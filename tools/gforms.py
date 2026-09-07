#!/usr/bin/env python3
"""用 Google Forms API 讀取表單結構與回覆（唯讀）。

Claude 的 Google Drive 連接器讀不到 Google 表單的題目與選項（表單內容不是 Drive
檔案內容），故自建此腳本。憑證與 token 放在 ~/.config/gdrive_api/，任何資料夾可跑。

前置（一次性，使用者自己做）：
  1. Google Cloud Console 建專案 → OAuth 同意畫面 User Type = External、把自己加為 Test user
  2. 建 OAuth 用戶端 ID（類型：桌面應用程式）→ 下載 JSON
  3. 改名 credentials.json 放到 ~/.config/gdrive_api/
  4. 在同一個 Google Cloud 專案啟用 Google Forms API
     https://console.developers.google.com/apis/api/forms.googleapis.com/overview
  scope 不用另外申請：Forms API 的 forms.get／forms.responses.list 都接受
  auth/drive scope（官方 discovery doc 確認），與 gdrive_update.py 共用同一份 token。

用法：
    R="uv run --with google-auth --with google-auth-oauthlib python tools/gforms.py"

    # 看整份表單：標題、每一題、選擇題的所有選項
    $R get --form-id <FORM_ID>

    # 只要某一題的選項，一行一個（給 skill 解析用）
    $R options --form-id <FORM_ID> --question 預約日期

    # 列出回覆（最新在前）
    $R responses --form-id <FORM_ID> --limit 20

  <FORM_ID> 從私有設定 config.json 的 room_form_id 取得，不要寫死在文件或程式裡。

注意：
  - 本腳本**唯讀**：只呼叫 forms.get 與 forms.responses.list，不會改動表單。
    若表單掛有自動維護選項的 Apps Script，更不要用 API 去寫，會打架。
  - 需要對該表單有讀取權限。
  - ⚠️ get／responses 的輸出含會員姓名、email、表單連結。不要貼進 issue、log
    或截圖，也不要把輸出存成 fixture 進版控。
  - credentials.json / token.json 都是機密，請勿外流或 commit。
"""
import argparse
import json
import os
import sys
from pathlib import Path

CONFIG_DIR = Path.home() / ".config" / "gdrive_api"
CREDENTIALS_FILE = CONFIG_DIR / "credentials.json"
TOKEN_FILE = CONFIG_DIR / "token.json"
# forms.get／forms.responses.list 都吃 drive scope，與 gdrive_update.py 共用 token
SCOPES = ["https://www.googleapis.com/auth/drive"]

API_ROOT = "https://forms.googleapis.com/v1"
ENABLE_URL = "https://console.developers.google.com/apis/api/forms.googleapis.com/overview"


def _is_invalid_grant(exc):
    """判斷 RefreshError 是否為 refresh_token 被撤銷（invalid_grant）。

    與 gdrive_update.py 同一套判斷：只有 invalid_grant 代表憑證失效需重新授權，
    invalid_client／invalid_scope 等設定錯誤與暫時性伺服器錯誤都不算。
    """
    for arg in exc.args:
        if isinstance(arg, dict) and arg.get("error") == "invalid_grant":
            return True
    return "invalid_grant" in str(exc)


def get_session():
    from google.auth.exceptions import RefreshError
    from google.auth.transport.requests import AuthorizedSession, Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    if not CREDENTIALS_FILE.exists():
        sys.exit(
            f"找不到 {CREDENTIALS_FILE}\n"
            "請先到 Google Cloud Console 建立 OAuth 桌面應用程式憑證，"
            "下載後改名 credentials.json 放到 ~/.config/gdrive_api/。"
        )

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError as exc:
                # 暫時性錯誤照拋，維持 fail fast；只有 refresh_token 被撤銷才重新授權
                if not _is_invalid_grant(exc):
                    raise
                print("Google 授權已失效，開瀏覽器重新登入…", file=sys.stderr)
                creds = None
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
        os.chmod(TOKEN_FILE, 0o600)
    return AuthorizedSession(creds)


def api_get(session, path, **params):
    """呼叫 Forms API，非 200 一律 fail fast 並印出可行動的錯誤訊息。"""
    resp = session.get(f"{API_ROOT}/{path}", params=params or None)
    if resp.status_code == 200:
        return resp.json()

    err = resp.json().get("error", {})
    message = err.get("message", resp.text)
    if "has not been used in project" in message or "is disabled" in message:
        sys.exit(
            f"Google Forms API 尚未啟用（HTTP {resp.status_code}）。\n"
            f"請到 {ENABLE_URL} 按「啟用」，等幾分鐘生效後重試。\n"
            f"原始訊息：{message}"
        )
    sys.exit(f"Forms API 失敗 HTTP {resp.status_code}：{err.get('status')} | {message}")


def iter_questions(form):
    """走訪表單所有題目，yield (title, question dict)。

    表單的 item 有多種型別（questionItem／questionGroupItem／pageBreakItem／
    textItem…），只有前兩種帶題目。
    """
    for item in form.get("items", []):
        title = item.get("title", "")
        if "questionItem" in item:
            yield title, item["questionItem"]["question"]
        for sub in item.get("questionGroupItem", {}).get("questions", []):
            yield title, sub


def choice_values(question):
    """選擇題 → 選項字串清單；非選擇題 → None。"""
    choice = question.get("choiceQuestion")
    if not choice:
        return None
    return [opt.get("value", "(其他)") for opt in choice.get("options", [])]


def find_question(form, keyword):
    """依題目標題關鍵字找唯一一題，找不到或找到多題都 fail fast。"""
    matches = [(t, q) for t, q in iter_questions(form) if keyword in t]
    if not matches:
        titles = [t for t, _ in iter_questions(form)]
        sys.exit(f"找不到題目含「{keyword}」。表單題目有：{titles}")
    if len(matches) > 1:
        sys.exit(f"題目含「{keyword}」有 {len(matches)} 題，請給更精確的關鍵字：{[t for t, _ in matches]}")
    return matches[0]


def cmd_get(args):
    form = api_get(get_session(), f"forms/{args.form_id}")
    if args.json:
        print(json.dumps(form, ensure_ascii=False, indent=2))
        return
    info = form.get("info", {})
    print(f"表單：{info.get('title')}")
    if info.get("documentTitle") and info["documentTitle"] != info.get("title"):
        print(f"檔名：{info['documentTitle']}")
    print(f"連結：{form.get('responderUri')}")
    print()
    for title, question in iter_questions(form):
        values = choice_values(question)
        required = "＊" if question.get("required") else "　"
        if values is None:
            print(f"{required}{title}（非選擇題）")
        else:
            print(f"{required}{title}（{len(values)} 個選項）")
            for v in values:
                print(f"    - {v}")


def cmd_options(args):
    form = api_get(get_session(), f"forms/{args.form_id}")
    title, question = find_question(form, args.question)
    values = choice_values(question)
    if values is None:
        sys.exit(f"「{title}」不是選擇題，沒有選項。")
    for v in values:
        print(v)


def cmd_responses(args):
    session = get_session()
    form = api_get(session, f"forms/{args.form_id}")
    label_by_id = {}
    for title, question in iter_questions(form):
        label_by_id[question["questionId"]] = title

    # ⚠️ Forms API 回傳的回覆「沒有排序」（實測時間亂序），
    # 所以不能只抓一頁再排——必須全部抓回來排完再截斷，否則 --limit 會拿到隨機子集。
    responses = []
    page_token = None
    while True:
        params = {"pageSize": 5000}
        if page_token:
            params["pageToken"] = page_token
        data = api_get(session, f"forms/{args.form_id}/responses", **params)
        responses.extend(data.get("responses", []))
        page_token = data.get("nextPageToken")
        if not page_token:
            break

    if not responses:
        print("（沒有回覆）")
        return

    responses.sort(key=lambda r: r.get("lastSubmittedTime", ""), reverse=True)
    total = len(responses)
    for r in responses[: args.limit]:
        print(f"— {r.get('lastSubmittedTime')}  (id={r.get('responseId')})")
        for qid, ans in r.get("answers", {}).items():
            values = [a.get("value") for a in ans.get("textAnswers", {}).get("answers", [])]
            # 表單上已刪除的題目仍會留在舊回覆裡，查不到標題就標示出來
            label = label_by_id.get(qid, f"(已刪除的題目 {qid})")
            print(f"    {label}：{'、'.join(values)}")
    if total > args.limit:
        print(f"\n（共 {total} 筆，只顯示最新 {args.limit} 筆）")


def main():
    p = argparse.ArgumentParser(description="讀取 Google 表單的題目、選項與回覆（唯讀）")
    sub = p.add_subparsers(dest="cmd", required=True)

    g = sub.add_parser("get", help="看整份表單的題目與選項")
    g.add_argument("--form-id", required=True, help="表單的 Google 檔案 ID")
    g.add_argument("--json", action="store_true", help="輸出原始 JSON")
    g.set_defaults(func=cmd_get)

    o = sub.add_parser("options", help="只印某一題的選項，一行一個")
    o.add_argument("--form-id", required=True, help="表單的 Google 檔案 ID")
    o.add_argument("--question", required=True, help="題目標題關鍵字（需唯一命中）")
    o.set_defaults(func=cmd_options)

    r = sub.add_parser("responses", help="列出回覆（最新在前）")
    r.add_argument("--form-id", required=True, help="表單的 Google 檔案 ID")
    r.add_argument("--limit", type=int, default=20, help="最多取幾筆（預設 20）")
    r.set_defaults(func=cmd_responses)

    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
