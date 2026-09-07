#!/usr/bin/env python3
"""用 Google Drive API 就地覆蓋既有 Drive 檔的內容（fileId／分享連結不變）。

Claude 的 Google Drive 連接器只能建新檔或改中繼資料，無法覆蓋既有檔內容，
故自建此腳本。token 全域快取在 ~/.config/gdrive_api/，任何資料夾可跑。

前置（一次性，使用者自己做）：
  1. Google Cloud Console 建專案 → 啟用 Google Drive API
  2. OAuth 同意畫面：User Type = External、把自己加為 Test user
  3. 建 OAuth 用戶端 ID（類型：桌面應用程式）→ 下載 JSON
  4. 把該 JSON 改名 credentials.json 放到 ~/.config/gdrive_api/
第一次跑會自動開瀏覽器要你登入授權，token 快取到 ~/.config/gdrive_api/token.json，之後免登入。

用法：
    R="uv run --with google-api-python-client --with google-auth-oauthlib --with google-auth python tools/gdrive_update.py"

    # 看目前授權的是哪個帳號（順便完成首次登入）
    $R whoami

    # 把 Drive 檔下載到本機（當範本用；不佔 Claude context）
    $R download --file-id <FILE_ID> --out ./scratch/template.pptx

    # 用本機檔覆蓋既有 Drive 檔（連結不變）
    $R update --file-id <FILE_ID> --local ./scratch/來賓名單投影片_<日期場次>.pptx

    # 同上但保留這次為永久版本（Drive「管理版本」看得到、不會被自動清）
    $R update --file-id <FILE_ID> --local <path> --keep-revision

  <FILE_ID> 從私有設定 config.json 取得（如 guest_slides_file_id），不要寫死在文件或程式裡。

注意：
  - 只覆蓋「內容」，不改檔名、不改父資料夾、不改權限，連結與 fileId 完全不變。
  - 需要對該檔有 writer 權限。
  - ⚠️ whoami 會印出授權帳號的姓名與 email。不要貼進 issue、log 或截圖。
  - credentials.json / token.json 都是機密，請勿外流或 commit。
"""
import argparse
import os
import sys
from pathlib import Path

CONFIG_DIR = Path.home() / ".config" / "gdrive_api"
CREDENTIALS_FILE = CONFIG_DIR / "credentials.json"
TOKEN_FILE = CONFIG_DIR / "token.json"
# 覆蓋既有任意檔需要 full drive scope（drive.file 只能碰自己建的檔）
SCOPES = ["https://www.googleapis.com/auth/drive"]

# 常見副檔名 → MIME（Drive 更新內容時建議一併指定，避免類型跑掉）
MIME_BY_EXT = {
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pdf": "application/pdf",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
}


def _is_invalid_grant(exc):
    """判斷 RefreshError 是否為 refresh_token 被撤銷（invalid_grant）。

    google-auth 會把 OAuth 錯誤塞進例外的其中一個 dict 引數，優先讀結構化
    欄位；取不到才退回字串比對。只有 invalid_grant 才代表憑證失效需重新授權，
    invalid_client／invalid_scope 等設定錯誤與暫時性伺服器錯誤都不算。
    """
    for arg in exc.args:
        if isinstance(arg, dict) and arg.get("error") == "invalid_grant":
            return True
    return "invalid_grant" in str(exc)


def get_service():
    from google.auth.exceptions import RefreshError
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

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
                # 暫時性錯誤（如 token endpoint 500）照拋，維持 fail fast；
                # 只有 refresh_token 被撤銷（Testing app 約每 7 天到期）才作廢重授權
                if not _is_invalid_grant(exc):
                    raise
                print("Drive 授權已失效，開瀏覽器重新登入…", file=sys.stderr)
                creds = None
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
        os.chmod(TOKEN_FILE, 0o600)
    return build("drive", "v3", credentials=creds)


def cmd_whoami(_args):
    service = get_service()
    about = service.about().get(fields="user(displayName,emailAddress)").execute()
    u = about["user"]
    print(f"已授權帳號：{u.get('displayName')} <{u.get('emailAddress')}>")


def cmd_download(args):
    import io
    from googleapiclient.http import MediaIoBaseDownload

    service = get_service()
    meta = service.files().get(fileId=args.file_id, fields="id,name,mimeType").execute()
    out = Path(args.out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    req = service.files().get_media(fileId=args.file_id)
    with io.FileIO(str(out), "wb") as fh:
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    print(f"✅ 已下載 {meta['name']} → {out}  ({out.stat().st_size} bytes)")


def cmd_update(args):
    from googleapiclient.http import MediaFileUpload

    local = Path(args.local).expanduser()
    if not local.exists():
        sys.exit(f"本機檔不存在：{local}")

    service = get_service()

    # 先確認目標檔存在、印出更新前資訊
    before = service.files().get(
        fileId=args.file_id, fields="id,name,modifiedTime,version,mimeType"
    ).execute()
    print(f"目標檔：{before['name']}  (id={before['id']})")
    print(f"更新前 version={before.get('version')}  modifiedTime={before.get('modifiedTime')}")

    mime = MIME_BY_EXT.get(local.suffix.lower())  # None → 讓 Drive 自行判斷
    media = MediaFileUpload(str(local), mimetype=mime, resumable=True)
    updated = service.files().update(
        fileId=args.file_id,
        media_body=media,
        keepRevisionForever=args.keep_revision,
        fields="id,name,modifiedTime,version",
    ).execute()

    print("✅ 已覆蓋內容（連結／fileId 不變）")
    print(f"更新後 version={updated.get('version')}  modifiedTime={updated.get('modifiedTime')}")
    print(f"開檔：https://drive.google.com/file/d/{updated['id']}/view")


def main():
    p = argparse.ArgumentParser(description="覆蓋既有 Google Drive 檔內容（連結不變）")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("whoami", help="顯示目前授權帳號（順便完成首次登入）").set_defaults(func=cmd_whoami)

    dl = sub.add_parser("download", help="把 Drive 檔下載到本機")
    dl.add_argument("--file-id", required=True, help="要下載的 Drive fileId")
    dl.add_argument("--out", required=True, help="本機輸出路徑")
    dl.set_defaults(func=cmd_download)

    up = sub.add_parser("update", help="用本機檔覆蓋既有 Drive 檔")
    up.add_argument("--file-id", required=True, help="要覆蓋的 Drive fileId")
    up.add_argument("--local", required=True, help="本機檔路徑")
    up.add_argument("--keep-revision", action="store_true", help="保留此次為永久版本")
    up.set_defaults(func=cmd_update)

    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
