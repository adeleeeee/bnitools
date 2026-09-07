#!/usr/bin/env python3
"""公開 repo 的敏感內容掃描（純標準庫）。

這個 repo 是公開的，但實際的分會名稱與 Google 檔案 ID 只能留在未進版控的
config.json。本工具在 commit 前擋下不小心貼進來的真實值。

設計原則：
  - **掃 Git index 的 staged blob，不掃工作樹**。工作樹乾淨但 staged 有問題也要擋。
  - **只用高信心規則阻擋**。單靠「長字串」缺乏辨識力（Python 的 \\w 還包含中文），
    一般長 token 只列為人工檢查提示，不阻擋。
  - **欄位名稱本身不是機密**。client_secret／refresh_token 這些符號在程式與測試裡
    正常出現；要擋的是帶有疑似實際值的憑證結構。
  - **輸出只印 檔名:行號:規則名，不回顯命中內容**，避免掃描報告本身變成外洩管道。
  - 本檔**不得包含任何真實 ID 清單**。需要比對已知私有值時，另行在本機讀私有設定
    並遮罩結果，不寫進公開 scanner。

用法：
    python3 tools/check_no_secrets.py                 # 掃 staged（pre-commit 用）
    python3 tools/check_no_secrets.py --all-tracked   # 掃目前 index 全部檔案
    python3 tools/check_no_secrets.py --history HEAD  # 掃該 ref 可達的歷史內容
    python3 tools/check_no_secrets.py --paths a.md b.py   # 掃指定檔案（工作樹）
    python3 tools/check_no_secrets.py --no-allowlist  # 停用白名單（測試偵測規則用）

退出碼：0＝乾淨；1＝有阻擋級發現；2＝操作失敗（git 失敗、無法解碼等需人工審查）。

⚠️ scanner 通過**不等於**證明乾淨。首次公開發布前仍須逐檔人工確認，並確認
將發布的 refs 範圍與作者 email 是否為有意公開的身分。
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
ALLOWLIST_FILE = REPO_ROOT / ".secretscan-allow.json"

# Google 檔案 ID 的形狀。刻意不含中文與其他 Unicode：只認 ASCII。
_ID = r"[A-Za-z0-9_-]{20,}"
# 只認 Google 網域，不把任意網站的 /d/ 或 id= 當成 Google 資源
_GOOGLE_HOST = r"(?:docs|drive|forms|sheets|script)\.google\.com"

# 設定欄位一旦出現實際值就要擋；佔位符 <...> 放行
_CONFIG_FIELDS = (
    "chapter_name|chapter_short|room_form_id|room_form_responder_url|"
    "room_sheet_id|guest_slides_file_id|guest_slides_parent_folder_id"
)

BLOCKING_RULES = [
    (
        "google-doc-url",
        re.compile(rf"https?://{_GOOGLE_HOST}/[^\s\"'<>)]*?/d/(?:e/)?{_ID}"),
    ),
    (
        "google-folder-url",
        re.compile(rf"https?://{_GOOGLE_HOST}/[^\s\"'<>)]*?/folders/{_ID}"),
    ),
    (
        "google-id-param",
        re.compile(rf"https?://{_GOOGLE_HOST}/[^\s\"'<>)]*?[?&]id={_ID}"),
    ),
    (
        # 設定欄位帶實際值（佔位符 "<...>" 不會命中，因為 < > 被排除在值之外）
        "config-value",
        re.compile(rf"\"(?:{_CONFIG_FIELDS})\"\s*:\s*\"(?![<\s])[^\"<>]{{3,}}\""),
    ),
    (
        # CLI 參數位置上的實際 ID。--form-id <FORM_ID>、"$FORM_ID"、args.form_id 都不會命中
        "cli-id-literal",
        re.compile(rf"--(?:form|file)-id[=\s]+\"?{_ID}"),
    ),
    (
        "email-address",
        re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"),
    ),
    (
        # 憑證結構「帶疑似實際值」才擋；裸欄位名不擋
        "credential-blob",
        re.compile(
            r"\"(?:client_secret|refresh_token|private_key)\"\s*:\s*\"[^\"]{8,}\""
            r"|\"type\"\s*:\s*\"authorized_user\""
        ),
    ),
]

# 只提示、不阻擋：單靠長度無法辨識，誤判成本高於漏判（有發布前人工確認兜底）。
# 進一步要求同時含數字與大寫字母，把 from_authorized_user_file、
# openxmlformats-officedocument 這類程式符號濾掉，提示才有訊號值。
ADVISORY_RULES = [
    (
        "long-token-advisory",
        re.compile(
            rf"(?<![A-Za-z0-9_-])(?=[A-Za-z0-9_-]*\d)(?=[A-Za-z0-9_-]*[A-Z])"
            rf"{_ID}(?![A-Za-z0-9_-])"
        ),
    ),
]

FORBIDDEN_NAMES = {
    "config.json", "token.json", "credentials.json", ".env", ".DS_Store",
}
FORBIDDEN_PATTERNS = [re.compile(r"^client_secret.*\.json$"), re.compile(r"^\.env\..+$")]


class ScanFailure(Exception):
    """操作失敗（git 失敗、無法解碼）。一律非零退出並要求人工審查。"""


def run_git(args: list[str], binary: bool = False):
    proc = subprocess.run(
        ["git", *args], cwd=REPO_ROOT, capture_output=True,
    )
    if proc.returncode != 0:
        raise ScanFailure(
            f"git {' '.join(args[:2])} 失敗（returncode={proc.returncode}）："
            f"{proc.stderr.decode('utf-8', 'replace').strip()}"
        )
    return proc.stdout if binary else proc.stdout.decode("utf-8", "replace")


def load_allowlist(enabled: bool) -> set[str]:
    """白名單：精確字串＋理由。禁止整檔豁免，也不得寫入任何真實 ID。"""
    if not enabled or not ALLOWLIST_FILE.exists():
        return set()
    data = json.loads(ALLOWLIST_FILE.read_text(encoding="utf-8"))
    allowed = set()
    for entry in data["allow"]:
        if not entry.get("reason"):
            raise ScanFailure(f"白名單項目缺少 reason：{entry.get('value')!r}")
        allowed.add(entry["value"])
    return allowed


def forbidden_name(path: str) -> bool:
    name = Path(path).name
    if name in FORBIDDEN_NAMES:
        return True
    return any(p.match(name) for p in FORBIDDEN_PATTERNS)


def scan_text(path: str, text: str, allowed: set[str]) -> tuple[list[str], list[str]]:
    """回傳 (阻擋發現, 提示發現)，每筆格式 path:line:rule。不回顯命中內容。"""
    blocking, advisory = [], []
    for lineno, line in enumerate(text.splitlines(), 1):
        for name, pattern in BLOCKING_RULES:
            for match in pattern.finditer(line):
                if match.group(0) in allowed:
                    continue
                blocking.append(f"{path}:{lineno}:{name}")
        for name, pattern in ADVISORY_RULES:
            for match in pattern.finditer(line):
                if match.group(0) in allowed:
                    continue
                advisory.append(f"{path}:{lineno}:{name}")
    return blocking, advisory


def decode_or_fail(path: str, blob: bytes) -> str:
    try:
        return blob.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ScanFailure(
            f"{path}：無法以 UTF-8 解碼（二進位或其他編碼），本工具不會默默跳過。"
            f"請人工審查此檔是否含敏感內容後再決定。（{exc.reason}）"
        ) from exc


def staged_paths() -> list[str]:
    """新增／修改／複製／改名的 staged 路徑。-z 處理含空白與非 ASCII 的路徑。"""
    out = run_git(["diff", "--cached", "-z", "--name-only", "--diff-filter=ACMR"])
    return [p for p in out.split("\0") if p]


def index_paths() -> list[str]:
    out = run_git(["ls-files", "-z"])
    return [p for p in out.split("\0") if p]


def read_index_blob(path: str) -> bytes:
    """讀 index 裡的內容，不是工作樹。首次 commit（尚無 HEAD）同樣適用。"""
    return run_git(["show", f":{path}"], binary=True)


def scan_index(paths: list[str], allowed: set[str]) -> tuple[list[str], list[str]]:
    blocking, advisory = [], []
    for path in paths:
        if forbidden_name(path):
            blocking.append(f"{path}:0:forbidden-filename")
            continue
        text = decode_or_fail(path, read_index_blob(path))
        b, a = scan_text(path, text, allowed)
        blocking.extend(b)
        advisory.extend(a)
    return blocking, advisory


def scan_worktree(paths: list[str], allowed: set[str]) -> tuple[list[str], list[str]]:
    blocking, advisory = [], []
    for path in paths:
        if forbidden_name(path):
            blocking.append(f"{path}:0:forbidden-filename")
            continue
        text = decode_or_fail(path, Path(path).read_bytes())
        b, a = scan_text(path, text, allowed)
        blocking.extend(b)
        advisory.extend(a)
    return blocking, advisory


def scan_history(ref: str, allowed: set[str]) -> tuple[list[str], list[str]]:
    """掃該 ref 可達的所有 blob 與 commit 訊息。

    刻意要求明確指定 ref，不用 --all／--mirror 順手帶出其他不打算發布的 refs。
    """
    blocking, advisory = [], []

    listing = run_git(["rev-list", "--objects", ref])
    entries = []
    for line in listing.splitlines():
        sha, _, path = line.partition(" ")
        if path:
            entries.append((sha, path))

    if entries:
        check_input = "\n".join(sha for sha, _ in entries) + "\n"
        proc = subprocess.run(
            ["git", "cat-file", "--batch-check=%(objectname) %(objecttype)"],
            cwd=REPO_ROOT, input=check_input.encode(), capture_output=True,
        )
        if proc.returncode != 0:
            raise ScanFailure("git cat-file --batch-check 失敗，無法判斷物件型別")
        kinds = {}
        for line in proc.stdout.decode().splitlines():
            sha, _, kind = line.partition(" ")
            kinds[sha] = kind.strip()

        seen_blobs = set()
        for sha, path in entries:
            if kinds.get(sha) != "blob" or sha in seen_blobs:
                continue
            seen_blobs.add(sha)
            if forbidden_name(path):
                blocking.append(f"{ref}~{path}:0:forbidden-filename-in-history")
                continue
            text = decode_or_fail(f"{ref}~{path}", run_git(["cat-file", "-p", sha], binary=True))
            b, a = scan_text(f"{ref}~{path}", text, allowed)
            blocking.extend(b)
            advisory.extend(a)

    messages = run_git(["log", "--format=%H%n%B", ref])
    b, a = scan_text(f"{ref}~commit-messages", messages, allowed)
    blocking.extend(b)
    advisory.extend(a)

    tags = run_git(["for-each-ref", "--format=%(refname) %(contents)", "refs/tags"])
    if tags.strip():
        b, a = scan_text("refs/tags", tags, allowed)
        blocking.extend(b)
        advisory.extend(a)

    return blocking, advisory


def main():
    p = argparse.ArgumentParser(description="公開 repo 的敏感內容掃描")
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--all-tracked", action="store_true", help="掃 index 中所有已追蹤檔案")
    mode.add_argument("--history", metavar="REF", help="掃該 ref 可達的歷史內容與 commit 訊息")
    mode.add_argument("--paths", nargs="+", help="掃指定檔案（讀工作樹，測試用）")
    p.add_argument("--no-allowlist", action="store_true", help="停用白名單（驗證偵測規則用）")
    args = p.parse_args()

    try:
        allowed = load_allowlist(enabled=not args.no_allowlist)
        if args.history:
            blocking, advisory = scan_history(args.history, allowed)
            scope = f"歷史 {args.history}"
        elif args.all_tracked:
            blocking, advisory = scan_index(index_paths(), allowed)
            scope = "index 全部已追蹤檔案"
        elif args.paths:
            blocking, advisory = scan_worktree(args.paths, allowed)
            scope = f"{len(args.paths)} 個指定檔案"
        else:
            paths = staged_paths()
            if not paths:
                print("沒有 staged 檔案，跳過掃描。")
                return 0
            blocking, advisory = scan_index(paths, allowed)
            scope = f"{len(paths)} 個 staged 檔案"
    except ScanFailure as exc:
        print(f"❌ 掃描失敗，需人工審查：{exc}", file=sys.stderr)
        return 2

    if advisory:
        print(f"⚠️  人工檢查提示（不阻擋，共 {len(advisory)} 處）：", file=sys.stderr)
        for item in advisory[:20]:
            print(f"    {item}", file=sys.stderr)
        if len(advisory) > 20:
            print(f"    …另有 {len(advisory) - 20} 處", file=sys.stderr)

    if blocking:
        print(f"\n🔴 偵測到不該進公開 repo 的內容（{scope}）：", file=sys.stderr)
        for item in blocking:
            print(f"    {item}", file=sys.stderr)
        print(
            "\n真實的分會名稱與 Google 檔案 ID 只能放未進版控的 config.json。\n"
            "文件與程式請改用 <FORM_ID> 這類佔位符，或由 skill 讀設定後傳參數。\n"
            "確定是誤判時，在 .secretscan-allow.json 加入精確字串與理由（不得整檔豁免）。",
            file=sys.stderr,
        )
        return 1

    print(f"✅ 掃描通過（{scope}）。注意：通過不等於證明乾淨，首次發布前仍須人工逐檔確認。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
