"""check_no_secrets 的規則與整合測試（純標準庫 unittest）。

全部使用合成值。這些合成值出現在本檔的原始碼裡，因此也會被 --all-tracked 掃到 ——
它們在 .secretscan-allow.json 有精確字串＋理由的豁免，**不是整檔豁免**。

偵測規則本身一律在 `--no-allowlist` 模式下驗證，避免測試只是確認合成值都被略過。

直接跑：
    python3 tools/tests/test_check_no_secrets.py
"""
import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
SCANNER = REPO_ROOT / "tools" / "check_no_secrets.py"

_spec = importlib.util.spec_from_file_location("check_no_secrets", SCANNER)
scanner = importlib.util.module_from_spec(_spec)
sys.modules["check_no_secrets"] = scanner
_spec.loader.exec_module(scanner)

# ── 合成樣本（皆為假值；.invalid 是 RFC 2606 保留網域）───────────────────────
FAKE_FORM_URL = "https://docs.google.com/forms/d/e/1EXAMPLEexampleEXAMPLE00000/viewform"
FAKE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1EXAMPLEsheet0000000000000/edit"
FAKE_FOLDER_URL = "https://drive.google.com/drive/folders/1EXAMPLEfolder000000000000"
FAKE_CLI = "--form-id 1EXAMPLEformid00000000000"
FAKE_EMAIL = "tester@example.invalid"
FAKE_CRED = '{"refresh_token": "EXAMPLE-not-a-real-token"}'


def rules_hit(text: str) -> set[str]:
    """在不套用白名單的情況下，回傳阻擋級規則名稱。"""
    blocking, _ = scanner.scan_text("sample.txt", text, allowed=set())
    return {item.rsplit(":", 1)[1] for item in blocking}


def advisories(text: str) -> set[str]:
    _, advisory = scanner.scan_text("sample.txt", text, allowed=set())
    return {item.rsplit(":", 1)[1] for item in advisory}


class BlockingRuleTest(unittest.TestCase):
    def test_google_doc_url_blocked(self):
        self.assertIn("google-doc-url", rules_hit(FAKE_FORM_URL))
        self.assertIn("google-doc-url", rules_hit(FAKE_SHEET_URL))

    def test_google_folder_url_blocked(self):
        self.assertIn("google-folder-url", rules_hit(FAKE_FOLDER_URL))

    def test_bare_cli_id_blocked(self):
        self.assertIn("cli-id-literal", rules_hit(FAKE_CLI))

    def test_email_blocked(self):
        self.assertIn("email-address", rules_hit(f"聯絡 {FAKE_EMAIL} 詢問"))

    def test_credential_blob_blocked(self):
        self.assertIn("credential-blob", rules_hit(FAKE_CRED))
        self.assertIn("credential-blob", rules_hit('{"type": "authorized_user"}'))

    def test_config_value_blocked(self):
        self.assertIn("config-value", rules_hit('"chapter_name": "某某分會"'))


class NoFalsePositiveTest(unittest.TestCase):
    def test_placeholders_pass(self):
        text = (
            '"room_form_id": "<小房間預約表單的 Google 檔案 ID>"\n'
            '"chapter_name": "<分會全名>"\n'
            "$R options --form-id <FORM_ID> --question 預約日期\n"
        )
        self.assertEqual(rules_hit(text), set())

    def test_shell_variable_and_attribute_pass(self):
        text = '--form-id "$FORM_ID"\n--file-id ${FILE_ID}\nargs.form_id\nargs.file_id\n'
        self.assertEqual(rules_hit(text), set())

    def test_bare_credential_field_names_pass(self):
        """欄位名稱本身不是機密，程式與測試會正常用到。"""
        text = (
            "creds = FakeCreds(refresh_token='rt')\n"
            "from_authorized_user_file(path, scopes)\n"
            "if 'refresh_token' in payload:\n"
        )
        self.assertEqual(rules_hit(text), set())

    def test_long_symbol_and_mime_not_blocked(self):
        text = (
            "from_authorized_user_file\n"
            "application/vnd.openxmlformats-officedocument.presentationml.presentation\n"
        )
        self.assertEqual(rules_hit(text), set())

    def test_long_chinese_paragraph_not_blocked(self):
        """Python 的 \\w 包含中文，若用 [-\\w]{25,} 當阻擋規則會誤中整段中文。"""
        text = "這是一段很長的中文說明文字用來確認掃描器不會把中文段落誤判成敏感字串總共超過二十五個字"
        self.assertEqual(rules_hit(text), set())
        self.assertEqual(advisories(text), set())

    def test_non_google_host_not_blocked(self):
        """任意網站的 /d/ 或 id= 不是 Google 資源。"""
        text = (
            "https://example.invalid/files/d/1SOMETHINGthatlookslikeanid00\n"
            "https://cdn.example.invalid/x?id=1SOMETHINGthatlookslikeanid00\n"
        )
        self.assertEqual(rules_hit(text), set())


class AdvisoryTest(unittest.TestCase):
    def test_id_shaped_token_is_advisory_only(self):
        """像 Google ID 的字串（含數字與大寫）只提示，不阻擋。"""
        text = "token=1AbCdEfGhIjKlMnOpQrStUvWx"
        self.assertEqual(rules_hit(text), set())
        self.assertIn("long-token-advisory", advisories(text))

    def test_code_symbols_are_not_even_advised(self):
        """程式符號沒有數字或大寫，不該淹沒提示訊號。"""
        for text in (
            "from_authorized_user_file",
            "openxmlformats-officedocument",
            "這是一段很長的中文說明文字用來確認掃描器不會把中文段落誤判成敏感字串",
        ):
            self.assertEqual(advisories(text), set(), text)


class AllowlistTest(unittest.TestCase):
    def test_allowlisted_exact_string_is_skipped(self):
        allowed = {"https://docs.google.com/forms/d/e/1EXAMPLEexampleEXAMPLE00000"}
        blocking, _ = scanner.scan_text("t.txt", FAKE_FORM_URL, allowed)
        self.assertEqual(blocking, [])

    def test_same_string_blocked_without_allowlist(self):
        blocking, _ = scanner.scan_text("t.txt", FAKE_FORM_URL, allowed=set())
        self.assertTrue(blocking)

    def test_allowlist_entry_requires_reason(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad = Path(tmp) / ".secretscan-allow.json"
            bad.write_text('{"allow": [{"value": "x"}]}', encoding="utf-8")
            orig = scanner.ALLOWLIST_FILE
            scanner.ALLOWLIST_FILE = bad
            try:
                with self.assertRaises(scanner.ScanFailure):
                    scanner.load_allowlist(enabled=True)
            finally:
                scanner.ALLOWLIST_FILE = orig


class ForbiddenNameTest(unittest.TestCase):
    def test_forbidden_names(self):
        for name in ("config.json", "token.json", "credentials.json", ".env", ".DS_Store"):
            self.assertTrue(scanner.forbidden_name(f"some/dir/{name}"), name)
        self.assertTrue(scanner.forbidden_name("client_secret_123.json"))
        self.assertTrue(scanner.forbidden_name(".env.production"))

    def test_allowed_names(self):
        for name in ("config.example.json", "README.md", "tools/gforms.py"):
            self.assertFalse(scanner.forbidden_name(name), name)


class DecodeFailureTest(unittest.TestCase):
    def test_binary_raises_scan_failure(self):
        with self.assertRaises(scanner.ScanFailure):
            scanner.decode_or_fail("logo.png", b"\x89PNG\r\n\x1a\n\xff\xfe\x00\x01")


class GitIntegrationTest(unittest.TestCase):
    """在臨時 git repo 上驗證：掃的是 index 而不是工作樹。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        (Path(self.tmp) / "tools").mkdir()
        shutil.copy(SCANNER, Path(self.tmp) / "tools" / "check_no_secrets.py")
        self.git("init", "-q")
        self.git("config", "user.email", "t@example.invalid")
        self.git("config", "user.name", "T")

    def git(self, *args):
        subprocess.run(["git", *args], cwd=self.tmp, check=True, capture_output=True)

    def scan(self, *extra):
        return subprocess.run(
            [sys.executable, "tools/check_no_secrets.py", *extra],
            cwd=self.tmp, capture_output=True, text=True,
        )

    def write(self, relpath, text):
        path = Path(self.tmp) / relpath
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        return path

    def test_clean_first_commit_passes(self):
        """首次 commit（尚無 HEAD）也要能掃。"""
        self.write("notes.md", "使用 <FORM_ID> 佔位符，不放真實 ID。\n")
        self.git("add", "notes.md")
        result = self.scan()
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_staged_sensitive_but_clean_worktree_still_blocked(self):
        """staged 有問題、工作樹已清乾淨 → 仍要擋（這正是掃 index 的理由）。"""
        path = self.write("notes.md", f"連結：{FAKE_FORM_URL}\n")
        self.git("add", "notes.md")
        path.write_text("已經清乾淨了\n", encoding="utf-8")  # 只改工作樹，不重新 add
        result = self.scan("--no-allowlist")
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("google-doc-url", result.stderr)

    def test_path_with_spaces_and_rename(self):
        self.write("有 空白 的 檔名.md", "乾淨內容\n")
        self.git("add", "有 空白 的 檔名.md")
        self.git("commit", "-qm", "init")
        self.git("mv", "有 空白 的 檔名.md", "改名 後的 檔名.md")
        result = self.scan()
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_forbidden_file_staged_is_blocked(self):
        self.write("config.json", "{}\n")
        self.git("add", "-f", "config.json")
        result = self.scan()
        self.assertEqual(result.returncode, 1)
        self.assertIn("forbidden-filename", result.stderr)

    def test_binary_staged_exits_two(self):
        (Path(self.tmp) / "logo.png").write_bytes(b"\x89PNG\r\n\x1a\n\xff\xfe\x00\x01")
        self.git("add", "logo.png")
        result = self.scan()
        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("人工審查", result.stderr)

    def test_history_mode_catches_removed_secret(self):
        """秘密在舊 commit、後來被移除 → 歷史模式仍要抓到。"""
        self.write("notes.md", f"連結：{FAKE_FORM_URL}\n")
        self.git("add", "notes.md")
        self.git("commit", "-qm", "oops")
        self.write("notes.md", "已移除\n")
        self.git("add", "notes.md")
        self.git("commit", "-qm", "clean up")
        staged = self.scan()
        self.assertEqual(staged.returncode, 0, "目前版本是乾淨的")
        history = self.scan("--history", "HEAD", "--no-allowlist")
        self.assertEqual(history.returncode, 1, history.stdout + history.stderr)
        self.assertIn("google-doc-url", history.stderr)

    def test_history_mode_scans_commit_messages(self):
        self.write("notes.md", "乾淨\n")
        self.git("add", "notes.md")
        self.git("commit", "-qm", f"參考 {FAKE_SHEET_URL}")
        result = self.scan("--history", "HEAD", "--no-allowlist")
        self.assertEqual(result.returncode, 1)
        self.assertIn("commit-messages", result.stderr)


class RealRepoTest(unittest.TestCase):
    """整合驗收：去識別後的 repo 本身必須通過掃描。"""

    def test_whole_repo_passes(self):
        result = subprocess.run(
            [sys.executable, "tools/check_no_secrets.py", "--all-tracked"],
            cwd=REPO_ROOT, capture_output=True, text=True,
        )
        if result.returncode == 2 and "not a git repository" in result.stderr.lower():
            self.skipTest("尚未 git init，略過全 repo 掃描")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
