"""bni_config 的設定邊界測試（純標準庫 unittest）。

全部使用**合成設定**寫在臨時目錄，不讀真實的 config.json。

直接跑：
    python3 tools/tests/test_bni_config.py
"""
import importlib.util
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

_TOOL = Path(__file__).resolve().parent.parent / "bni_config.py"
_spec = importlib.util.spec_from_file_location("bni_config", _TOOL)
bni_config = importlib.util.module_from_spec(_spec)
sys.modules["bni_config"] = bni_config
_spec.loader.exec_module(bni_config)

# 合成設定。刻意用 (鍵, 值) 序對組成，不寫成 `"欄位": "值"` 的字面形狀 ——
# 那個形狀正是 check_no_secrets 的 config-value 阻擋規則要抓的東西，
# 測試檔沒有理由長得像一份設定檔。
_SYNTHETIC_FIELDS = [
    ("chapter_name", "示範分會"),
    ("chapter_short", "示範"),
    ("timezone", "Asia/Taipei"),
    ("room_form_id", "SYNTHETIC-form"),
    ("room_form_responder_url", "https://example.invalid/form"),
    ("room_sheet_id", "SYNTHETIC-sheet"),
    ("guest_slides_file_id", "SYNTHETIC-slides"),
    ("guest_slides_parent_folder_id", "SYNTHETIC-folder"),
    ("signin_sheet_file_id", "SYNTHETIC-signin"),
    ("member_roster_doc_id", "SYNTHETIC-roster-doc"),
    ("member_roster_sheet_id", "SYNTHETIC-roster-sheet"),
    ("slides25s_ppt_folder_id", "SYNTHETIC-25s-folder"),
    ("slides25s_upload_folder_id", "SYNTHETIC-25s-upload"),
    ("dismember_folder_id", "SYNTHETIC-dismember"),
]
FULL = {"schema_version": 1, **dict(_SYNTHETIC_FIELDS)}


class ValidateTest(unittest.TestCase):
    def test_all_flows_accept_full_config(self):
        for flow in ("bni-local", "bni-cloud", "bnippt", "bnimembers", "bni25s", "assemble"):
            self.assertEqual(bni_config.validate(dict(FULL), flow), FULL, flow)

    def test_minimal_config_per_flow(self):
        """每個流程只要它自己的必要欄位就夠，不強迫填滿無關欄位。"""
        for flow, fields in bni_config.REQUIRED_BY_FLOW.items():
            minimal = {k: FULL[k] for k in fields}
            self.assertEqual(bni_config.validate(minimal, flow), minimal, flow)

    def test_cloud_does_not_require_form_id_or_slides(self):
        fields = bni_config.REQUIRED_BY_FLOW["bni-cloud"]
        self.assertNotIn("room_form_id", fields)
        self.assertFalse([f for f in fields if f.startswith("guest_slides")])

    def test_missing_field_names_the_field(self):
        cfg = dict(FULL)
        del cfg["room_sheet_id"]
        with self.assertRaises(bni_config.ConfigError) as ctx:
            bni_config.validate(cfg, "bni-local")
        self.assertIn("room_sheet_id", str(ctx.exception))

    def test_placeholder_is_rejected(self):
        cfg = dict(FULL, room_sheet_id="<預約彙整試算表的 Google 檔案 ID>")
        with self.assertRaises(bni_config.ConfigError) as ctx:
            bni_config.validate(cfg, "bni-local")
        self.assertIn("佔位符", str(ctx.exception))
        self.assertIn("room_sheet_id", str(ctx.exception))

    def test_blank_is_rejected(self):
        cfg = dict(FULL, chapter_short="   ")
        with self.assertRaises(bni_config.ConfigError) as ctx:
            bni_config.validate(cfg, "bni-local")
        self.assertIn("空白", str(ctx.exception))

    def test_wrong_type_is_rejected(self):
        cfg = dict(FULL, chapter_short=123)
        with self.assertRaises(bni_config.ConfigError):
            bni_config.validate(cfg, "bni-local")

    def test_missing_schema_version(self):
        cfg = dict(FULL)
        del cfg["schema_version"]
        with self.assertRaises(bni_config.ConfigError) as ctx:
            bni_config.validate(cfg, "bni-local")
        self.assertIn("schema_version", str(ctx.exception))

    def test_bool_schema_version_rejected(self):
        """True 在 Python 是 int 的子類，不能被當成版本 1。"""
        with self.assertRaises(bni_config.ConfigError):
            bni_config.validate(dict(FULL, schema_version=True), "bni-local")

    def test_unsupported_schema_version(self):
        with self.assertRaises(bni_config.ConfigError) as ctx:
            bni_config.validate(dict(FULL, schema_version=99), "bni-local")
        self.assertIn("99", str(ctx.exception))

    def test_invalid_timezone(self):
        with self.assertRaises(bni_config.ConfigError) as ctx:
            bni_config.validate(dict(FULL, timezone="Mars/Olympus"), "bni-local")
        self.assertIn("timezone", str(ctx.exception))

    def test_unknown_flow(self):
        with self.assertRaises(bni_config.ConfigError):
            bni_config.validate(dict(FULL), "nope")


class ResolvePathTest(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.root, ignore_errors=True)

    def write(self, relpath, data=FULL):
        path = self.root / relpath
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return path

    def test_repo_config_used_when_no_explicit(self):
        self.write("config.json")
        self.assertEqual(
            bni_config.resolve_config_path(None, self.root), self.root / "config.json"
        )

    def test_explicit_used_when_no_repo_config(self):
        other = self.write("elsewhere/cloud.json")
        self.assertEqual(
            bni_config.resolve_config_path(str(other), self.root), other.resolve()
        )

    def test_two_sources_conflict(self):
        """兩來源互斥：不自行挑一個。"""
        self.write("config.json")
        other = self.write("elsewhere/cloud.json")
        with self.assertRaises(bni_config.ConfigError) as ctx:
            bni_config.resolve_config_path(str(other), self.root)
        self.assertIn("兩個設定來源", str(ctx.exception))

    def test_same_path_is_not_conflict(self):
        repo = self.write("config.json")
        self.assertEqual(
            bni_config.resolve_config_path(str(repo), self.root), repo.resolve()
        )

    def test_missing_everything_explains_how_to_fix(self):
        with self.assertRaises(bni_config.ConfigError) as ctx:
            bni_config.resolve_config_path(None, self.root)
        self.assertIn("config.example.json", str(ctx.exception))

    def test_explicit_path_not_found(self):
        with self.assertRaises(bni_config.ConfigError):
            bni_config.resolve_config_path(str(self.root / "nope.json"), self.root)

    def test_load_rejects_invalid_json(self):
        (self.root / "config.json").write_text("{not json", encoding="utf-8")
        with self.assertRaises(bni_config.ConfigError):
            bni_config.load("bni-local", None, self.root)

    def test_load_rejects_non_object(self):
        (self.root / "config.json").write_text("[1, 2]", encoding="utf-8")
        with self.assertRaises(bni_config.ConfigError):
            bni_config.load("bni-local", None, self.root)

    def test_load_happy_path(self):
        self.write("config.json")
        self.assertEqual(bni_config.load("bnippt", None, self.root)["schema_version"], 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
