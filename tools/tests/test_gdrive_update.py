"""gdrive_update.get_service() 的 OAuth 自癒回歸測試（純標準庫 unittest）。

直接跑（不需 pytest）：
    uv run --with google-auth python tools/tests/test_gdrive_update.py

四條路徑全部以假物件替換外部 API，不碰真網路、不彈瀏覽器：
  1. refresh 成功            → 不啟動 OAuth、正常寫回 token
  2. invalid_grant（撤銷）    → 啟動重新 OAuth、以新 credentials 覆寫 token
  3. 其他 RefreshError（暫時性）→ 原樣拋出、不 OAuth、不寫 token（fail fast）
  4. 無 token（首次）          → 走首次授權流程並寫入 token
"""
import importlib.util
import sys
import types
import unittest
from pathlib import Path

from google.auth.exceptions import RefreshError

_TOOL = Path(__file__).resolve().parent.parent / "gdrive_update.py"
_spec = importlib.util.spec_from_file_location("gdrive_update", _TOOL)
gdrive_update = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gdrive_update)


class FakeCreds:
    def __init__(self, valid, expired=False, refresh_token=None, refresh_exc=None, tag="loaded"):
        self.valid = valid
        self.expired = expired
        self.refresh_token = refresh_token
        self._refresh_exc = refresh_exc
        self.tag = tag

    def refresh(self, request):
        if self._refresh_exc:
            raise self._refresh_exc
        self.valid = True

    def to_json(self):
        return '{"tag": "%s"}' % self.tag


class GetServiceTest(unittest.TestCase):
    def setUp(self):
        self._orig_modules = {}
        self.state = {"oauth_started": False, "written": None}

    def tearDown(self):
        for name, old in self._orig_modules.items():
            if old is None:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = old

    def _install_fakes(self, loaded_creds):
        """把 get_service() 內部 lazy import 的 google 套件全換成假物件。"""
        state = self.state
        outer = self

        req_mod = types.ModuleType("google.auth.transport.requests")
        req_mod.Request = lambda: object()
        exc_mod = types.ModuleType("google.auth.exceptions")
        exc_mod.RefreshError = RefreshError

        creds_mod = types.ModuleType("google.oauth2.credentials")

        class Credentials:
            @staticmethod
            def from_authorized_user_file(path, scopes):
                return loaded_creds

        creds_mod.Credentials = Credentials

        flow_mod = types.ModuleType("google_auth_oauthlib.flow")

        class InstalledAppFlow:
            @staticmethod
            def from_client_secrets_file(path, scopes):
                return InstalledAppFlow()

            def run_local_server(self, port=0):
                state["oauth_started"] = True
                return FakeCreds(valid=True, tag="fresh_oauth")

        flow_mod.InstalledAppFlow = InstalledAppFlow

        disc_mod = types.ModuleType("googleapiclient.discovery")
        disc_mod.build = lambda *a, **k: ("service", a, k)

        fakes = {
            "google": types.ModuleType("google"),
            "google.auth": types.ModuleType("google.auth"),
            "google.auth.exceptions": exc_mod,
            "google.auth.transport": types.ModuleType("google.auth.transport"),
            "google.auth.transport.requests": req_mod,
            "google.oauth2": types.ModuleType("google.oauth2"),
            "google.oauth2.credentials": creds_mod,
            "google_auth_oauthlib": types.ModuleType("google_auth_oauthlib"),
            "google_auth_oauthlib.flow": flow_mod,
            "googleapiclient": types.ModuleType("googleapiclient"),
            "googleapiclient.discovery": disc_mod,
        }
        for name, mod in fakes.items():
            self._orig_modules[name] = sys.modules.get(name)
            sys.modules[name] = mod

        # 假的 CREDENTIALS_FILE / TOKEN_FILE，攔截 exists() 與 write_text()
        class FakeCredFile:
            def exists(self):
                return True

        class FakeTokenFile:
            def exists(self):
                return loaded_creds is not None

            def write_text(self, text):
                state["written"] = text

        self._patch(gdrive_update, "CREDENTIALS_FILE", FakeCredFile())
        self._patch(gdrive_update, "TOKEN_FILE", FakeTokenFile())
        self._patch(gdrive_update.os, "chmod", lambda *a, **k: None)

    def _patch(self, obj, name, val):
        old = getattr(obj, name)
        self.addCleanup(setattr, obj, name, old)
        setattr(obj, name, val)

    def test_refresh_success(self):
        self._install_fakes(
            FakeCreds(valid=False, expired=True, refresh_token="rt", tag="refreshed")
        )
        svc = gdrive_update.get_service()
        self.assertEqual(svc[0], "service")
        self.assertFalse(self.state["oauth_started"], "正常續期不該啟動 OAuth")
        self.assertIn("refreshed", self.state["written"])

    def test_invalid_grant_reauth(self):
        self._install_fakes(
            FakeCreds(valid=False, expired=True, refresh_token="rt",
                      refresh_exc=RefreshError("invalid_grant: revoked", {"error": "invalid_grant"}))
        )
        svc = gdrive_update.get_service()
        self.assertEqual(svc[0], "service")
        self.assertTrue(self.state["oauth_started"], "撤銷應啟動重新 OAuth")
        self.assertIn("fresh_oauth", self.state["written"])

    def test_other_refresh_error_fail_fast(self):
        self._install_fakes(
            FakeCreds(valid=False, expired=True, refresh_token="rt",
                      refresh_exc=RefreshError("internal_failure", {"error": "internal_failure"}))
        )
        with self.assertRaises(RefreshError):
            gdrive_update.get_service()
        self.assertFalse(self.state["oauth_started"], "暫時性錯誤不該啟動 OAuth")
        self.assertIsNone(self.state["written"], "暫時性錯誤不該寫 token")

    def test_no_token_first_auth(self):
        self._install_fakes(None)
        svc = gdrive_update.get_service()
        self.assertEqual(svc[0], "service")
        self.assertTrue(self.state["oauth_started"], "無 token 應走首次授權")
        self.assertIn("fresh_oauth", self.state["written"])


class IsInvalidGrantTest(unittest.TestCase):
    def test_structured_field(self):
        exc = RefreshError("boom", {"error": "invalid_grant"})
        self.assertTrue(gdrive_update._is_invalid_grant(exc))

    def test_string_fallback(self):
        exc = RefreshError("invalid_grant: Token has been expired or revoked.")
        self.assertTrue(gdrive_update._is_invalid_grant(exc))

    def test_other_error_is_false(self):
        exc = RefreshError("nope", {"error": "invalid_client"})
        self.assertFalse(gdrive_update._is_invalid_grant(exc))


if __name__ == "__main__":
    unittest.main(verbosity=2)
