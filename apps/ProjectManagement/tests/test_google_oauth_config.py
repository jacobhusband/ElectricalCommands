import datetime
import sys
import types
import unittest
from unittest.mock import patch


def _ensure_google_genai_stub():
    try:
        from google import genai as _genai  # noqa: F401
        from google.genai import types as _types  # noqa: F401
        return
    except Exception:
        google_module = sys.modules.get("google")
        if google_module is None:
            google_module = types.ModuleType("google")
            google_module.__path__ = []
            sys.modules["google"] = google_module

        genai_module = types.ModuleType("google.genai")
        genai_types_module = types.ModuleType("google.genai.types")
        genai_module.types = genai_types_module
        google_module.genai = genai_module

        sys.modules["google.genai"] = genai_module
        sys.modules["google.genai.types"] = genai_types_module


def _ensure_webview_stub():
    try:
        import webview  # noqa: F401
        return
    except Exception:
        webview_module = types.ModuleType("webview")
        webview_module.windows = []
        webview_module.create_window = lambda *args, **kwargs: None
        webview_module.start = lambda *args, **kwargs: None
        sys.modules["webview"] = webview_module


def _ensure_dotenv_stub():
    try:
        from dotenv import load_dotenv as _load_dotenv  # noqa: F401
        return
    except Exception:
        dotenv_module = types.ModuleType("dotenv")
        dotenv_module.load_dotenv = lambda *args, **kwargs: False
        sys.modules["dotenv"] = dotenv_module


def _ensure_requests_stub():
    try:
        import requests  # noqa: F401
        return
    except Exception:
        requests_module = types.ModuleType("requests")
        requests_module.get = lambda *args, **kwargs: None
        requests_module.post = lambda *args, **kwargs: None
        sys.modules["requests"] = requests_module


def _ensure_pydantic_stub():
    try:
        from pydantic import BaseModel as _BaseModel, Field as _Field  # noqa: F401
        return
    except Exception:
        pydantic_module = types.ModuleType("pydantic")

        class BaseModel:
            pass

        def Field(*args, **kwargs):
            if args:
                return args[0]
            return kwargs.get("default")

        pydantic_module.BaseModel = BaseModel
        pydantic_module.Field = Field
        sys.modules["pydantic"] = pydantic_module


_ensure_google_genai_stub()
_ensure_webview_stub()
_ensure_dotenv_stub()
_ensure_requests_stub()
_ensure_pydantic_stub()

import main as main_module
from main import Api


class _FakeResponse:
    def __init__(self, status_code=200, payload=None, text=""):
        self.status_code = status_code
        self._payload = payload or {}
        self.text = text

    def json(self):
        return self._payload


class GoogleOAuthConfigTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_exchange_google_auth_code_omits_client_secret_when_not_configured(self):
        captured = {}

        def fake_post(url, data=None, timeout=None):
            captured["url"] = url
            captured["data"] = dict(data or {})
            captured["timeout"] = timeout
            return _FakeResponse(payload={"access_token": "token"})

        with patch.object(self.api, "_get_google_oauth_client_secret", return_value=""), patch.object(
            main_module.requests, "post", side_effect=fake_post
        ):
            result = self.api._exchange_google_auth_code(
                "client-id",
                "auth-code",
                "verifier",
                "http://127.0.0.1/callback",
            )

        self.assertEqual("token", result["access_token"])
        self.assertNotIn("client_secret", captured["data"])
        self.assertEqual("client-id", captured["data"]["client_id"])
        self.assertEqual("auth-code", captured["data"]["code"])

    def test_exchange_google_auth_code_includes_client_secret_when_configured(self):
        captured = {}

        def fake_post(url, data=None, timeout=None):
            captured["data"] = dict(data or {})
            return _FakeResponse(payload={"access_token": "token"})

        with patch.object(
            self.api, "_get_google_oauth_client_secret", return_value="secret-value"
        ), patch.object(main_module.requests, "post", side_effect=fake_post):
            self.api._exchange_google_auth_code(
                "client-id",
                "auth-code",
                "verifier",
                "http://127.0.0.1/callback",
            )

        self.assertEqual("secret-value", captured["data"]["client_secret"])

    def test_extract_google_error_message_mentions_missing_client_secret_configuration(self):
        response = _FakeResponse(
            status_code=400,
            payload={"error_description": "client_secret is missing"},
            text="client_secret is missing",
        )

        with patch.object(self.api, "_get_google_oauth_client_secret", return_value=""):
            message = self.api._extract_google_error_message(response)

        self.assertIn(main_module.GOOGLE_OAUTH_CLIENT_SECRET_ENV, message)
        self.assertIn("verify the Google OAuth client settings", message)

    def test_refresh_google_auth_record_includes_client_secret_when_configured(self):
        captured = {}
        stale_expiry = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(minutes=10)
        existing_auth = {
            "provider": "google",
            "subject": "subject-1",
            "email": "user@example.com",
            "displayName": "User Example",
            "avatarUrl": "https://example.com/avatar.png",
            "signedInAt": "2026-01-01T00:00:00Z",
            "refreshToken": "refresh-token",
            "expiresAt": "2026-01-01T00:00:00Z",
        }

        def fake_post(url, data=None, timeout=None):
            captured["data"] = dict(data or {})
            return _FakeResponse(
                payload={
                    "access_token": "refreshed-token",
                    "expires_in": 3600,
                    "token_type": "Bearer",
                    "scope": "openid email profile",
                }
            )

        with patch.object(self.api, "_get_google_oauth_client_id", return_value="client-id"), patch.object(
            self.api, "_get_google_oauth_client_secret", return_value="secret-value"
        ), patch.object(main_module, "parse_utc_iso", return_value=stale_expiry), patch.object(
            main_module.requests, "post", side_effect=fake_post
        ), patch.object(
            self.api, "_build_google_auth_record", return_value={"provider": "google"}
        ) as build_auth, patch.object(
            self.api, "_persist_google_auth_record"
        ) as persist_auth:
            refreshed = self.api._refresh_google_auth_record_if_needed(existing_auth)

        self.assertEqual("secret-value", captured["data"]["client_secret"])
        self.assertEqual("refresh-token", captured["data"]["refresh_token"])
        self.assertEqual("google", refreshed["provider"])
        self.assertEqual(existing_auth["signedInAt"], refreshed["signedInAt"])
        build_auth.assert_called_once()
        persist_auth.assert_called_once_with(refreshed)


if __name__ == "__main__":
    unittest.main()
