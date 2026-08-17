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
        import pydantic  # noqa: F401
        return
    except Exception:
        pydantic_module = types.ModuleType("pydantic")

        class _BaseModel:
            def __init__(self, **kwargs):
                for key, value in kwargs.items():
                    setattr(self, key, value)

        def _field(*args, **kwargs):
            return None

        pydantic_module.BaseModel = _BaseModel
        pydantic_module.Field = _field
        sys.modules["pydantic"] = pydantic_module


_ensure_google_genai_stub()
_ensure_webview_stub()
_ensure_dotenv_stub()
_ensure_requests_stub()
_ensure_pydantic_stub()

import main as main_module
from main import Api


def _sample_payload(**overrides):
    payload = {
        "scope": "incomplete",
        "today": "2026-08-05",
        "deliverableCount": 20,
        "omittedCount": 12,
        "buckets": [
            {
                "bucket": "missedHardDeadline",
                "totalCount": 1,
                "deliverables": [
                    {
                        "projectId": "241039",
                        "projectName": "BofA - Crenshaw Ctr.",
                        "deliverableName": "PCC",
                        "due": "2026-05-08",
                        "hardDue": "",
                        "statusText": "On hold",
                    },
                ],
            },
            {
                "bucket": "upcoming",
                "totalCount": 1,
                "deliverables": [],
            },
            {
                "bucket": "undated",
                "totalCount": 1,
                "deliverables": [
                    {
                        "projectId": "260564",
                        "projectName": "Signage Project",
                        "deliverableName": "Signage",
                        "due": "",
                        "hardDue": "",
                        "statusText": "None",
                    },
                ],
            },
        ],
    }
    payload.update(overrides)
    return payload


class DeliverableStatusBriefingPromptTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_prompt_groups_by_bucket_and_reports_omissions(self):
        prompt = self.api._build_deliverable_status_summary_prompt(_sample_payload())

        self.assertIn("Today is 2026-08-05.", prompt)
        self.assertIn("Scope: all incomplete deliverables.", prompt)
        self.assertIn("=== MISSED HARD DEADLINE (1 of 1 listed) ===", prompt)
        self.assertIn("=== NO DUE DATE SET (1 of 1 listed) ===", prompt)
        self.assertIn("[241039]", prompt)
        self.assertIn("internal due 2026-05-08 | hard due none | status On hold", prompt)
        self.assertIn("Listed below: 2.", prompt)
        self.assertIn("Not listed: 12.", prompt)
        # Empty buckets contribute no section at all.
        self.assertNotIn("=== UPCOMING", prompt)
        # Most urgent bucket must come first.
        self.assertLess(
            prompt.index("=== MISSED HARD DEADLINE"),
            prompt.index("=== NO DUE DATE SET"),
        )

    def test_prompt_explains_that_none_status_means_unset(self):
        prompt = self.api._build_deliverable_status_summary_prompt(_sample_payload())
        self.assertIn(
            '"status None" means no status label was ever set',
            prompt,
        )
        self.assertIn("Judge urgency from the dates, not the status.", prompt)

    def test_prompt_survives_an_empty_payload(self):
        prompt = self.api._build_deliverable_status_summary_prompt({})
        self.assertIn("(no deliverables listed)", prompt)

    def test_normalize_status_summary_strips_markdown_bullets(self):
        headline, paragraphs = self.api._normalize_deliverable_status_summary(
            {"headline": "## Two deadlines blown", "paragraphs": ["- one", "* two", "   "]}
        )
        self.assertEqual("Two deadlines blown", headline)
        self.assertEqual(["one", "two"], paragraphs)

    def test_normalize_status_summary_splits_a_single_string(self):
        headline, paragraphs = self.api._normalize_deliverable_status_summary(
            {"headline": "H", "paragraphs": "first block\n\nsecond block"}
        )
        self.assertEqual("H", headline)
        self.assertEqual(["first block", "second block"], paragraphs)


class DeliverableStatusBriefingApiTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def _patched_call(self, payload, fake_client):
        generate_content_calls = []

        class FakeModels:
            def generate_content(self, **kwargs):
                generate_content_calls.append(kwargs)
                return fake_client(**kwargs)

        client = types.SimpleNamespace(models=FakeModels())

        with patch.object(
            self.api, "_ensure_aiohttp", return_value=None
        ), patch(
            "main.genai.Client", return_value=client, create=True
        ), patch(
            "main.types.GenerateContentConfig",
            side_effect=lambda **kwargs: kwargs,
            create=True,
        ), patch(
            "main.types.HttpOptions",
            side_effect=lambda **kwargs: kwargs,
            create=True,
        ), patch(
            "main.types.Content",
            side_effect=lambda **kwargs: kwargs,
            create=True,
        ), patch(
            "main.types.Part",
            types.SimpleNamespace(from_text=lambda text: text),
            create=True,
        ):
            result = self.api.generate_deliverable_status_summary(payload)

        return result, generate_content_calls

    def test_returns_headline_and_paragraphs_on_success(self):
        def fake(**kwargs):
            return types.SimpleNamespace(
                text='{"headline": "H", "paragraphs": ["A", "B"]}'
            )

        payload = _sample_payload(apiKey="test-key")
        result, calls = self._patched_call(payload, fake)

        self.assertEqual("success", result["status"])
        self.assertEqual("H", result["headline"])
        self.assertEqual(["A", "B"], result["paragraphs"])
        self.assertTrue(result["generatedAt"])
        self.assertEqual("gemini-3-flash-preview", calls[0]["model"])
        # Prose, not structured extraction - deliberately not temperature 0.
        self.assertEqual(0.2, calls[0]["config"]["temperature"])
        self.assertEqual("application/json", calls[0]["config"]["response_mime_type"])

    def test_rejects_an_empty_selection_without_calling_the_model(self):
        result, calls = self._patched_call(
            {"buckets": [], "apiKey": "test-key"}, lambda **kwargs: None
        )
        self.assertEqual("error", result["status"])
        self.assertIn("at least one deliverable", result["message"])
        self.assertEqual([], calls)

    def test_missing_api_key_returns_an_error_rather_than_raising(self):
        with patch.dict(main_module.os.environ, {}, clear=True):
            result, calls = self._patched_call(
                _sample_payload(apiKey=""), lambda **kwargs: None
            )
        self.assertEqual("error", result["status"])
        self.assertIn("API key is not configured", result["message"])
        self.assertEqual([], calls)

    def test_expired_key_error_points_at_google_ai_studio(self):
        def fake(**kwargs):
            raise Exception("API key expired. Please renew the API key.")

        result, _ = self._patched_call(_sample_payload(apiKey="k"), fake)
        self.assertEqual("error", result["status"])
        self.assertIn("Google AI Studio", result["message"])

    def test_rate_limit_error_is_mapped(self):
        def fake(**kwargs):
            raise Exception("429 quota exceeded for this project")

        result, _ = self._patched_call(_sample_payload(apiKey="k"), fake)
        self.assertEqual("error", result["status"])
        self.assertIn("rate limit", result["message"].lower())

    def test_empty_model_response_is_an_error(self):
        def fake(**kwargs):
            return types.SimpleNamespace(text="")

        result, _ = self._patched_call(_sample_payload(apiKey="k"), fake)
        self.assertEqual("error", result["status"])
        self.assertIn("empty briefing", result["message"])

    def test_invalid_json_response_is_an_error(self):
        def fake(**kwargs):
            return types.SimpleNamespace(text="not json at all")

        result, _ = self._patched_call(_sample_payload(apiKey="k"), fake)
        self.assertEqual("error", result["status"])
        self.assertIn("invalid JSON", result["message"])


if __name__ == "__main__":
    unittest.main()
