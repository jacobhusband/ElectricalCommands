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


class _FixedDateTime(datetime.datetime):
    @classmethod
    def now(cls, tz=None):
        base = cls(2026, 3, 23, 12, 0, 0, tzinfo=datetime.timezone.utc)
        return base.astimezone(tz) if tz else base


class DesktopOutlookScanTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_get_outlook_scan_capability_reports_desktop_availability(self):
        with patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(True, ""),
        ):
            result = self.api.get_outlook_scan_capability()

        self.assertEqual(
            {
                "status": "success",
                "desktopAvailable": True,
                "desktopReason": "",
            },
            result,
        )

    def test_scan_outlook_inbox_returns_desktop_unavailable_error(self):
        with patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(False, "Classic Outlook is not installed on this machine."),
        ), patch.object(self.api, "_send_outlook_scan_progress") as progress_mock:
            result = self.api.scan_outlook_inbox(
                {"scanDate": "2026-03-20", "projectContext": []},
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        self.assertEqual("error", result["status"])
        self.assertEqual(main_module.OUTLOOK_SCAN_SOURCE_DESKTOP, result["source"])
        self.assertIn("Classic Outlook is not installed", result["message"])
        self.assertEqual("2026-03-20", result["scanDate"])
        self.assertEqual("week", result["timeframe"])
        self.assertEqual([], result["suggestions"])
        self.assertEqual([], result["skippedMessages"])
        self.assertIn("candidateCount", result)
        self.assertIn("scannedCount", result)
        self.assertTrue(
            any(call.args and call.args[0] == "error" for call in progress_mock.call_args_list)
        )

    def test_scan_outlook_inbox_returns_desktop_read_error_without_fallback(self):
        with patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(True, ""),
        ), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            side_effect=RuntimeError("MAPI unavailable"),
        ), patch.object(self.api, "_send_outlook_scan_progress"):
            result = self.api.scan_outlook_inbox(
                {"scanDate": "2026-03-20", "projectContext": []},
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        self.assertEqual("error", result["status"])
        self.assertEqual(main_module.OUTLOOK_SCAN_SOURCE_DESKTOP, result["source"])
        self.assertEqual(
            "Desktop Outlook could not be read: MAPI unavailable",
            result["message"],
        )
        self.assertNotIn("Microsoft", result["message"])

    def test_scan_outlook_inbox_uses_desktop_messages_when_available(self):
        payload = {
            "scanDate": "2026-03-20",
            "projectContext": [{"id": "proj-1", "deliverables": [{"name": "Submittal"}]}],
        }
        message_summaries = [{"id": "msg-1", "subject": "Need updated submittal"}]
        expected_result = {
            "status": "success",
            "source": main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
            "scanDate": "2026-03-20",
            "timeframe": "week",
            "suggestions": [],
            "skippedMessages": [],
        }

        with patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(True, ""),
        ), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            return_value=(message_summaries, False),
        ) as list_mock, patch.object(
            self.api,
            "_analyze_outlook_scan_batch",
            return_value=expected_result,
        ) as analyze_mock, patch.object(self.api, "_send_outlook_scan_progress"):
            result = self.api.scan_outlook_inbox(
                payload,
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        self.assertEqual(expected_result, result)
        list_mock.assert_called_once_with(
            timeframe="week",
            limit=main_module.OUTLOOK_SCAN_FETCH_LIMIT,
            scan_date="2026-03-20",
        )
        args, kwargs = analyze_mock.call_args
        self.assertEqual(message_summaries, args[0])
        self.assertEqual("proj-1", args[2][0]["id"])
        self.assertEqual("", args[2][0]["name"])
        self.assertEqual("Submittal", args[2][0]["deliverables"][0]["name"])
        self.assertEqual("api-key", args[3])
        self.assertEqual("Jacob Husband", args[4])
        self.assertEqual(["Electrical"], args[5])
        self.assertEqual("week", args[6])
        self.assertEqual(main_module.OUTLOOK_SCAN_SOURCE_DESKTOP, args[7])
        self.assertFalse(kwargs["has_more"])
        self.assertEqual("2026-03-20", kwargs["scan_date"])

    def test_scan_outlook_inbox_invalid_scan_date_defaults_to_today(self):
        with patch.object(
            main_module.datetime,
            "datetime",
            _FixedDateTime,
        ), patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(True, ""),
        ), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            return_value=([], False),
        ) as list_mock, patch.object(
            self.api,
            "_analyze_outlook_scan_batch",
            return_value={"status": "success", "suggestions": [], "skippedMessages": []},
        ), patch.object(self.api, "_send_outlook_scan_progress"):
            self.api.scan_outlook_inbox(
                {"scanDate": "not-a-date", "projectContext": []},
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        list_mock.assert_called_once_with(
            timeframe="week",
            limit=main_module.OUTLOOK_SCAN_FETCH_LIMIT,
            scan_date="2026-03-23",
        )

    def test_open_outlook_desktop_message_displays_mail_item(self):
        class _FakeMailItem:
            def __init__(self):
                self.displayed = False

            def Display(self):
                self.displayed = True

        class _FakeNamespace:
            def __init__(self, item):
                self.item = item
                self.last_id = None

            def GetItemFromID(self, message_id):
                self.last_id = message_id
                return self.item

        mail_item = _FakeMailItem()
        namespace = _FakeNamespace(mail_item)

        with patch.object(
            self.api,
            "_get_desktop_outlook_namespace",
            return_value=(None, namespace),
        ), patch.object(
            self.api,
            "_run_with_outlook_com",
            side_effect=lambda callback: callback(),
        ):
            result = self.api.open_outlook_desktop_message({"messageId": "abc123"})

        self.assertEqual({"status": "success"}, result)
        self.assertEqual("abc123", namespace.last_id)
        self.assertTrue(mail_item.displayed)


class OutlookScanBodyReductionTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def _build_summary(self, **overrides):
        summary = {
            "id": "msg-1",
            "subject": "Need updated submittal",
            "bodyPreview": "Need updated submittal by Friday.",
            "receivedDateTime": "2026-03-20T17:00:00Z",
            "internetMessageId": "<msg-1@example.com>",
            "conversationId": "conversation-1",
            "hasAttachments": True,
            "source": main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
            "from": {
                "name": "Alice Smith",
                "address": "alice@example.com",
            },
        }
        summary.update(overrides)
        return summary

    def test_build_outlook_scan_analysis_body_strips_reply_history_and_signature(self):
        raw_body = """Please issue the revised lighting submittal by Friday.

Thanks,
Jacob Husband
Project Manager
ACIES
555-111-2222

From: Architect <architect@example.com>
Sent: Thursday, March 20, 2026 8:00 AM
To: Jacob Husband <jacob@example.com>
Subject: Re: Lighting submittal
"""

        result = self.api._build_outlook_scan_analysis_body(raw_body)

        self.assertEqual("Please issue the revised lighting submittal by Friday.", result)

    def test_build_outlook_scan_analysis_body_strips_disclaimer_and_device_footer(self):
        raw_body = """Need panel schedule update today.

Sent from my iPhone

This email and any attachments are intended only for the named recipient. If you are not the intended recipient, please delete this email immediately and notify the sender."""

        result = self.api._build_outlook_scan_analysis_body(raw_body)

        self.assertEqual("Need panel schedule update today.", result)

    def test_build_outlook_scan_analysis_body_preserves_basic_text_without_signature(self):
        raw_body = """Need updated panel schedules.
Please coordinate with the architect today."""

        result = self.api._build_outlook_scan_analysis_body(raw_body)

        self.assertEqual(
            "Need updated panel schedules.\nPlease coordinate with the architect today.",
            result,
        )

    def test_build_outlook_scan_analysis_body_falls_back_when_cleaning_empties_message(self):
        raw_body = """CAUTION: External Email
Use caution with links and attachments."""

        result = self.api._build_outlook_scan_analysis_body(raw_body)

        self.assertEqual(
            "CAUTION: External Email\nUse caution with links and attachments.",
            result,
        )

    def test_analyze_outlook_scan_batch_uses_minimal_prompt_payload(self):
        summary = self._build_summary()
        captured = {}

        def _capture_prompt(prompt, api_key):
            captured["prompt"] = prompt
            return []

        with patch.object(
            self.api,
            "_request_outlook_scan_batch_suggestions",
            side_effect=_capture_prompt,
        ), patch.object(self.api, "_send_outlook_scan_progress"):
            result = self.api._analyze_outlook_scan_batch(
                [summary],
                lambda current: (
                    current,
                    """Need updated submittal by Friday.

Thanks,
Alice Smith
Project Engineer
alice@example.com""",
                ),
                [{"id": "proj-1", "name": "Client Project", "deliverables": []}],
                "api-key",
                "Jacob Husband",
                ["Electrical"],
                "week",
                main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
                scan_date="2026-03-20",
            )

        prompt = captured["prompt"]
        self.assertEqual("success", result["status"])
        self.assertIn('"messageRef":"E001"', prompt)
        self.assertIn('"from":{"name":"Alice Smith"}', prompt)
        self.assertIn('"body":"Need updated submittal by Friday."', prompt)
        self.assertNotIn('"messageId"', prompt)
        self.assertNotIn('"internetMessageId"', prompt)
        self.assertNotIn('"conversationId"', prompt)
        self.assertNotIn('"address"', prompt)
        self.assertNotIn("Thanks,", prompt)

    def test_analyze_outlook_scan_batch_retries_with_reduced_prompt_after_timeout(self):
        summary = self._build_summary()
        timeout_error = (
            "AI error: 504 DEADLINE_EXCEEDED. "
            "{'error': {'code': 504, 'message': 'Deadline expired before operation could complete.', "
            "'status': 'DEADLINE_EXCEEDED'}}"
        )
        prompts = []
        long_body = "Need updated submittal by Friday. " * 200

        def _request(prompt, api_key):
            prompts.append(prompt)
            if len(prompts) == 1:
                raise RuntimeError(timeout_error)
            return []

        with patch.object(
            self.api,
            "_request_outlook_scan_batch_suggestions",
            side_effect=_request,
        ), patch.object(self.api, "_send_outlook_scan_progress") as progress_mock:
            result = self.api._analyze_outlook_scan_batch(
                [summary],
                lambda current: (current, long_body),
                [],
                "api-key",
                "Jacob Husband",
                ["Electrical"],
                "week",
                main_module.OUTLOOK_SCAN_SOURCE_DESKTOP,
                scan_date="2026-03-20",
            )

        self.assertEqual("success", result["status"])
        self.assertEqual(2, len(prompts))
        self.assertLess(len(prompts[1]), len(prompts[0]))
        self.assertTrue(
            any(
                len(call.args) >= 2
                and call.args[0] == "reviewing_ai"
                and "Retrying with a reduced text-only batch" in call.args[1]
                for call in progress_mock.call_args_list
            )
        )

    def test_scan_outlook_inbox_returns_clear_error_after_timeout_retry_failure(self):
        summary = self._build_summary()
        timeout_error = (
            "AI error: 504 DEADLINE_EXCEEDED. "
            "{'error': {'code': 504, 'message': 'Deadline expired before operation could complete.', "
            "'status': 'DEADLINE_EXCEEDED'}}"
        )

        with patch.object(
            self.api,
            "_get_desktop_outlook_availability",
            return_value=(True, ""),
        ), patch.object(
            self.api,
            "_list_desktop_outlook_inbox_messages",
            return_value=([summary], False),
        ), patch.object(
            self.api,
            "_get_desktop_outlook_message_body_text",
            return_value=(summary, "Need updated submittal by Friday. " * 200),
        ), patch.object(
            self.api,
            "_request_outlook_scan_batch_suggestions",
            side_effect=RuntimeError(timeout_error),
        ) as request_mock, patch.object(self.api, "_send_outlook_scan_progress"):
            result = self.api.scan_outlook_inbox(
                {"scanDate": "2026-03-20", "projectContext": []},
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        self.assertEqual("error", result["status"])
        self.assertIn("reduced text-only batch", result["message"])
        self.assertEqual(2, request_mock.call_count)
        self.assertEqual([], result["suggestions"])


class DeliverableAiPromptContractTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_build_email_analysis_prompt_separates_tasks_from_notes(self):
        prompt = self.api._build_email_analysis_prompt(
            "Please revise the lighting plans and wait for architect approval before final issue.",
            "Jacob Husband",
            ["Electrical"],
        )

        self.assertIn(
            'every actionable Electrical engineering step needed to complete the deliverable',
            prompt,
        )
        self.assertIn(
            'Provide only non-actionable context the user should note',
            prompt,
        )
        self.assertIn(
            'Do NOT put action items, next steps, or completion work in notes',
            prompt,
        )
        self.assertIn('If there is no notable non-actionable context, return "".', prompt)
        self.assertIn(
            'Example: ["Revise lighting plans per architect comments", "Update panel schedules", "Issue updated PDF set for resubmittal"]',
            prompt,
        )
        self.assertIn('Example: "Awaiting architect approval before final issue."', prompt)
        self.assertNotIn('summary of the email', prompt)

    def test_build_email_analysis_prompt_includes_known_project_matching_context(self):
        prompt = self.api._build_email_analysis_prompt(
            "Subject: Canoga BofA submittal\nPlease update the lighting package.",
            "Jacob Husband",
            ["Electrical"],
            [
                {
                    "id": "250597",
                    "name": "BofA Sherman",
                    "nick": "Canoga",
                    "path": r"M:\Projects\250597 BofA Sherman",
                }
            ],
        )

        self.assertIn("KNOWN_PROJECTS:", prompt)
        self.assertIn('"id":"250597"', prompt)
        self.assertIn('"name":"BofA Sherman"', prompt)
        self.assertIn('"nick":"Canoga"', prompt)
        self.assertIn('"pathLeaf":"250597 BofA Sherman"', prompt)
        self.assertIn("Prefer exact known-project identity", prompt)
        self.assertIn('return that known project\'s exact "id", "name", and "path"', prompt)

    def test_normalize_email_intake_project_context_ignores_malformed_and_deliverables(self):
        result = self.api._normalize_email_intake_project_context(
            [
                "not-a-project",
                {"deliverables": [{"name": "Submittal"}]},
                {
                    "id": " 250597 ",
                    "name": " BofA Sherman ",
                    "nick": " Canoga ",
                    "path": r"M:\Projects\250597 BofA Sherman",
                    "deliverables": [{"name": "Should not be included"}],
                },
            ]
        )

        self.assertEqual(
            [
                {
                    "id": "250597",
                    "name": "BofA Sherman",
                    "nick": "Canoga",
                    "path": r"M:\Projects\250597 BofA Sherman",
                    "pathLeaf": "250597 BofA Sherman",
                }
            ],
            result,
        )
        self.assertNotIn("deliverables", result[0])

    def test_normalize_email_intake_project_context_caps_project_count(self):
        result = self.api._normalize_email_intake_project_context(
            [
                {"id": f"P-{index:04d}", "name": f"Project {index}"}
                for index in range(main_module.EMAIL_INTAKE_PROJECT_CONTEXT_MAX_PROJECTS + 5)
            ]
        )

        self.assertEqual(main_module.EMAIL_INTAKE_PROJECT_CONTEXT_MAX_PROJECTS, len(result))
        self.assertEqual("P-0000", result[0]["id"])

    def test_process_email_with_ai_remains_backward_compatible_without_project_context(self):
        expected = {
            "id": "",
            "name": "Client Project",
            "due": "",
            "path": "",
            "deliverable": "",
            "tasks": [],
            "notes": "",
        }

        with patch.object(
            self.api,
            "_extract_project_data_from_email_text",
            return_value=expected,
        ) as extract_mock:
            result = self.api.process_email_with_ai(
                "Please update the submittal.",
                "api-key",
                "Jacob Husband",
                ["Electrical"],
            )

        self.assertEqual({"status": "success", "data": expected}, result)
        extract_mock.assert_called_once_with(
            "Please update the submittal.",
            "api-key",
            "Jacob Husband",
            ["Electrical"],
            None,
        )

    def test_build_outlook_scan_batch_prompt_separates_tasks_from_notes(self):
        prompt = self.api._build_outlook_scan_batch_prompt(
            included_emails=[],
            included_deliverables=[],
            user_name="Jacob Husband",
            discipline=["Electrical"],
            timeframe="week",
            scan_date="2026-03-20",
        )

        self.assertIn(
            'every actionable Electrical engineering step needed to complete the deliverable',
            prompt,
        )
        self.assertIn(
            'Provide only non-actionable context the user should note',
            prompt,
        )
        self.assertIn(
            'Do NOT put action items, next steps, or completion work in "notes"',
            prompt,
        )
        self.assertIn('If there is no notable non-actionable context, return "".', prompt)
        self.assertNotIn('summary of the email', prompt)

    def test_normalize_email_project_data_keeps_action_only_notes_blank(self):
        result = self.api._normalize_email_project_data(
            {
                "id": "250597",
                "name": "Client Project",
                "deliverable": "Submittal",
                "tasks": [
                    "Update CAD per architect comments",
                    "Issue revised submittal package",
                ],
                "notes": "",
            }
        )

        self.assertEqual("", result["notes"])
        self.assertEqual(
            [
                {"text": "Update CAD per architect comments", "done": False, "links": []},
                {"text": "Issue revised submittal package", "done": False, "links": []},
            ],
            result["tasks"],
        )

    def test_normalize_email_project_data_preserves_context_notes_and_tasks(self):
        result = self.api._normalize_email_project_data(
            {
                "tasks": [
                    {"text": "Revise one-line diagram", "done": True, "links": ["ref-1"]},
                    "Coordinate revised load letter",
                ],
                "notes": "Awaiting utility response before final issue.",
            }
        )

        self.assertEqual("Awaiting utility response before final issue.", result["notes"])
        self.assertEqual(
            [
                {"text": "Revise one-line diagram", "done": True, "links": ["ref-1"]},
                {"text": "Coordinate revised load letter", "done": False, "links": []},
            ],
            result["tasks"],
        )

    def test_normalize_outlook_scan_batch_suggestions_preserves_context_notes_and_tasks(self):
        result = self.api._normalize_outlook_scan_batch_suggestions(
            {
                "suggestions": [
                    {
                        "projectId": "proj-1",
                        "projectName": "Client Project",
                        "projectPath": "M:\\Project",
                        "deliverable": "Bulletin #2",
                        "due": "03/30/26",
                        "tasks": [
                            "Revise branch circuiting",
                            {"text": "Update fixture schedule", "done": False, "links": ["email-1"]},
                        ],
                        "notes": "Pricing must be called out separately from the bulletin package.",
                        "supportingMessageRefs": ["E001"],
                    }
                ]
            }
        )

        self.assertEqual(1, len(result))
        self.assertEqual(
            [
                {"text": "Revise branch circuiting", "done": False, "links": []},
                {"text": "Update fixture schedule", "done": False, "links": ["email-1"]},
            ],
            result[0]["deliverable"]["tasks"],
        )
        self.assertEqual(
            "Pricing must be called out separately from the bulletin package.",
            result[0]["deliverable"]["notes"],
        )


if __name__ == "__main__":
    unittest.main()
