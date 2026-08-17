import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
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

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import main as main_module
from main import Api


class _FakeMailItem:
    def __init__(self, subject="Quarterly RFI", entry_id="ENTRY-1", save_error=None):
        self.Subject = subject
        self.EntryID = entry_id
        self._save_error = save_error
        self.saved_to = None
        self.saved_format = None

    def SaveAs(self, path, fmt):
        if self._save_error is not None:
            raise self._save_error
        self.saved_to = path
        self.saved_format = fmt
        with open(path, "wb") as handle:
            handle.write(b"fake-msg")


class _FakeSelection:
    def __init__(self, items):
        self._items = list(items)
        self.Count = len(self._items)

    def Item(self, index):
        return self._items[index - 1]


class _FakeExplorer:
    def __init__(self, selection):
        self.Selection = selection


class _FakeApplication:
    def __init__(self, explorer):
        self._explorer = explorer

    def ActiveExplorer(self):
        return self._explorer


class SaveActiveOutlookSelectionTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def _run(self, mail_items, temp_dir, context=None):
        application = _FakeApplication(_FakeExplorer(_FakeSelection(mail_items)))
        with patch.object(main_module, "TASKS_FILE", str(Path(temp_dir) / "tasks.json")), \
                patch.object(
                    Api,
                    "_get_desktop_outlook_namespace",
                    lambda _self: (application, object()),
                ), \
                patch.object(
                    Api,
                    "_get_desktop_outlook_internet_message_id",
                    lambda _self, _item: "<abc@contoso.com>",
                ):
            return self.api.save_active_outlook_selection(context)

    def test_selection_is_saved_as_a_managed_msg_copy(self):
        with tempfile.TemporaryDirectory(prefix="acies-page-email-") as temp_dir:
            item = _FakeMailItem()
            result = self._run(
                [item],
                temp_dir,
                {"projectId": "proj_9", "deliverableId": "proj_9_subpage_a"},
            )

            self.assertEqual("success", result["status"])
            ref = result["emailRef"]
            self.assertEqual("saved-file", ref["source"])
            self.assertEqual("Quarterly RFI", ref["label"])
            self.assertEqual("ENTRY-1", ref["messageId"])
            self.assertEqual("<abc@contoso.com>", ref["internetMessageId"])
            self.assertTrue(ref["url"].startswith("file:///"))

            saved_path = Path(ref["raw"])
            self.assertTrue(saved_path.exists())
            self.assertEqual(3, item.saved_format)  # olMSG
            # Lands under email-links/<projectHint>/<ownerHint>/
            expected_dir = Path(temp_dir) / "email-links" / "proj_9" / "proj_9_subpage_a"
            self.assertEqual(expected_dir.resolve(), saved_path.parent.resolve())

    def test_path_components_are_sanitized_and_contained(self):
        with tempfile.TemporaryDirectory(prefix="acies-page-email-") as temp_dir:
            item = _FakeMailItem(subject='RE: bid/quote <urgent>')
            result = self._run(
                [item],
                temp_dir,
                {"projectId": "../../escape", "deliverableId": "..\\..\\also"},
            )

            self.assertEqual("success", result["status"])
            saved_path = Path(result["emailRef"]["raw"]).resolve()
            email_root = (Path(temp_dir) / "email-links").resolve()
            self.assertEqual(
                str(email_root), os.path.commonpath([str(email_root), str(saved_path)])
            )
            for illegal in ('/', '\\', '<', '>', ':'):
                self.assertNotIn(illegal, saved_path.name)

    def test_save_failure_degrades_to_an_entry_id_reference(self):
        with tempfile.TemporaryDirectory(prefix="acies-page-email-") as temp_dir:
            item = _FakeMailItem(save_error=RuntimeError("Outlook refused"))
            result = self._run([item], temp_dir, {"projectId": "proj_9"})

            self.assertEqual("success", result["status"])
            ref = result["emailRef"]
            # open_outlook_desktop_message still resolves this shape.
            self.assertEqual("outlook-desktop", ref["source"])
            self.assertEqual("ENTRY-1", ref["raw"])
            self.assertEqual("ENTRY-1", ref["messageId"])

    def test_empty_selection_reports_an_error(self):
        with tempfile.TemporaryDirectory(prefix="acies-page-email-") as temp_dir:
            result = self._run([], temp_dir)
            self.assertEqual("error", result["status"])
            self.assertIn("selected", result["message"].lower())


class PageEmailPdfRenderTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def _render(self, html):
        return self.api._prepare_project_page_pdf_html(html)

    def test_chip_renders_as_a_labelled_email_reference(self):
        rendered = self._render(
            '<p data-important="true">Call the GC'
            '<span class="page-email" data-email-raw="C:\\mail\\x.msg" '
            'data-email-label="RE: Panel schedule" contenteditable="false">'
            '<span class="page-email-icon">@</span>'
            '<button data-email-action="open">RE: Panel schedule</button>'
            '<button data-email-action="remove">x</button>'
            "</span></p>"
        )
        self.assertIn("<b>Email:</b> RE: Panel schedule", rendered)
        # The chip's own controls must not survive into the PDF.
        self.assertNotIn("data-email-action", rendered)
        self.assertNotIn("page-email-icon", rendered)
        # The important flag on the surrounding block is untouched.
        self.assertIn("pdf-important-flag", rendered)
        self.assertIn("Call the GC", rendered)

    def test_chip_without_a_label_falls_back_to_the_file_name(self):
        rendered = self._render(
            '<p><span class="page-email" data-email-raw="C:\\mail\\2026_bid.msg">'
            "<button data-email-action=\"open\">x</button></span></p>"
        )
        self.assertIn("<b>Email:</b> 2026_bid.msg", rendered)


if __name__ == "__main__":
    unittest.main()
