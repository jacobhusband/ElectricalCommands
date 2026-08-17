import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch


def _ensure_optional_dependency_stubs():
    try:
        from google import genai as _genai  # noqa: F401
        from google.genai import types as _types  # noqa: F401
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

    try:
        import webview  # noqa: F401
    except Exception:
        webview_module = types.ModuleType("webview")
        webview_module.windows = []
        webview_module.create_window = lambda *args, **kwargs: None
        webview_module.start = lambda *args, **kwargs: None
        sys.modules["webview"] = webview_module

    try:
        from dotenv import load_dotenv as _load_dotenv  # noqa: F401
    except Exception:
        dotenv_module = types.ModuleType("dotenv")
        dotenv_module.load_dotenv = lambda *args, **kwargs: False
        sys.modules["dotenv"] = dotenv_module


_ensure_optional_dependency_stubs()

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import openpyxl

import main as main_module
from main import Api


class PageWorkbookBackendTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    @staticmethod
    def _patched_data_path(temp_dir):
        return patch.object(
            main_module,
            "get_app_data_path",
            lambda name="tasks.json": str(Path(temp_dir) / name),
        )

    def test_create_inspect_open_and_delete_workbook(self):
        with tempfile.TemporaryDirectory(prefix="acies-page-files-") as temp_dir:
            with self._patched_data_path(temp_dir):
                first = self.api.create_page_workbook("proj_260243", "Load: Calculations.xlsx")
                second = self.api.create_page_workbook("proj_260243", "Load: Calculations")

                self.assertEqual("success", first["status"])
                self.assertEqual("Load_ Calculations.xlsx", first["fileName"])
                self.assertNotEqual(first["fileRef"], second["fileRef"])
                self.assertTrue(first["fileRef"].startswith("page_files/proj_260243/"))

                first_path = Path(self.api._resolve_page_file_path(first["fileRef"]))
                self.assertTrue(first_path.is_file())
                self.assertIn(str(Path(temp_dir, "page_files").resolve()), str(first_path.resolve()))
                workbook = openpyxl.load_workbook(first_path)
                try:
                    self.assertEqual(["Sheet1"], workbook.sheetnames)
                finally:
                    workbook.close()

                info = self.api.get_page_file_info(first["fileRef"])
                self.assertEqual("success", info["status"])
                self.assertTrue(info["exists"])
                self.assertGreater(info["sizeBytes"], 0)

                with patch.object(self.api, "open_path", return_value={"status": "success"}) as opener:
                    opened = self.api.open_page_file(first["fileRef"])
                self.assertEqual("success", opened["status"])
                opener.assert_called_once_with(str(first_path))

                deleted = self.api.delete_page_file(first["fileRef"])
                self.assertEqual("success", deleted["status"])
                self.assertTrue(deleted["deleted"])
                self.assertFalse(first_path.exists())
                self.assertTrue(Path(self.api._resolve_page_file_path(second["fileRef"])).is_file())

                deleted_again = self.api.delete_page_file(first["fileRef"])
                self.assertEqual("success", deleted_again["status"])
                self.assertFalse(deleted_again["deleted"])

    def test_workbook_paths_and_names_are_sandboxed(self):
        with tempfile.TemporaryDirectory(prefix="acies-page-files-safety-") as temp_dir:
            with self._patched_data_path(temp_dir):
                self.assertEqual(
                    "error",
                    self.api.create_page_workbook("proj_x", "   ")["status"],
                )
                saved = self.api.create_page_workbook("proj_x/../secret", "CON")
                self.assertEqual("success", saved["status"])
                self.assertEqual("_CON.xlsx", saved["fileName"])
                self.assertNotIn("..", saved["fileRef"])

                bad_refs = [
                    "../secret.xlsx",
                    "page_files/proj/id/../../secret.xlsx",
                    "page_files\\proj\\id\\..\\secret.xlsx",
                    str(Path(temp_dir) / "secret.xlsx"),
                    "page_files/proj/id/not-a-workbook.xls",
                ]
                for bad_ref in bad_refs:
                    self.assertEqual("error", self.api.get_page_file_info(bad_ref)["status"])
                    self.assertEqual("error", self.api.open_page_file(bad_ref)["status"])
                    self.assertEqual("error", self.api.delete_page_file(bad_ref)["status"])

    def test_batch_delete_reports_invalid_refs_without_hiding_successes(self):
        with tempfile.TemporaryDirectory(prefix="acies-page-files-batch-") as temp_dir:
            with self._patched_data_path(temp_dir):
                saved = self.api.create_page_workbook("page_global_abc", "Budget")
                result = self.api.delete_page_files([saved["fileRef"], "../outside.xlsx"])

                self.assertEqual("partial", result["status"])
                self.assertEqual(1, result["deleted"])
                self.assertEqual(1, len(result["failed"]))
                self.assertFalse(os.path.exists(self.api._resolve_page_file_path(saved["fileRef"])))

    def test_link_existing_workbook_opens_original_and_only_removes_link(self):
        with tempfile.TemporaryDirectory(prefix="acies-page-file-links-") as temp_dir:
            existing_path = Path(temp_dir) / "Existing Budget.xlsm"
            existing_path.write_bytes(b"existing workbook placeholder")
            with self._patched_data_path(temp_dir):
                linked = self.api.link_page_workbook(f'"{existing_path}"')

                self.assertEqual("success", linked["status"])
                self.assertEqual("external", linked["storageType"])
                self.assertTrue(linked["fileRef"].startswith("page_file_links/"))
                self.assertNotIn(str(existing_path), linked["fileRef"])

                info = self.api.get_page_file_info(linked["fileRef"])
                self.assertEqual("success", info["status"])
                self.assertTrue(info["exists"])
                self.assertEqual("external", info["storageType"])
                self.assertEqual(existing_path.name, info["fileName"])

                with patch.object(self.api, "open_path", return_value={"status": "success"}) as opener:
                    opened = self.api.open_page_file(linked["fileRef"])
                self.assertEqual("success", opened["status"])
                opener.assert_called_once_with(str(existing_path))

                removed = self.api.delete_page_file(linked["fileRef"])
                self.assertEqual("success", removed["status"])
                self.assertEqual("external", removed["storageType"])
                self.assertFalse(removed["externalFileDeleted"])
                self.assertTrue(existing_path.is_file())

                missing_info = self.api.get_page_file_info(linked["fileRef"])
                self.assertEqual("success", missing_info["status"])
                self.assertFalse(missing_info["exists"])
                self.assertEqual("external", missing_info["storageType"])

    def test_link_existing_workbook_validates_manual_paths(self):
        with tempfile.TemporaryDirectory(prefix="acies-page-file-link-validation-") as temp_dir:
            text_path = Path(temp_dir) / "notes.txt"
            text_path.write_text("not a workbook", encoding="utf-8")
            missing_path = Path(temp_dir) / "missing.xlsx"
            csv_path = Path(temp_dir) / "schedule.csv"
            csv_path.write_text("Item,Value\nA,1\n", encoding="utf-8")
            with self._patched_data_path(temp_dir):
                self.assertEqual("error", self.api.link_page_workbook("relative.xlsx")["status"])
                self.assertEqual("error", self.api.link_page_workbook(str(missing_path))["status"])
                self.assertEqual("error", self.api.link_page_workbook(str(text_path))["status"])
                linked_csv = self.api.link_page_workbook(str(csv_path))
                self.assertEqual("success", linked_csv["status"])
                self.assertEqual("schedule.csv", linked_csv["fileName"])


if __name__ == "__main__":
    unittest.main()
