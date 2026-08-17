import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
REQUIREMENTS_PATH = REPO_ROOT / "requirements.txt"
MAIN_PY_PATH = REPO_ROOT / "main.py"
ROOT_SPEC_PATH = REPO_ROOT / "ACIES Scheduler.spec"
BUILD_SPEC_PATH = REPO_ROOT / "build-config" / "ACIES Scheduler.spec"
BUILD_SCRIPT_PATH = REPO_ROOT / "build-config" / "build.ps1"


class HeicSupportConfigTests(unittest.TestCase):
    def test_requirements_include_pillow_heif(self):
        text = REQUIREMENTS_PATH.read_text(encoding="utf-8")
        self.assertIn("pillow-heif", text)

    def test_pyinstaller_specs_bundle_heif_modules(self):
        root_spec = ROOT_SPEC_PATH.read_text(encoding="utf-8")
        build_spec = BUILD_SPEC_PATH.read_text(encoding="utf-8")

        for text in (root_spec, build_spec):
            self.assertIn("pillow_heif", text)
            self.assertIn("_pillow_heif", text)

    def test_build_script_uses_checked_in_spec_and_repairs_heif_dependencies(self):
        text = BUILD_SCRIPT_PATH.read_text(encoding="utf-8")

        self.assertIn("build-config\\ACIES Scheduler.spec", text)
        self.assertIn("pip install -r requirements.txt", text)
        self.assertIn("pillow_heif", text)
        self.assertIn("_pillow_heif", text)
        self.assertIn('Join-Path $bundleInternal "pillow_heif"', text)
        self.assertIn('_pillow_heif*.pyd', text)
        self.assertIn("archive_viewer", text)

    def test_expense_image_file_dialog_accepts_heic_and_heif(self):
        text = MAIN_PY_PATH.read_text(encoding="utf-8")
        self.assertIn(
            "Image Files (*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.heic;*.heif)",
            text,
        )


if __name__ == "__main__":
    unittest.main()
