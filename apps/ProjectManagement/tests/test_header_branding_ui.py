import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SPEC_PATH = REPO_ROOT / "ACIES Scheduler.spec"


class HeaderBrandingUiTests(unittest.TestCase):
    def test_header_and_loader_logo_use_assets_path(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertEqual(2, html.count('src="./assets/acies.png"'))

    def test_version_chip_is_in_header_actions_before_update_button(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        header_start_index = html.index('<div class="header-start">')
        header_actions_index = html.index('<div class="header-actions">')
        version_chip_index = html.index('id="versionChip"')
        app_update_index = html.index('id="appUpdateBtn"')

        header_start_section = html[header_start_index:header_actions_index]
        header_actions_section = html[header_actions_index:app_update_index]

        self.assertGreater(version_chip_index, header_actions_index)
        self.assertLess(version_chip_index, app_update_index)
        self.assertNotIn('id="versionChip"', header_start_section)
        self.assertIn('id="versionChip"', header_actions_section)

    def test_spec_packages_logo_inside_assets_folder(self):
        spec = SPEC_PATH.read_text(encoding="utf-8")

        self.assertIn("('assets\\\\acies.png', 'assets')", spec)


if __name__ == "__main__":
    unittest.main()
