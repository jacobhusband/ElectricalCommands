import pathlib
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]


class EmailCaptureUiRemovalTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.html = (ROOT / "index.html").read_text(encoding="utf-8")
        cls.script = (ROOT / "script.js").read_text(encoding="utf-8")
        cls.styles = (ROOT / "styles.css").read_text(encoding="utf-8")

    def test_email_tab_and_panel_are_removed(self):
        for fragment in (
            'data-tab="email"',
            'id="emailTabBadge"',
            'id="email-panel"',
            'id="emailInboxFilters"',
            'id="emailCaptureRescanBtn"',
            'id="emailCaptureProgress"',
            'id="emailInboxList"',
            'id="followUpsSection"',
            'id="followUpsList"',
            'id="addFollowUpBtn"',
        ):
            self.assertNotIn(fragment, self.html)

    def test_email_tab_dialogs_are_removed(self):
        for fragment in (
            '<dialog id="followUpDlg">',
            'id="followUpTitleInput"',
            'id="followUpWaitingOnInput"',
            'id="followUpDueInput"',
            'id="followUpSaveBtn"',
            '<dialog id="emailProjectPickerDlg">',
            'id="emailProjectPickerSearch"',
            'id="emailProjectPickerList"',
        ):
            self.assertNotIn(fragment, self.html)

    def test_email_tab_script_and_styles_are_removed(self):
        for fragment in (
            "let emailCaptureState = {",
            "window.updateEmailCaptureProgress = function",
            "function renderEmailInboxView() {",
            "function renderFollowUpsView() {",
            "function initEmailCaptureUi() {",
            "function startEmailCaptureAutoScan() {",
            'tab === "email"',
            "run_email_capture_scan",
            "list_captured_emails",
            "get_email_capture_badge_counts",
        ):
            self.assertNotIn(fragment, self.script)
        for selector in (
            ".tab-badge",
            ".email-inbox-card",
            ".follow-up-card",
            ".email-directedness-pill",
        ):
            self.assertNotIn(selector, self.styles)

    def test_manual_email_intake_remains_available(self):
        self.assertIn('id="outlookScanBtn"', self.html)
        self.assertIn('<dialog id="outlookScanDlg">', self.html)
        self.assertIn('id="emailIntakePastePanel"', self.html)

    def test_legacy_outlook_scan_ids_remain_absent(self):
        for legacy_id in (
            'id="outlookScanRunBtn"',
            'id="outlookScanProgress"',
            'id="outlookScanReport"',
            'id="outlookScanSuggestions"',
            'id="emailIntakeScanPanel"',
            'id="outlookScanDate"',
        ):
            self.assertNotIn(legacy_id, self.html)
        self.assertNotIn("Scan Outlook Day", self.html)


if __name__ == "__main__":
    unittest.main()
