import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class OutlookScanUiTests(unittest.TestCase):
    def test_email_intake_markup_exists_and_scan_markup_is_removed(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('id="outlookScanBtn"', html)
        self.assertIn('id="outlookScanDlg"', html)
        self.assertIn("Email Intake", html)
        self.assertIn('id="emailIntakePastePanel"', html)
        self.assertIn('id="emailArea"', html)
        self.assertIn('id="btnProcessEmail"', html)

        self.assertNotIn("Scan Outlook Day", html)
        self.assertNotIn('id="aiSpinner"', html)
        self.assertNotIn('id="emailIntakeModeGroup"', html)
        self.assertNotIn('id="emailIntakeModePaste"', html)
        self.assertNotIn('id="emailIntakeModeScan"', html)
        self.assertNotIn('id="emailIntakeScanPanel"', html)
        self.assertNotIn('id="outlookScanCapabilityStatus"', html)
        self.assertNotIn('id="outlookScanCapabilityDetails"', html)
        self.assertNotIn('id="outlookScanDate"', html)
        self.assertNotIn('id="outlookScanRunBtn"', html)
        self.assertNotIn('id="outlookScanProgress"', html)
        self.assertNotIn('id="outlookScanReport"', html)
        self.assertNotIn('id="outlookScanSuggestions"', html)
        self.assertNotIn('id="outlookScanSkipped"', html)
        self.assertNotIn('id="aiBtn"', html)
        self.assertNotIn('id="emailDlg"', html)
        self.assertNotIn('id="outlookScanConnectBtn"', html)
        self.assertNotIn('id="outlookScanSignOutBtn"', html)
        self.assertNotIn('id="settings_microsoftAuthStatus"', html)
        self.assertNotIn('id="settings_microsoftAuthActionBtn"', html)

    def test_email_intake_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("const DEFAULT_OUTLOOK_SCAN_CAPABILITY = {", script)
        self.assertIn('mode: "paste"', script)
        self.assertIn('function normalizeEmailIntakeMode(value = "") {', script)
        self.assertIn("function renderOutlookScanUi() {", script)
        self.assertIn('async function openOutlookScanDialog(mode = "paste") {', script)
        self.assertIn("function beginEmailIntakeActivity() {", script)
        self.assertIn("function completeEmailIntakeActivity(", script)
        self.assertIn("function failEmailIntakeActivity(", script)
        self.assertIn('label: "Email Intake"', script)
        self.assertIn('message: "Processing with AI..."', script)
        self.assertIn("function buildEmailIntakeProjectContext() {", script)
        self.assertIn("async function processEmailIntakePaste() {", script)
        self.assertIn("const projectContext = buildEmailIntakeProjectContext();", script)
        self.assertIn("getActiveDisciplineList(),\n      projectContext", script)
        self.assertIn('closeDlg("outlookScanDlg");', script)
        self.assertIn("handleAiProjectResult(res.data || {});", script)
        self.assertIn('outlookScanBtn.onclick = () => openOutlookScanDialog();', script)
        self.assertIn('btnProcessEmail.onclick = () => processEmailIntakePaste();', script)

        self.assertNotIn('document.getElementById("emailIntakeModeScan")', script)
        self.assertNotIn('outlookScanRunBtn.onclick = () => runOutlookInboxScan();', script)
        self.assertNotIn("DEFAULT_MICROSOFT_AUTH_STATE", script)
        self.assertNotIn("renderMicrosoftAuthUi", script)
        self.assertNotIn("loadMicrosoftAuthState", script)
        self.assertNotIn("handleMicrosoftSignIn", script)
        self.assertNotIn("handleMicrosoftSignOut", script)
        self.assertNotIn('closeDlg("emailDlg")', script)
        self.assertNotIn("No summary available.", script)

    def test_email_intake_project_context_builder_is_compact_identity_only(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("const EMAIL_INTAKE_PROJECT_CONTEXT_MAX_PROJECTS = 200;", script)
        self.assertIn("const EMAIL_INTAKE_PROJECT_CONTEXT_MAX_CHARS = 25000;", script)

        start = script.index("function buildEmailIntakeProjectContext() {")
        end = script.index("async function processEmailIntakePaste()", start)
        body = script[start:end]

        self.assertIn('id: String(project.id || "").trim(),', body)
        self.assertIn('name: String(project.name || "").trim(),', body)
        self.assertIn('nick: String(project.nick || "").trim(),', body)
        self.assertIn("pathLeaf: getWindowsPathLeaf(path),", body)
        self.assertIn("JSON.stringify(candidate)", body)
        self.assertNotIn("deliverables", body)
        self.assertNotIn("tasks", body)

    def test_email_intake_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".email-intake-panel[hidden] {", css)
        self.assertIn(".email-intake-panel-head {", css)

        self.assertNotIn(".email-intake-mode-switch {", css)
        self.assertNotIn(".email-intake-mode-option {", css)
        self.assertNotIn(".email-intake-mode-label {", css)
        self.assertNotIn(".email-intake-scan-status {", css)
        self.assertNotIn(".outlook-scan-toolbar {", css)
        self.assertNotIn(".outlook-scan-results {", css)
        self.assertNotIn(".outlook-scan-report summary {", css)
        self.assertNotIn(".outlook-scan-skipped-item {", css)


if __name__ == "__main__":
    unittest.main()
