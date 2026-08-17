import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ActivityTrayUiTests(unittest.TestCase):
    def test_activity_tray_markup_exists(self):
        text = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('id="activityTray"', text)
        self.assertIn('id="activityTrayToggle"', text)
        self.assertIn('id="activityTrayCounts"', text)
        self.assertIn('id="activityTrayBody"', text)
        self.assertIn('id="activityTrayEmpty"', text)
        self.assertIn('id="activityTrayList"', text)
        self.assertIn('id="activityTrayClearAll"', text)
        self.assertIn('class="activity-tray-header"', text)
        self.assertIn("Activity", text)
        self.assertIn("No activity yet.", text)
        self.assertIn("Clear All", text)

    def test_activity_tray_styles_exist(self):
        text = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".activity-tray", text)
        self.assertIn(".activity-tray.is-collapsed .activity-tray-body", text)
        self.assertIn(".activity-card", text)
        self.assertIn(".activity-card-progress-bar", text)
        self.assertIn(".activity-card-action.accept", text)
        self.assertIn(".activity-card-action.rerun", text)
        self.assertIn(".activity-tray-clear", text)

    def test_activity_tray_script_helpers_exist(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("const activityTrayState = {", text)
        self.assertIn("function initActivityTray() {", text)
        self.assertIn("function beginActivity({", text)
        self.assertIn("function updateActivity(activityId, patch = {}) {", text)
        self.assertIn("function completeActivity(activityId, patch = {}) {", text)
        self.assertIn("function failActivity(activityId, patch = {}) {", text)
        self.assertIn("function acceptActivity(activityId) {", text)
        self.assertIn("function clearAllActivityNotifications() {", text)
        self.assertIn("function handleActivityTrayClearAll() {", text)
        self.assertIn("function handleActivityTrayRerun(activityId) {", text)
        self.assertIn("async function handleActivityTrayOpenFolder(activityId) {", text)
        self.assertIn("async function handleActivityTrayCopyCombinedPdf(activityId) {", text)
        self.assertIn("async function handleActivityTrayOpenCombinedPdf(activityId) {", text)
        self.assertIn("function handleActivityTrayAccept(activityId) {", text)
        self.assertIn("function renderActivityTray() {", text)
        self.assertIn("ACTIVITY_RERUN_TOOL_IDS", text)
        self.assertIn("rerunDefaultPath", text)
        self.assertIn("rerunLaunchContext", text)
        self.assertIn("combinedPdfPath", text)
        self.assertIn('textContent: "Copy Combined PDF"', text)
        self.assertIn('textContent: "Open Combined PDF"', text)
        self.assertIn('textContent: "Rerun"', text)
        self.assertIn('textContent: "Accept"', text)
        self.assertIn('"data-activity-action": "copy-combined-pdf"', text)
        self.assertIn('"data-activity-action": "open-combined-pdf"', text)
        self.assertIn('"data-activity-action": "open"', text)
        self.assertIn('"data-activity-action": "rerun"', text)
        self.assertIn('await handleActivityTrayRerun(activityId);', text)
        self.assertIn('rawMessage.startsWith("INPUT_FOLDER:")', text)
        self.assertIn(
            'const result = await window.pywebview.api.open_path(activity.openFolderPath);',
            text,
        )
        self.assertIn(
            "window.pywebview.api.copy_file_to_clipboard(",
            text,
        )
        self.assertIn(
            "const result = await window.pywebview.api.open_path(activity.combinedPdfPath);",
            text,
        )
        self.assertIn(
            'if (result && String(result.status || "").trim().toLowerCase() !== "success") {',
            text,
        )

    def test_workflow_activity_title_is_rendered(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("function getWorkflowDisplayName(workflow) {", text)
        self.assertIn("function getWorkflowActivityLabel(workflow) {", text)
        self.assertIn("function getActivityLabelFromPayload(toolId, payload = {}, existing = null) {", text)
        self.assertIn('normalizedToolId === "toolWorkflow" && workflowTitle', text)
        self.assertIn("const activityTitle =", text)
        self.assertIn("textContent: activityTitle", text)


if __name__ == "__main__":
    unittest.main()
