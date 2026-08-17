import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"


class PanelScheduleAiMultiPhotoUiTests(unittest.TestCase):
    def test_panel_schedule_markup_mentions_multi_photo_guidance(self):
        text = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn(
            "Upload one or more breaker photos and one or more directory photos",
            text,
        )
        self.assertIn(
            "upper half, middle, and bottom half photos.",
            text,
        )
        self.assertIn(
            "circuits 1-42 on one image and 43-84 on",
            text,
        )
        self.assertIn("Drag &amp; drop photos or click to select", text)
        self.assertIn("No photos", text)
        self.assertIn("JPG, PNG, HEIC, HEIF supported.", text)
        self.assertIn("pasted screenshots append", text)
        self.assertIn("Progress and results appear in the activity tray.", text)
        self.assertIn("Use the tray action to open the output folder after completion.", text)

    def test_panel_schedule_script_uses_array_backed_photo_state_and_plural_payloads(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("breakerPaths: [],", text)
        self.assertIn("directoryPaths: [],", text)
        self.assertIn("breakerFiles: [],", text)
        self.assertIn("directoryFiles: [],", text)
        self.assertIn("launchContext: null,", text)
        self.assertIn("function setCircuitBreakerFiles(kind, files) {", text)
        self.assertIn("function setCircuitBreakerPaths(kind, paths) {", text)
        self.assertIn("function getCircuitBreakerLaunchDefaultDirectory() {", text)
        self.assertIn("allow_multiple: true,", text)
        self.assertIn("Image Files (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff;*.heic;*.heif)", text)
        self.assertIn("default_directory: defaultDirectory,", text)
        self.assertIn("input.multiple = true;", text)
        self.assertIn("const files = Array.from(e.dataTransfer?.files || []);", text)
        self.assertIn("if (files.length) setCircuitBreakerFiles(kind, files);", text)
        self.assertIn('void selectCircuitBreakerImage("breaker");', text)
        self.assertIn('void selectCircuitBreakerImage("directory");', text)
        self.assertIn("defaultDirectory || null,", text)
        self.assertIn(
            "breakerPaths: [...normalizeCircuitBreakerPaths(panel.breakerPaths)],",
            text,
        )
        self.assertIn(
            "directoryPaths: [...normalizeCircuitBreakerPaths(panel.directoryPaths)],",
            text,
        )
        self.assertIn("breakerPaths: firstPanel.breakerPaths || [],", text)
        self.assertIn("directoryPaths: firstPanel.directoryPaths || [],", text)
        self.assertIn("function filesToUploadPayloads(files) {", text)
        self.assertIn("activeJobId: \"\",", text)
        self.assertIn("panelSchedulePollTimer: 0,", text)
        self.assertIn("lastHandledTerminalJobId: \"\",", text)
        self.assertIn("lastPanelScheduleStatus: null,", text)
        self.assertIn("pasteTargetKind: \"\",", text)
        self.assertIn("activeActivityId: \"\",", text)
        self.assertIn("function getCircuitBreakerClipboardImageFiles(clipboardData) {", text)
        self.assertIn("function resolveCircuitBreakerPasteTargetKind(eventTarget = null) {", text)
        self.assertIn("function appendCircuitBreakerFiles(kind, files) {", text)
        self.assertIn("...normalizeCircuitBreakerFiles(panel.breakerFiles),", text)
        self.assertIn("...normalizeCircuitBreakerFiles(panel.directoryFiles),", text)
        self.assertIn("function handleCircuitBreakerPaste(e) {", text)
        self.assertIn('circuitBreakerDlg.addEventListener("paste", handleCircuitBreakerPaste);', text)
        self.assertIn("appendCircuitBreakerFiles(targetKind, imageFiles);", text)
        self.assertIn("breakerUploads,", text)
        self.assertIn("directoryUploads,", text)
        self.assertIn("function schedulePanelScheduleStatusPoll(jobId, delay = 1000) {", text)
        self.assertIn("window.pywebview?.api?.get_panel_schedule_background_status", text)
        self.assertIn("window.updateActivityStatus({", text)
        self.assertIn('"toolCircuitBreaker"', text)
        self.assertIn("function handlePanelScheduleBackgroundUpdate(payload, { source = \"push\" } = {}) {", text)
        self.assertIn("window.handlePanelScheduleResult = async function (payload) {", text)
        self.assertIn("await handlePanelScheduleBackgroundUpdate(payload, { source: \"push\" });", text)
        self.assertIn('payload.activityId = activityId;', text)
        self.assertIn('label: "Panel Schedule AI"', text)
        self.assertIn("completedCount,", text)
        self.assertIn("circuitBreakerState.lastHandledTerminalJobId === jobId", text)
        self.assertNotIn("await window.pywebview.api.open_path(folder);", text)


if __name__ == "__main__":
    unittest.main()
