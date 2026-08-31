import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class CanvasPanelScheduleUiTests(unittest.TestCase):
    """Canvas selection -> Panel Schedule AI confirmation dialog."""

    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_confirmation_dialog_markup_exists(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        for expected in (
            'id="canvasPanelScheduleDlg"',
            'id="cpsLoading"',
            'id="cpsError"',
            'id="cpsContent"',
            'id="cpsPanelName"',
            'id="cpsInputModeField"',
            'id="cpsInputModeExisting"',
            'id="cpsImageList"',
            'id="cpsOutputPath"',
            'id="cpsWarning"',
            'id="cpsConfirmBtn"',
            'id="cpsCancelBtn"',
        ):
            self.assertIn(expected, html)

        # Sits between the Panel Schedule AI modal and the lighting scheduler.
        self.assertLess(
            html.index('id="circuitBreakerDlg"'), html.index('id="canvasPanelScheduleDlg"')
        )
        self.assertLess(
            html.index('id="canvasPanelScheduleDlg"'), html.index('id="lightingScheduleDlg"')
        )

    def test_dialog_state_and_render_functions_exist(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "const canvasPanelScheduleState = {",
            "function initCanvasPanelScheduleDialog() {",
            "async function openCanvasPanelScheduleDialog() {",
            "function renderCanvasPanelScheduleDialog(",
            "function getCanvasPanelScheduleValidation() {",
            "async function refreshCanvasPanelScheduleOutputPath() {",
            "async function confirmCanvasPanelSchedule() {",
            "function showCanvasPanelScheduleFindings(result) {",
            "window.pywebview.api.resolve_canvas_panel_schedule_output(",
        ):
            self.assertIn(expected, script)

    def test_open_guards_against_an_already_running_job(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        open_block = self._block(
            script,
            "async function openCanvasPanelScheduleDialog() {",
            "function handleCanvasPanelClassificationRecord(record) {",
        )
        self.assertIn("if (circuitBreakerState.activeJobId) {", open_block)
        self.assertIn('toast("Panel Schedule AI is already running.");', open_block)
        self.assertIn("getSelectedCanvasTextNodes()", open_block)

    def test_analysis_runs_in_the_background_without_blocking_the_dialog(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        open_block = self._block(
            script,
            "async function openCanvasPanelScheduleDialog() {",
            "function handleCanvasPanelClassificationRecord(record) {",
        )
        # The menu handler starts a job and reports through the tray; it never awaits
        # the AI with a modal up.
        self.assertIn(
            "window.pywebview.api.run_canvas_panel_classification_background(",
            open_block,
        )
        self.assertIn("beginActivity({", open_block)
        self.assertIn("scheduleCanvasPanelClassificationPoll(", open_block)
        self.assertNotIn("classify_canvas_panel_selection(", open_block)

        for expected in (
            "function scheduleCanvasPanelClassificationPoll(jobId, delay = 1500) {",
            "async function pollCanvasPanelClassificationStatus(jobId) {",
            "window.pywebview.api.get_canvas_panel_classification_status(jobId)",
            "window.handleCanvasPanelClassificationResult = function (record) {",
        ):
            self.assertIn(expected, script)

    def test_findings_auto_open_only_on_the_originating_page(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        record_block = self._block(
            script,
            "function handleCanvasPanelClassificationRecord(record) {",
            "window.handleCanvasPanelClassificationResult = function (record) {",
        )
        self.assertIn("getCanvasPanelSchedulePageKey() === job.pageKey", record_block)
        self.assertIn('document.querySelector("dialog[open]")', record_block)
        self.assertIn("canvasPanelReview: !showNow,", record_block)
        self.assertIn("if (showNow) showCanvasPanelScheduleFindings(result);", record_block)
        # Both the push callback and the poll land here; the job handle is cleared first
        # so the second delivery is a no-op.
        self.assertIn("canvasPanelScheduleState.job = null;", record_block)

    def test_tray_offers_a_review_action_for_deferred_findings(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function handleActivityTrayReviewCanvasPanel(activityId) {",
            'data-activity-action": "review-canvas-panel",',
            'textContent: "Review findings",',
            'if (action === "review-canvas-panel") {',
            # upsertActivity whitelists fields, so the flag has to be declared there.
            "canvasPanelReview: Boolean(",
        ):
            self.assertIn(expected, script)

    def test_validation_mirrors_the_backend_photo_requirements(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        validation_block = self._block(
            script,
            "function getCanvasPanelScheduleValidation() {",
            "function renderCanvasPanelScheduleRow(image) {",
        )
        self.assertIn(
            'if (state.inputMode === "existing_directory" && directories === 0) {',
            validation_block,
        )
        self.assertIn(
            'if (state.inputMode === "field_photos" && (breakers === 0 || directories === 0)) {',
            validation_block,
        )
        # Both document kinds count as a circuit list.
        self.assertIn("isCanvasPanelDirectoryRole(image.role)", validation_block)

    def test_as_built_and_field_directory_are_separate_roles(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            'as_built_schedule: "As-built schedule (printed)",',
            'field_directory: "Field directory card",',
            'panel_label: "Panel label / nameplate",',
            'main_breaker: "Main breaker photo",',
            'const CANVAS_PANEL_DIRECTORY_ROLES = ["as_built_schedule", "field_directory"];',
            "function isCanvasPanelDirectoryRole(role) {",
            "function deriveCanvasPanelInputMode(images) {",
        ):
            self.assertIn(expected, script)

        derive_block = self._block(
            script,
            "function deriveCanvasPanelInputMode(images) {",
            "const canvasPanelScheduleState = {",
        )
        # A printed as-built outranks breaker photos, matching main.py.
        self.assertIn('if (asBuilts > 0) return "existing_directory";', derive_block)
        self.assertIn('if (breakers === 0) return "existing_directory";', derive_block)
        self.assertIn('return "field_photos";', derive_block)

        confirm_block = self._block(
            script,
            "async function confirmCanvasPanelSchedule() {",
            "const debouncedSaveLightingSchedule",
        )
        self.assertIn('image.role === "panel_label"', confirm_block)
        self.assertIn('image.role === "main_breaker"', confirm_block)
        self.assertIn("panelLabelPaths,", confirm_block)
        self.assertIn("mainBreakerPaths,", confirm_block)

    def test_correcting_a_role_moves_the_mode_unless_overridden(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        init_block = self._block(
            script,
            "function initCanvasPanelScheduleDialog() {",
            "function closeCanvasPanelScheduleDialog() {",
        )
        self.assertIn("canvasPanelScheduleState.modeTouched = true;", init_block)
        self.assertIn("if (!canvasPanelScheduleState.modeTouched) {", init_block)
        self.assertIn("deriveCanvasPanelInputMode(", init_block)

    def test_existing_directory_mode_drops_breaker_paths(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        confirm_block = self._block(
            script,
            "async function confirmCanvasPanelSchedule() {",
            "const debouncedSaveLightingSchedule",
        )
        self.assertIn('state.inputMode === "existing_directory"', confirm_block)
        self.assertIn("? []", confirm_block)
        self.assertIn('outputMode: "new",', confirm_block)
        self.assertIn('panelId: "canvas_panel_1",', confirm_block)
        self.assertIn("revealOnComplete: true,", confirm_block)

    def test_completion_reveals_the_new_file(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn('revealOnCompleteJobId: "",', script)
        self.assertIn("async function revealPanelScheduleOutput(targetPath) {", script)
        self.assertIn("window.pywebview.api.reveal_path(path)", script)

        update_block = self._block(
            script,
            "async function handlePanelScheduleBackgroundUpdate(",
            "async function runCircuitBreakerInBackground() {",
        )
        self.assertIn(
            "const shouldReveal = circuitBreakerState.revealOnCompleteJobId === jobId;",
            update_block,
        )
        self.assertIn("if (shouldReveal && successCount > 0) {", update_block)

    def test_submit_helper_is_shared_with_the_modal_flow(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("async function submitPanelScheduleJob({", script)
        self.assertIn(
            "circuitBreakerState.revealOnCompleteJobId = revealOnComplete ? jobId : \"\";",
            script,
        )

        run_block = self._block(
            script,
            "async function runCircuitBreakerInBackground() {",
            "async function submitPanelScheduleJob({",
        )
        self.assertIn("submitPanelScheduleJob({", run_block)
        self.assertIn("breakerPaths: firstPanel.breakerPaths || [],", run_block)
        self.assertIn("closeCircuitBreaker();", run_block)

    def test_canvas_panel_schedule_dialog_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".canvas-panel-schedule-dialog {",
            ".cps-shell {",
            ".cps-image-list {",
            ".cps-image-row {",
            ".cps-image-row.is-unavailable {",
            ".cps-image-thumb {",
            ".cps-warning {",
            ".cps-error {",
            ".cps-footer {",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
