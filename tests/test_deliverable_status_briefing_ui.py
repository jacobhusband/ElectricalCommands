import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class DeliverableStatusBriefingUiTests(unittest.TestCase):
    """Incomplete deliverables export, revived from the orphaned notepad dialog."""

    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_projects_toolbar_exposes_the_deliverable_export_button(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        block = self._block(
            html, '<div class="toolbar-actions">', '<div class="projects-view-toggle"'
        )
        self.assertIn('id="exportDeliverablesBtn"', block)
        self.assertIn('title="Export incomplete deliverables to Excel"', block)
        self.assertIn('aria-label="Export deliverables"', block)
        # It sits next to the stats button, not buried elsewhere in the toolbar.
        self.assertLess(
            block.index('id="statsBtn"'), block.index('id="exportDeliverablesBtn"')
        )

    def test_dialog_has_scope_chips_and_a_summary_panel(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        block = self._block(
            html,
            '<dialog id="deliverableNotepadDlg"',
            '<dialog id="copyProjectLocallyDlg"',
        )
        self.assertIn('id="deliverableNotepadScopeIncompleteBtn"', block)
        self.assertIn('data-notepad-scope="incomplete"', block)
        self.assertIn('id="deliverableNotepadScopeAllBtn"', block)
        self.assertIn('data-notepad-scope="all"', block)
        self.assertIn('class="projects-filter-chip is-active"', block)
        self.assertIn('id="deliverableNotepadSummaryBtn"', block)
        self.assertIn('id="deliverableNotepadSummaryPanel"', block)
        self.assertIn('id="deliverableNotepadSummaryBody"', block)
        self.assertIn('id="deliverableNotepadSummaryMeta"', block)
        self.assertIn('aria-live="polite"', block)
        self.assertIn('id="deliverableNotepadAvailablePanelTitle"', block)
        # Scope is a chip group, not a dropdown.
        self.assertNotIn("<select", block)
        # The briefing reads above the export button, not below it.
        self.assertLess(
            block.index("deliverableNotepadSummaryPanel"),
            block.index('class="modal-footer"'),
        )

    def test_opener_preselects_every_incomplete_deliverable(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        block = self._block(
            script,
            "function openDeliverablesExcelDialog() {",
            "function setDeliverableNotepadScope(",
        )
        self.assertIn("buildDeliverableNotepadEntries(\n    db,\n    deliverableNotepadScope\n  )", block)
        self.assertIn(
            "deliverableNotepadSelectedEntryIds = deliverableNotepadEntries.map(",
            block,
        )
        self.assertIn("dialog.showModal();", block)
        # An all-finished database falls back to the full list instead of dead-ending.
        self.assertIn('deliverableNotepadScope = "all";', block)

    def test_entry_builder_filters_by_scope(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn(
            'function buildDeliverableNotepadEntries(projects = db, scope = "incomplete") {',
            script,
        )
        block = self._block(
            script,
            "function buildDeliverableNotepadEntries(projects = db",
            "function toDeliverableSummaryIsoDate(",
        )
        self.assertIn('scope === "all" || !isFinished(deliverable)', block)
        self.assertIn("dueBucket: getDeliverableSummaryBucket(deliverable),", block)
        self.assertIn('projectPath: String(project?.path || "").trim(),', block)

    def test_project_path_reaches_the_excel_payload(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        block = self._block(
            script,
            "function buildSelectedDeliverablesExcelRows(entries = []) {",
            "function createDeliverableNotepadListItem(item) {",
        )
        self.assertIn('projectPath: String(item?.projectPath || "").trim(),', block)

    def test_summary_bucket_checks_undated_before_due_state(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        block = self._block(
            script,
            "function getDeliverableSummaryBucket(deliverable) {",
            "const DELIVERABLE_SUMMARY_BUCKET_ORDER",
        )
        # dueState() returns "ok" for undated items, so this ordering matters.
        self.assertLess(
            block.index('return "undated"'),
            block.index("deliverableDueState(deliverable)"),
        )
        self.assertIn('if (state === "critical") return "missedHardDeadline";', block)

    def test_summary_request_caps_rows_and_reports_omissions(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("const DELIVERABLE_SUMMARY_MAX_ROWS = 150;", script)
        block = self._block(
            script,
            "function buildDeliverableSummaryRequest(selectedItems = []) {",
            "function buildSelectedDeliverablesExcelRows(",
        )
        self.assertIn("let remaining = DELIVERABLE_SUMMARY_MAX_ROWS;", block)
        self.assertIn("omittedCount += all.length - take;", block)
        self.assertIn("compareDeliverablesByDue(a?.deliverable, b?.deliverable)", block)

    def test_export_attaches_the_summary_only_once_generated(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        block = self._block(
            script,
            "async function exportSelectedDeliverablesToExcel() {",
            "function clearDeliverableNotepadSummary() {",
        )
        self.assertIn("const payload = { entries };", block)
        self.assertIn("if (deliverableNotepadSummary) {", block)
        self.assertIn("payload.summary = {", block)
        self.assertIn("export_deliverables_excel(payload)", block)

    def test_briefing_failure_never_blocks_the_export(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        block = self._block(
            script,
            "async function generateDeliverableStatusBriefing() {",
            "function compareDeliverablesByDue(a, b) {",
        )
        self.assertIn("} finally {", block)
        self.assertIn("deliverableNotepadSummaryLoading = false;", block)
        self.assertLess(block.index("} finally {"), block.rindex("deliverableNotepadSummaryLoading = false;"))
        # The AI path must not touch the export button's disabled state.
        self.assertNotIn("deliverableNotepadExportBtn", block)
        self.assertIn("generate_deliverable_status_summary", block)

    def test_summary_is_cleared_when_the_selection_changes(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        add_block = self._block(
            script,
            "function addDeliverablesToNotepadSelection() {",
            "function removeDeliverablesFromNotepadSelection() {",
        )
        remove_block = self._block(
            script,
            "function removeDeliverablesFromNotepadSelection() {",
            "async function exportSelectedDeliverablesToExcel() {",
        )
        self.assertIn("clearDeliverableNotepadSummary();", add_block)
        self.assertIn("clearDeliverableNotepadSummary();", remove_block)

    def test_dialog_render_syncs_scope_chips_and_panel_title(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        block = self._block(
            script,
            "function renderDeliverableNotepadDialog() {",
            "function renderDeliverableNotepadSummary() {",
        )
        self.assertIn('"All incomplete deliverables are already selected for export."', block)
        self.assertIn('deliverableNotepadAvailablePanelTitle', block)
        self.assertIn('chip.classList.toggle("is-active", isActive);', block)
        self.assertIn('chip.setAttribute("aria-pressed", isActive ? "true" : "false");', block)
        self.assertIn("renderDeliverableNotepadSummary();", block)

    def test_notepad_summary_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        self.assertIn(".deliverable-notepad-toolbar {", css)
        self.assertIn(".deliverable-notepad-summary {", css)
        self.assertIn(".deliverable-notepad-summary-header {", css)
        self.assertIn(".deliverable-notepad-summary-body p {", css)
        self.assertIn(".deliverable-notepad-summary-headline {", css)
        self.assertIn('[data-theme="light"] .deliverable-notepad-summary {', css)

    def test_buttons_are_wired(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn(
            "exportDeliverablesBtn.onclick = () => openDeliverablesExcelDialog();",
            script,
        )
        self.assertIn(
            "deliverableNotepadSummaryBtn.onclick = () =>\n      generateDeliverableStatusBriefing();",
            script,
        )
        self.assertIn(
            "setDeliverableNotepadScope(chip.dataset.notepadScope)",
            script,
        )


if __name__ == "__main__":
    unittest.main()
