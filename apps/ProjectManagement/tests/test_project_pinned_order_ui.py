import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectPinnedDeliverablesUiTests(unittest.TestCase):
    def test_list_view_uses_deliverable_pin_state_for_pinned_section(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "isPinnedDeliverable: isDeliverablePinned(deliverable),",
            "const aPinned = !!a?.isPinnedDeliverable;",
            "const bPinned = !!b?.isPinnedDeliverable;",
            "if (row?.isPinnedDeliverable || !shouldSortCompletedProjectsLast()) return 0;",
            "const {\n      project,\n      projectIndex,\n      projectListContext,\n      deliverable,\n      dueDate,\n      isPinnedDeliverable,",
            "if (isPinnedDeliverable) {",
            'appendSectionSeparator("Pinned Deliverables");',
        ):
            self.assertIn(expected, script)

        self.assertNotIn("appendSectionSeparator(\"Pinned Projects\");", script)
        self.assertNotIn("const pinned = getPinnedProjectsInManualOrder(items);", script)
        self.assertNotIn("items = pinned.concat(unpinned);", script)
        self.assertNotIn("isPinnedProject:", script)

    def test_card_view_uses_same_deliverable_rows_as_list_view(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        card_view_start = script.index("function renderCardView(")
        card_view_end = script.index("function setProjectsViewMode(", card_view_start)
        card_view_block = script[card_view_start:card_view_end]
        render_start = script.index("function render() {")
        render_end = script.index("function renderStatusToggles(", render_start)
        render_block = script[render_start:render_end]

        for expected in (
            "function renderCardView(items = db, projectListContextMap = null) {",
            "const deliverableRows = buildProjectDeliverableRowEntries(",
            "sortProjectDeliverableRows(deliverableRows);",
            "for (const { project, deliverable, dueDate, isPinnedDeliverable } of deliverableRows) {",
            "if (pinnedShown && isPinnedDeliverable) {",
            'buckets.get("pinned").push({',
            "const bucketRows = buckets.get(col.key);",
            "sortCardViewBucketRows(col.key, bucketRows);",
        ):
            self.assertIn(expected, card_view_block)

        self.assertIn("renderCardView(items, projectListContextMap);", render_block)

    def test_card_view_sorts_due_dates_youngest_first_by_default(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        helper_start = script.index("function compareProjectDeliverableRowsByDueAsc")
        helper_end = script.index("function clearPinnedProjectDragStyles()", helper_start)
        helper_block = script[helper_start:helper_end]

        for expected in (
            "function compareProjectDeliverableRowsByDueAsc(a, b) {",
            "compareDueDateValues(",
            '"asc"',
            "function sortCardViewBucketRows(columnKey, rows = []) {",
            'columnKey === "pinned"',
            "comparePinnedDeliverableCardRows(a, b)",
            "compareProjectDeliverableRowsByDueAsc(a, b)",
        ):
            self.assertIn(expected, helper_block)

    def test_pinned_deliverables_default_to_due_asc_then_manual_order(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        helper_start = script.index("function normalizePinnedDeliverableOrder")
        helper_end = script.index("function clearPinnedProjectDragStyles()", helper_start)
        helper_block = script[helper_start:helper_end]

        for expected in (
            "function normalizePinnedDeliverableOrder(value) {",
            "function comparePinnedDeliverableCardRows(a, b) {",
            "normalizePinnedDeliverableOrder(",
            "return aPinnedOrder - bPinnedOrder;",
            "return compareProjectDeliverableRowsByDueAsc(a, b);",
            "function movePinnedDeliverableToTarget(",
            "deliverable.pinnedOrder = index;",
        ):
            self.assertIn(expected, helper_block)

    def test_card_view_drop_from_pinned_column_unpins_deliverable_before_status_update(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        drop_start = script.index('host.addEventListener("drop", async (e) => {')
        drop_end = script.index("function render() {", drop_start)
        drop_block = script[drop_start:drop_end]

        self.assertIn(
            'const { deliverableId, sourceColumnKey } = kanbanDragState;',
            drop_block,
        )
        self.assertIn("setDeliverablePinnedState(deliverable, true);", drop_block)
        self.assertIn(
            'if (sourceColumnKey === "pinned" && deliverable.pinned) {',
            drop_block,
        )
        self.assertIn("setDeliverablePinnedState(deliverable, false);", drop_block)
        self.assertLess(
            drop_block.index("setDeliverablePinnedState(deliverable, false);"),
            drop_block.index("setSingleStatus(deliverable, targetKey);"),
        )
        self.assertNotIn("setProjectPinnedState(project", drop_block)

    def test_card_view_drop_replaces_status_and_status_tag_together(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        drop_start = script.index('host.addEventListener("drop", async (e) => {')
        drop_end = script.index("function render() {", drop_start)
        drop_block = script[drop_start:drop_end]

        self.assertIn("setSingleStatus(deliverable, targetKey);", drop_block)
        self.assertNotIn("deliverable.statuses = [targetKey];", drop_block)
        self.assertNotIn("syncStatusArrays(deliverable);", drop_block)

    def test_card_view_drop_reorders_deliverables_within_pinned_column(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        drop_start = script.index('host.addEventListener("drop", async (e) => {')
        drop_end = script.index("function render() {", drop_start)
        drop_block = script[drop_start:drop_end]
        dragover_start = script.index('host.addEventListener("dragover", (e) => {')
        dragover_end = script.index('host.addEventListener("dragleave"', dragover_start)
        dragover_block = script[dragover_start:dragover_end]

        for expected in (
            'kanbanDragState.sourceColumnKey === "pinned"',
            'col.dataset.columnKey === "pinned"',
            'targetCard.dataset.deliverableId !== kanbanDragState.deliverableId',
            "getKanbanCardDropPosition(e, targetCard)",
            "kanban-drop-before",
            "kanban-drop-after",
        ):
            self.assertIn(expected, dragover_block)

        for expected in (
            'if (targetKey === "pinned" && sourceColumnKey === "pinned") {',
            "const moved = movePinnedDeliverableToTarget(",
            "dropPosition !== \"after\"",
            "await save();",
            "renderProjectsPreservingExpandedDeliverables();",
        ):
            self.assertIn(expected, drop_block)

        self.assertLess(
            drop_block.index('if (targetKey === "pinned" && sourceColumnKey === "pinned") {'),
            drop_block.index("if (!targetKey || targetKey === sourceColumnKey) return;"),
        )
        self.assertIn(".kanban-column .deliverable-card-new.kanban-drop-before {", css)
        self.assertIn(".kanban-column .deliverable-card-new.kanban-drop-after {", css)

    def test_project_pin_ui_is_removed(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for removed in (
            'id="pinUrgentDeliverablesBtn"',
            'class="pin-header"',
            'class="pin-btn"',
            'class="cell-select"',
            'aria-label="Pin project"',
        ):
            self.assertNotIn(removed, html)

        for removed in (
            "pinUrgentDeliverablesBtn",
            'tr.classList.toggle("is-pinned-project", isPinned);',
            "if (isPinned) enablePinnedProjectRowDrag(tr, pinBtn, project);",
            "setProjectPinnedState(project, !project?.pinned, db);",
        ):
            self.assertNotIn(removed, script)

        for removed in (
            ".pin-btn {",
            ".pin-header {",
            ".table th.cell-select {",
            ".row td.cell-select {",
            ".project-row.is-pinned-project .pin-btn {",
        ):
            self.assertNotIn(removed, css)


if __name__ == "__main__":
    unittest.main()
