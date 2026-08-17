import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverablePinnedWorkItemsUiTests(unittest.TestCase):
    @staticmethod
    def _css_block(css: str, selector: str) -> str:
        block_start = css.index(selector)
        block_end = css.index("}", block_start)
        return css[block_start:block_end]

    @staticmethod
    def _script_block(script: str, start_marker: str, end_marker: str) -> str:
        block_start = script.index(start_marker)
        block_end = script.index(end_marker, block_start)
        return script[block_start:block_end]

    def test_projects_card_status_and_pinned_wiring_exists_without_notes_ui(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertEqual(script.count("function renderDeliverableCard("), 1)
        self.assertEqual(script.count("function createTasksPreview("), 1)
        self.assertNotIn("function createNotesSection(", script)
        self.assertNotIn("function createNotesSectionLegacy(", script)

        for expected in (
            "function getPinnedDeliverablePreviewItems(deliverable) {",
            "function createWorkItemPinButton({",
            "function renderDeliverableStatusBadges(container, deliverable) {",
            "function createDeliverableStatusSection(deliverable, project, card) {",
            "function renderDeliverablePinnedPreview(container, deliverable) {",
            "function updateDeliverableWorkItemUi(card, deliverable) {",
            "function renderDeliverableCard(deliverable, isPrimary, project) {",
            'className: "deliverable-status-inline-group",',
            'className: "deliverable-pinned-inline-group"',
            'pillIcon.classList.add("deliverable-pinned-inline-pill-icon");',
            'className: `deliverable-pinned-inline-pill ${',
            'className: "deliverable-pinned-inline-pill-text"',
            "renderProjectsPreservingExpandedDeliverables();",
            'card.dataset.deliverableId = deliverableId;',
        ):
            self.assertIn(expected, script)

        current_card_block = self._script_block(
            script,
            "function renderDeliverableCard(deliverable, isPrimary, project) {",
            "function normalizeProjectMatchValue(value) {",
        )
        self.assertIn(
            "const actionRow = createDeliverableCardTopActions(deliverable, project, card);",
            current_card_block,
        )
        self.assertIn("const statusSection = createDeliverableStatusSection(", current_card_block)
        self.assertIn("card.append(actionRow, header, statusSection);", current_card_block)
        self.assertNotIn("notesSection", current_card_block)

    def test_status_change_rerenders_projects_while_restoring_expanded_cards(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        helper_block = self._script_block(
            script,
            "function getExpandedProjectDeliverableIds() {",
            "function renderDeliverableCardLegacy(deliverable, isPrimary, project) {",
        )
        for expected in (
            'document.getElementById("tbody")',
            '.querySelectorAll(".deliverable-card-new[data-deliverable-id]")',
            'if (card.classList.contains("details-collapsed")) return;',
            "setDeliverableDetailsCollapsed(card, false);",
            "function renderProjectsPreservingExpandedDeliverables() {",
            "const expandedIds = getExpandedProjectDeliverableIds();",
            "render();",
            "restoreExpandedProjectDeliverables(expandedIds);",
        ):
            self.assertIn(expected, helper_block)

        status_dropdown_block = self._script_block(
            script,
            "function createStatusDropdown(deliverable, project, card) {",
            "function createDeliverableStatusSection(deliverable, project, card) {",
        )
        for expected in (
            "await save();",
            "setDeliverableStatusDropdownState(dropdown, false);",
            "renderProjectsPreservingExpandedDeliverables();",
        ):
            self.assertIn(expected, status_dropdown_block)

    def test_projects_pinned_work_items_styles_exist_and_notes_styles_removed(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".deliverable-status-row {",
            ".deliverable-status-row[hidden] {",
            ".deliverable-status-badges {",
            ".deliverable-status-inline-group {",
            ".deliverable-status-dropdown {",
            ".deliverable-card-action-row {",
            ".deliverable-card-action-group {",
            ".work-item-actions {",
            ".work-item-pin-btn {",
            ".work-item-pin-btn.is-pinned {",
            ".deliverable-pinned-inline-group {",
            ".deliverable-pinned-inline-pill {",
            ".deliverable-pinned-inline-pill-icon {",
            ".deliverable-pinned-inline-pill-text {",
            ".deliverable-pinned-inline-pill.is-task {",
            ".deliverable-pinned-item-attachments {",
            ".deliverable-pinned-link,",
        ):
            self.assertIn(expected, css)

        for removed in (
            ".deliverable-notes-footer {",
            ".deliverable-note-add-btn {",
            ".deliverable-note-row {",
            ".deliverable-note-delete-btn {",
            ".deliverable-note-add-input {",
            ".deliverable-notes-toggle {",
            ".deliverable-pinned-preview-heading {",
            ".deliverable-pinned-preview {",
            ".deliverable-pinned-item-kind {",
        ):
            self.assertNotIn(removed, css)

        action_row_block = self._css_block(css, ".deliverable-card-action-row {")
        self.assertIn("gap: 0.3rem;", action_row_block)
        self.assertIn("margin-bottom: 0.2rem;", action_row_block)

        card_block = self._css_block(css, ".deliverable-card-new {")
        self.assertIn("padding: 0.5rem 0.6rem;", card_block)

    def test_work_item_pin_button_supports_async_toggle_rollback(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        pin_button_block = self._script_block(
            script,
            "function createWorkItemPinButton({",
            "function getModalDeliverableCardContext(card) {",
        )

        for expected in (
            "let isTogglePending = false;",
            'button.setAttribute("aria-busy", String(isTogglePending));',
            "if (isTogglePending) return;",
            "const previousPinned = currentPinned;",
            "const nextPinned = !currentPinned;",
            "const resolvedPinned = onToggle ? await onToggle(nextPinned) : nextPinned;",
            'if (typeof resolvedPinned === "boolean") {',
            "currentPinned = previousPinned;",
            'console.warn("Failed to toggle work item pin state:", error);',
        ):
            self.assertIn(expected, pin_button_block)


if __name__ == "__main__":
    unittest.main()
