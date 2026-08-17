import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverableInlineEditUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    @staticmethod
    def _css_block(css: str, selector: str) -> str:
        block_start = css.index(selector)
        block_end = css.index("}", block_start)
        return css[block_start:block_end]

    def test_inline_edit_helper_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        helper_block = self._block(
            script,
            "function createWorkItemDoneCheckbox({",
            "function createTasksPreview(deliverable, card, project = null) {",
        )

        for expected in (
            "function createWorkItemDoneCheckbox({",
            'className: "work-item-checkbox",',
            "function createInlineWorkItemTextControl({",
            'className: "work-item-text-wrap"',
            'className: `work-item-text-trigger ${textClassName}`.trim(),',
            'className: "work-item-text-input",',
            'title = "Click to edit",',
            "const nextText = trimmedText || previousText;",
            "await onCommit(nextText, previousText);",
            'void finishEditing("commit");',
            'void finishEditing("cancel");',
            '} else if (event.key === "Escape") {',
        ):
            self.assertIn(expected, helper_block)

    def test_projects_tab_no_longer_renders_note_inline_editing(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        card_block = self._block(
            script,
            "function renderDeliverableCard(deliverable, isPrimary, project) {",
            "function normalizeProjectMatchValue(value) {",
        )

        self.assertNotIn("function createNotesSection(", script)
        self.assertNotIn("const notesSection", card_block)
        self.assertIn("card.append(actionRow, header, statusSection);", card_block)
        self.assertNotIn("createTasksPreview(", card_block)
        self.assertNotIn("createProgressSection(", card_block)

    def test_modal_task_inline_edit_script_wiring_exists_without_note_editor(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        modal_task_block = self._block(
            script,
            "function renderModalDeliverableTaskList(card, container, options = {}) {",
            "function commitPendingModalDeliverableTask(card, container, options = {}) {",
        )

        for expected in (
            'const textControl = createInlineWorkItemTextControl({',
            'title: "Click to edit task",',
            "liveTask.text = nextText;",
        ):
            self.assertIn(expected, modal_task_block)

        self.assertNotIn("function renderModalDeliverableNoteList(", script)
        self.assertNotIn("function commitPendingModalDeliverableNote(", script)

    def test_inline_edit_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".work-item-checkbox {",
            ".work-item-checkbox:focus-visible {",
            ".work-item-text-wrap {",
            ".work-item-text-trigger {",
            ".work-item-text-trigger:hover {",
            ".work-item-text-trigger:focus-visible {",
            ".work-item-text-input {",
        ):
            self.assertIn(expected, css)

        item_block = self._css_block(css, ".deliverable-task-item {")
        self.assertNotIn("cursor: pointer;", item_block)
        self.assertNotIn("user-select: none;", item_block)
        self.assertNotIn(".deliverable-note-item", css)

        text_trigger_block = self._css_block(css, ".work-item-text-trigger {")
        self.assertIn("cursor: text;", text_trigger_block)
        self.assertIn("background: transparent;", text_trigger_block)

        checkbox_block = self._css_block(css, ".work-item-checkbox {")
        self.assertIn("accent-color: var(--accent);", checkbox_block)
        self.assertIn("cursor: pointer;", checkbox_block)

    def test_card_view_text_editors_do_not_start_kanban_drag(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        drag_block = self._block(
            script,
            "let kanbanDragState = null;",
            "function render() {",
        )

        for expected in (
            "const KANBAN_TEXT_INPUT_TYPES = new Set([",
            "function getKanbanTextEditorFromTarget(target) {",
            'element.closest("textarea, input, [contenteditable]")',
            'card.dataset.kanbanDragSuppressed = "1";',
            "card.draggable = false;",
            'host.addEventListener("pointerdown", handleEditorPointerStart, true);',
            'host.addEventListener("mousedown", handleEditorPointerStart, true);',
            'card.dataset.kanbanDragSuppressed === "1"',
            "e.preventDefault();",
            "e.stopPropagation();",
        ):
            self.assertIn(expected, drag_block)

        self.assertLess(
            drag_block.index(
                'host.addEventListener("pointerdown", handleEditorPointerStart, true);'
            ),
            drag_block.index('host.addEventListener("dragstart", (e) => {'),
        )


if __name__ == "__main__":
    unittest.main()
