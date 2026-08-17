import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectModalNoteUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_modal_note_editor_removed_but_note_data_preserved(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for removed in (
            "flushPendingModalDeliverableNotes();",
            "function renderModalDeliverableNoteList(card, container, options = {}) {",
            "function commitPendingModalDeliverableNote(card, container, options = {}) {",
            "function flushPendingModalDeliverableNotes() {",
            "renderModalDeliverableNoteList(card, noteList);",
        ):
            self.assertNotIn(removed, script)

        read_form_block = self._block(
            script,
            "function readForm() {",
            "function addRefRowFrom(",
        )
        self.assertIn("const sourceNoteItems = existingDeliverable", read_form_block)
        self.assertIn("normalizeDeliverableNoteItems(", read_form_block)
        self.assertIn("notes: existingDeliverable?.notes || buildDeliverableNotesText(noteItems, \"\"),", read_form_block)
        self.assertIn("noteItems,", read_form_block)
        self.assertIn("dueDate: normalizedNoteItem.dueDate,", read_form_block)
        self.assertIn("attachments: normalizedNoteItem.attachments,", read_form_block)
        self.assertIn("emailRefs: normalizedNoteItem.emailRefs,", read_form_block)

    def test_modal_note_editor_markup_and_styles_removed(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertNotIn('<div class="title section-header">Notes</div>', html)
        self.assertNotIn('<div class="deliverable-note-list stack"></div>', html)
        self.assertNotIn(".deliverable-note-list {", css)
        self.assertNotIn(".deliverable-note-item {", css)
        self.assertNotIn(".deliverable-notes-list-edit {", css)
        self.assertNotIn(".note-add-bullet {", css)


if __name__ == "__main__":
    unittest.main()
