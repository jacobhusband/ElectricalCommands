import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverableNoteDueDateUiTests(unittest.TestCase):
    @staticmethod
    def _block(text, start_marker, end_marker):
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_note_data_model_still_preserves_due_date(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        for expected in (
            "function normalizeDeliverableNoteDueDate(",
            'dueDate: normalizeDeliverableNoteDueDate(noteItem?.dueDate)',
            "dueDate: normalizeDeliverableNoteDueDate(out.dueDate)",
        ):
            self.assertIn(expected, script)

        sync_block = self._block(
            script,
            "function syncDeliverableWorkItemFields(",
            "function getPinnedFirstEntries(",
        )
        self.assertIn("noteItem.dueDate = normalizedNoteItem.dueDate;", sync_block)

    def test_note_sort_helper_keeps_due_date_order_for_preserved_data(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        sort_block = self._block(
            script,
            "function getDeliverableNoteSortEntries(",
            "function getPinnedDeliverablePreviewItems(",
        )
        self.assertIn("normalizeDeliverableNoteItem(item)", sort_block)
        self.assertIn("parseDueStr(left.item?.dueDate)", sort_block)
        self.assertIn("parseDueStr(right.item?.dueDate)", sort_block)
        self.assertIn("return leftDue - rightDue;", sort_block)
        self.assertIn("return left.index - right.index;", sort_block)

    def test_modal_save_preserves_existing_note_due_date_without_note_ui(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        read_form_block = self._block(
            script,
            "function readForm(",
            "function addRefRowFrom(",
        )
        self.assertIn("const sourceNoteItems = existingDeliverable", read_form_block)
        self.assertIn("dueDate: normalizedNoteItem.dueDate,", read_form_block)
        self.assertIn("noteItems,", read_form_block)

    def test_note_due_date_ui_removed(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertNotIn("function createNoteDueDateControl(", script)
        self.assertNotIn("function renderModalDeliverableNoteList(", script)
        self.assertNotIn("function createNotesSection(", script)
        for removed in (
            ".note-due-wrap {",
            ".note-due-badge {",
            ".note-due-input {",
            ".note-due-add-btn {",
        ):
            self.assertNotIn(removed, css)


if __name__ == "__main__":
    unittest.main()
