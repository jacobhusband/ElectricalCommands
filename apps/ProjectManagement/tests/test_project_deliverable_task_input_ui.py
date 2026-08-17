import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class ProjectDeliverableTaskInputUiTests(unittest.TestCase):
    def test_projects_note_icon_input_removed_from_cards(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertNotIn("function createNotesSection(", script)
        self.assertNotIn('className: "deliverable-note-add-btn"', script)
        self.assertNotIn('className: "deliverable-note-add-input"', script)
        self.assertNotIn('addNoteBtn.addEventListener("click"', script)
        self.assertNotIn("isAddingNote = true;", script)
        self.assertNotIn("notesExpanded = true;", script)
        self.assertNotIn("deliverable.noteItems.push({", script)


if __name__ == "__main__":
    unittest.main()
