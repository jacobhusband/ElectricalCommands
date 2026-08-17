import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"


class ProjectModalTaskUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_modal_task_section_is_removed_but_legacy_tasks_are_preserved(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        save_block = self._block(
            script,
            "function onSaveProject() {",
            "// ===================== FORM HANDLING =====================",
        )
        read_form_block = self._block(
            script,
            "function readForm() {",
            "function addRefRowFrom(L = {}) {",
        )

        self.assertNotIn("flushPendingModalDeliverableTasks();", save_block)
        self.assertNotIn('<div class="title section-header">Tasks</div>', html)
        self.assertNotIn('<div class="deliverable-task-list stack"></div>', html)
        self.assertIn("existingDeliverable.tasks", read_form_block)
        self.assertIn("getModalDeliverableTaskItems(card)", read_form_block)
        self.assertIn("tasksMigratedToNotes: existingDeliverable?.tasksMigratedToNotes === true,", read_form_block)


if __name__ == "__main__":
    unittest.main()
