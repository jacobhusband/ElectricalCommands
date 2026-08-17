import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class ProjectDeliverableNotesMigrationUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_legacy_tasks_are_migrated_to_notes_once_and_preserved(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        migration_block = self._block(
            script,
            "function migrateDeliverableTasksToNotes(deliverable) {",
            "function buildDeliverableNotesText(noteItems = [], fallback = \"\") {",
        )
        normalize_block = self._block(
            script,
            "function normalizeDeliverable(deliverable = {}) {",
            "function createDeliverable(seed = {}) {",
        )

        for expected in (
            "if (deliverable.tasksMigratedToNotes === true) return deliverable;",
            "const tasks = Array.isArray(deliverable.tasks) ? deliverable.tasks : [];",
            "const existingNoteTexts = new Set(",
            "if (!text || existingNoteTexts.has(text)) return;",
            "pinned: !!normalizedTask.pinned,",
            "attachments: normalizedTask.attachments,",
            "links: normalizedTask.links,",
            "emailRefs: normalizedTask.emailRefs,",
            "deliverable.tasksMigratedToNotes = true;",
        ):
            self.assertIn(expected, migration_block)

        self.assertIn("tasksMigratedToNotes: deliverable.tasksMigratedToNotes === true,", normalize_block)
        self.assertIn("syncDeliverableWorkItemFields(out);", normalize_block)
        self.assertIn("tasks: Array.isArray(deliverable.tasks)", normalize_block)

    def test_pinned_preview_no_longer_uses_notes_or_hidden_tasks(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        preview_block = self._block(
            script,
            "function getPinnedDeliverablePreviewItems(deliverable) {",
            "function normalizePinnedProjectOrder(value) {",
        )

        self.assertIn("syncDeliverableWorkItemFields(deliverable);", preview_block)
        self.assertIn("return [];", preview_block)
        self.assertNotIn("getPinnedDeliverableNoteEntries(deliverable)", preview_block)
        self.assertNotIn("getPinnedDeliverableTaskEntries(deliverable)", preview_block)


if __name__ == "__main__":
    unittest.main()
