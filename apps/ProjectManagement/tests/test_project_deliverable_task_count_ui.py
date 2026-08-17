import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class ProjectDeliverableTaskCountUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_current_card_render_does_not_render_task_count_or_progress(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        header_block = self._block(
            script,
            "function createCardHeader(deliverable, isPrimary, card, project) {",
            "function createExpandToggle(card) {",
        )
        card_block = self._block(
            script,
            "function renderDeliverableCard(deliverable, isPrimary, project) {",
            "function normalizeProjectMatchValue(value) {",
        )

        self.assertNotIn("createDeliverableTaskCountBadge(deliverable)", header_block)
        self.assertNotIn("leftSection.appendChild(taskCountBadge);", header_block)
        self.assertNotIn("createProgressSection(deliverable)", card_block)
        self.assertNotIn("createTasksPreview(deliverable", card_block)
        self.assertIn("card.append(actionRow, header, statusSection);", card_block)


if __name__ == "__main__":
    unittest.main()
