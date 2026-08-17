import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDeliverableTaskPreviewUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_projects_deliverable_cards_no_longer_render_notes_footer(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        card_block = self._block(
            script,
            "function renderDeliverableCard(deliverable, isPrimary, project) {",
            "function normalizeProjectMatchValue(value) {",
        )

        self.assertNotIn("function createNotesSection(", script)
        self.assertNotIn("const notesSection =", card_block)
        self.assertIn("card.append(actionRow, header, statusSection);", card_block)
        self.assertNotIn("deliverable-note", card_block)

    def test_projects_notes_footer_styles_removed(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for removed in (
            ".deliverable-notes-footer {",
            ".deliverable-notes-footer-controls {",
            ".deliverable-note-add-btn {",
            ".deliverable-note-row {",
            ".deliverable-note-row .note-text {",
            ".deliverable-note-delete-btn {",
            ".deliverable-note-add-input {",
            ".deliverable-notes-toggle {",
        ):
            self.assertNotIn(removed, css)


if __name__ == "__main__":
    unittest.main()
