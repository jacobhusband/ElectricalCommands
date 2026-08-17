import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class PageTaggedItemsUiTests(unittest.TestCase):
    """Action/coordination tagging UI is removed; legacy markup opens as text."""

    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_tag_authoring_helpers_and_toolbar_removed(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for removed in (
            "const PAGE_ITEM_TAGS = {",
            "function insertPageItem(tag) {",
            "function togglePageItem(box) {",
            "function handlePageItemEnter(editor) {",
            "function clearPageItemTag(editor) {",
            "function upgradeLegacyPageTodos(editor) {",
            'else if (btn.dataset.pageAction === "action")',
            'else if (btn.dataset.pageAction === "coordination")',
            'else if (btn.dataset.pageAction === "untag")',
        ):
            self.assertNotIn(removed, script)

        self.assertNotIn('data-page-action="action"', html)
        self.assertNotIn('data-page-action="coordination"', html)
        self.assertNotIn('data-page-action="untag"', html)
        self.assertNotIn("action/coordination items", html)
        self.assertNotIn("action items", html)
        self.assertNotIn("coordination items", html)

        self.assertNotIn(".page-item {", css)
        self.assertNotIn(".page-item-box {", css)
        self.assertNotIn(".page-item-badge {", css)
        self.assertNotIn(".page-item-text {", css)

    def test_legacy_tagged_items_flatten_to_plain_lines(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        flatten_block = self._block(
            script,
            "function flattenLegacyPageItems(editor) {",
            "// --- Rich-text commands ---",
        )

        self.assertIn('editor.querySelectorAll(".page-todo").forEach', flatten_block)
        self.assertIn('flatten(todo, ".page-todo-text");', flatten_block)
        self.assertIn('editor.querySelectorAll(".page-item").forEach', flatten_block)
        self.assertIn('flatten(item, ".page-item-text");', flatten_block)
        self.assertIn("line.innerHTML = textEl ? textEl.innerHTML : \"<br>\";", flatten_block)
        self.assertIn("node.replaceWith(line);", flatten_block)
        self.assertIn("flattenLegacyPageItems(editor);", script)


if __name__ == "__main__":
    unittest.main()
