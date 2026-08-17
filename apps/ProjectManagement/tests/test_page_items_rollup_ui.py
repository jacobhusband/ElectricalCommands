import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class PageItemsRollupUiTests(unittest.TestCase):
    """The cross-project action/coordination roll-up view has been removed."""

    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_rollup_view_mode_and_markup_removed(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for removed in (
            "function renderPageItemsRollupView() {",
            "function refreshRollupIfActive() {",
            "function extractPageTaggedItems(html) {",
            "function collectRollupSources() {",
            "function renderRollupSourcePanel(source) {",
            'setProjectsViewMode("rollup")',
            'projectsViewMode === "rollup"',
            'document.getElementById("viewToggleRollup")',
            'document.getElementById("projectsRollupView")',
        ):
            self.assertNotIn(removed, script)

        normalize_block = self._block(
            script,
            "function normalizeProjectsViewMode(value) {",
            "function setProjectsViewMode(mode, options = {}) {",
        )
        self.assertIn('if (value === "card") return "card";', normalize_block)
        self.assertNotIn('"rollup"', normalize_block)
        self.assertNotIn('"coordination"', normalize_block)

        self.assertNotIn('id="viewToggleRollup"', html)
        self.assertNotIn('id="projectsRollupView"', html)
        self.assertNotIn(".projects-rollup-view {", css)
        self.assertNotIn(".rollup-source-panel {", css)
        self.assertNotIn(".rollup-item-row {", css)

    def test_legacy_coordination_migration_uses_plain_page_text(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        coord_block = self._block(
            script,
            "function migrateCoordinationItemsToPage(out) {",
            "function normalizeProject(project) {",
        )

        self.assertIn('const block = `<h2>Coordination</h2>${plainTextToPageHtml(lines.join("\\n"))}`;', coord_block)
        self.assertIn('if (item?.done === true) parts.push("(done)");', coord_block)
        self.assertNotIn('data-tag="coordination"', coord_block)
        self.assertNotIn('class="page-item', coord_block)


if __name__ == "__main__":
    unittest.main()
