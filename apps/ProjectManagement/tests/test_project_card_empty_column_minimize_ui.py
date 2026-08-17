import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectCardEmptyColumnMinimizeUiTests(unittest.TestCase):
    def test_empty_column_minimize_toggle_is_default_on_and_persisted(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn('id="projectsEmptyColumnsToggle"', html)
        self.assertIn('aria-label="Minimize empty columns"', html)
        self.assertIn('aria-pressed="true"', html)

        for expected in (
            "minimizeEmptyProjectColumns: true,",
            "let minimizeEmptyProjectColumns = true;",
            "userSettings.minimizeEmptyProjectColumns !== false",
            "userSettings.minimizeEmptyProjectColumns = minimizeEmptyProjectColumns;",
            "function setMinimizeEmptyProjectColumns(enabled, options = {}) {",
            "setMinimizeEmptyProjectColumns(!minimizeEmptyProjectColumns)",
        ):
            self.assertIn(expected, script)

    def test_empty_column_hide_toggle_is_default_off_and_persisted(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn('id="projectsHideEmptyColumnsToggle"', html)
        self.assertIn('aria-label="Hide empty columns"', html)
        self.assertIn('aria-pressed="false"', html)

        for expected in (
            "hideEmptyProjectColumns: false,",
            "let hideEmptyProjectColumns = false;",
            "userSettings.hideEmptyProjectColumns === true",
            "userSettings.hideEmptyProjectColumns = hideEmptyProjectColumns;",
            "function setHideEmptyProjectColumns(enabled, options = {}) {",
            "setHideEmptyProjectColumns(!hideEmptyProjectColumns)",
        ):
            self.assertIn(expected, script)

    def test_empty_columns_are_minimized_without_removing_drop_target(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        render_start = script.index("function renderCardView(")
        render_end = script.index("function setProjectsViewMode(", render_start)
        render_block = script[render_start:render_end]
        drag_start = script.index('host.addEventListener("dragover", (e) => {')
        drag_end = script.index('host.addEventListener("drop", async (e) => {', drag_start)
        drag_block = script[drag_start:drag_end]

        self.assertIn('className: `kanban-column kanban-column--${slug}`', render_block)
        self.assertIn('column.classList.toggle("is-empty", bucketRows.length === 0);', render_block)
        self.assertIn('const cardsHost = el("div", { className: "kanban-column__cards" });', render_block)
        self.assertIn('className: "kanban-column__empty"', render_block)
        self.assertIn('e.target.closest(".kanban-column")', drag_block)

        self.assertIn("--projects-card-empty-column-width: 112px;", css)
        self.assertIn(
            ".projects-card-view.minimize-empty-columns .kanban-column.is-empty {",
            css,
        )
        self.assertIn("flex-basis: var(--projects-card-empty-column-width);", css)

    def test_empty_columns_can_be_hidden_after_bucket_counts_are_known(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        render_start = script.index("function renderCardView(")
        render_end = script.index("function setProjectsViewMode(", render_start)
        render_block = script[render_start:render_end]

        for expected in (
            "const renderColumns =\n    hideEmptyProjectColumns === true",
            "visibleColumns.filter(",
            "(column) => (buckets.get(column.key) || []).length > 0",
            "for (const col of renderColumns) {",
            'className: "projects-card-view__empty-board"',
            "No deliverables in visible columns.",
        ):
            self.assertIn(expected, render_block)

        self.assertIn(
            ".projects-card-view.hide-empty-columns .kanban-column.is-empty {",
            css,
        )
        self.assertIn("display: none;", css)


if __name__ == "__main__":
    unittest.main()
