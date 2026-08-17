import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectCardScrollUiTests(unittest.TestCase):
    def test_projects_tab_can_use_controlled_full_width_layout(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(
            '<body data-active-tab="projects" data-projects-layout-width="wide">',
            html,
        )
        self.assertIn('id="projectsWidthToggle"', html)
        self.assertIn('aria-pressed="true"', html)
        self.assertIn(
            'body[data-active-tab="projects"][data-projects-layout-width="wide"] main {',
            css,
        )

        layout_start = css.index(
            'body[data-active-tab="projects"][data-projects-layout-width="wide"] main {'
        )
        layout_end = css.index(":root {", layout_start)
        layout_block = css[layout_start:layout_end]
        self.assertIn("max-width: none;", layout_block)
        self.assertIn("width: 100%;", layout_block)

        self.assertIn("projectsWideLayout: true,", script)
        self.assertIn("let projectsWideLayout = true;", script)
        self.assertIn("userSettings.projectsWideLayout !== false", script)
        self.assertIn("function updateProjectsLayoutWidthUi() {", script)
        self.assertIn(
            'document.body.dataset.projectsLayoutWidth = isWide ? "wide" : "contained";',
            script,
        )
        self.assertIn("function setProjectsWideLayout(enabled, options = {}) {", script)
        self.assertIn("setProjectsWideLayout(!projectsWideLayout)", script)
        self.assertIn("document.body.dataset.activeTab =", script)
        self.assertIn(
            'document.querySelector(".main-tab-btn.active")?.dataset.tab || "projects"',
            script,
        )
        self.assertIn("document.body.dataset.activeTab = tab;", script)

    def test_card_view_horizontal_scroll_is_free_scrolling(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        view_start = css.index(".projects-card-view {")
        view_end = css.index(".projects-card-view[hidden]", view_start)
        view_block = css[view_start:view_end]
        column_start = css.index(".kanban-column {")
        column_end = css.index(".kanban-card-project-meta", column_start)
        column_block = css[column_start:column_end]

        self.assertIn("overflow-x: auto;", view_block)
        self.assertNotIn("scroll-snap-type", view_block)
        self.assertNotIn("scroll-snap-align", column_block)

    def test_card_view_fills_available_window_height(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        view_start = css.index(".projects-card-view {")
        view_end = css.index(".projects-card-view[hidden]", view_start)
        view_block = css[view_start:view_end]
        column_start = css.index(".kanban-column {")
        column_end = css.index(".kanban-card-project-meta", column_start)
        column_block = css[column_start:column_end]

        self.assertIn("--projects-card-view-height: max(", view_block)
        self.assertIn(
            "calc(100vh - var(--header-height) - var(--toolbar-height) - 5.5rem)",
            view_block,
        )
        self.assertIn("min-height: var(--projects-card-view-height);", view_block)
        self.assertIn("min-height: var(--projects-card-view-height);", column_block)
        self.assertIn("max-height: var(--projects-card-view-height);", column_block)
        self.assertNotIn("min-height: 300px;", view_block)
        self.assertNotIn("max-height: calc(100vh - 240px);", column_block)
        self.assertNotIn("min-height: 200px;", column_block)


if __name__ == "__main__":
    unittest.main()
