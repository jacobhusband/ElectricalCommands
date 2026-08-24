import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectPaginationUiTests(unittest.TestCase):
    def test_projects_list_has_accessible_pagination_controls(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        table_end = html.index("</table>", html.index('class="table projects-table"'))
        empty_state_start = html.index('<div id="emptyState"', table_end)
        pagination_block = html[table_end:empty_state_start]

        self.assertIn('id="projectsPagination"', pagination_block)
        self.assertIn('aria-label="Projects list pagination"', pagination_block)
        self.assertIn('id="projectsPaginationSummary"', pagination_block)
        self.assertIn('aria-live="polite"', pagination_block)
        self.assertIn('id="projectsPageSize"', pagination_block)
        self.assertIn('<option value="25" selected>25</option>', pagination_block)
        self.assertIn('id="projectsPageFirst"', pagination_block)
        self.assertIn('id="projectsPagePrev"', pagination_block)
        self.assertIn('id="projectsPageNext"', pagination_block)
        self.assertIn('id="projectsPageLast"', pagination_block)

    def test_render_pages_filtered_sorted_deliverables_before_building_cards(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        render_start = script.index("function render() {")
        render_end = script.index("function renderStatusToggles(", render_start)
        render_block = script[render_start:render_end]

        self.assertIn("function paginateProjectsListItems(items = []) {", script)
        self.assertIn("function updateProjectsPaginationUi(", script)
        self.assertIn("const deliverableRows = buildProjectDeliverableRowEntries(", render_block)
        self.assertIn("sortProjectDeliverableRows(deliverableRows);", render_block)
        self.assertIn("const pagination = paginateProjectsListItems(deliverableRows);", render_block)
        self.assertIn("paginatedDeliverableRows: pagination.items,", render_block)
        self.assertIn("const pagination = paginateProjectsListItems(items);", render_block)
        self.assertIn('updateProjectsPaginationUi(pagination, "projects");', render_block)
        self.assertNotIn("getAllDeliverables().forEach", render_block)

    def test_pagination_resets_for_search_and_filter_changes(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        filter_start = script.index("function setProjectsFilterValue(")
        filter_end = script.index("function getProjectsFilterOptions(", filter_start)
        search_start = script.index("function handleProjectSearchInput() {")
        search_end = script.index("function getActiveHelpTopic()", search_start)

        self.assertIn("resetProjectsListPagination();", script[filter_start:filter_end])
        self.assertIn("resetProjectsListPagination();", script[search_start:search_end])
        self.assertIn('getElementById("projectsPageSize")', script)
        self.assertIn('getElementById("projectsPageNext")', script)

    def test_pagination_styles_support_hidden_and_wrapped_controls(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".projects-pagination {", css)
        self.assertIn(".projects-pagination[hidden] {", css)
        self.assertIn(".projects-pagination-controls {", css)
        self.assertIn(".projects-pagination-btn:disabled {", css)


if __name__ == "__main__":
    unittest.main()
