import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class ConsolidationMigrationsTests(unittest.TestCase):
    """One-time migrations fold legacy notes/coordination data into pages."""

    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_migration_functions_exist_and_are_guarded(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("function migrateProjectNotesToPage(out) {", script)
        self.assertIn("function migrateCoordinationItemsToPage(out) {", script)
        self.assertIn("function migrateDeliverablePagesToProjectSubpages(out) {", script)

        notes_block = self._block(
            script,
            "function migrateProjectNotesToPage(out) {",
            "function migrateCoordinationItemsToPage(out) {",
        )
        self.assertIn("if (!out || out.notesMigratedToPage) return;", notes_block)
        self.assertIn("out.notesMigratedToPage = true;", notes_block)
        self.assertIn("<h2>Notes</h2>", notes_block)
        self.assertIn("plainTextToPageHtml(text)", notes_block)

        coord_block = self._block(
            script,
            "function migrateCoordinationItemsToPage(out) {",
            "function normalizeProject(project) {",
        )
        self.assertIn("if (!out || out.coordinationMigratedToPage) return;", coord_block)
        self.assertIn("out.coordinationMigratedToPage = true;", coord_block)
        self.assertIn("out.coordinationItems", coord_block)
        # Party, due date, and done state are folded into plain page text.
        self.assertIn("parts.push(`[${party}]`)", coord_block)
        self.assertIn("parts.push(`(due ${due})`)", coord_block)
        self.assertIn('if (item?.done === true) parts.push("(done)");', coord_block)
        self.assertIn('plainTextToPageHtml(lines.join("\\n"))', coord_block)
        self.assertNotIn('data-tag="coordination"', coord_block)
        self.assertNotIn('class="page-item', coord_block)
        self.assertIn("<h2>Coordination</h2>", coord_block)

        deliverable_page_block = self._block(
            script,
            "function migrateDeliverablePagesToProjectSubpages(out) {",
            "function normalizeProject(project) {",
        )
        self.assertIn("if (!out || out.deliverablePagesMigratedToProjectSubpages) return;", deliverable_page_block)
        self.assertIn("out.deliverablePagesMigratedToProjectSubpages = true;", deliverable_page_block)
        self.assertIn("existingSources", deliverable_page_block)
        self.assertIn("sourceDeliverableId", deliverable_page_block)
        self.assertIn("createProjectSubpage({", deliverable_page_block)
        self.assertIn("html: page.html,", deliverable_page_block)

    def test_normalize_project_runs_migrations(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        normalize_block = self._block(
            script,
            "function normalizeProject(project) {",
            "function getLatestDueDeliverableId(deliverables = []) {",
        )
        self.assertIn("migrateProjectNotesToPage(out);", normalize_block)
        self.assertIn("migrateCoordinationItemsToPage(out);", normalize_block)
        self.assertIn("migrateDeliverablePagesToProjectSubpages(out);", normalize_block)
        self.assertIn("subpages: normalizeProjectSubpages(project.subpages),", normalize_block)
        # The coordination field is no longer normalized into the project.
        self.assertNotIn("coordinationItems: normalizeCoordinationItems", normalize_block)

    def test_legacy_notes_tabs_migrate_to_pages(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        legacy_block = self._block(
            script,
            "function migrateLegacyNotesToPages(source = {}) {",
            "function readGlobalPagesData(raw = {}) {",
        )
        self.assertIn("createGlobalPage({ title: tab, html: plainTextToPageHtml(text) })", legacy_block)
        # An empty default "General" tab is skipped to keep first-run clean.
        self.assertIn('if (!text && tab === "General") return;', legacy_block)

    def test_readform_preserves_page_and_guards(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        read_block = self._block(
            script,
            "function readForm() {",
            "function addRefRowFrom(L = {}) {",
        )
        self.assertIn("page: existingProject?.page,", read_block)
        self.assertIn("subpages: existingProject?.subpages || [],", read_block)
        self.assertIn(
            "notesMigratedToPage: existingProject?.notesMigratedToPage === true,",
            read_block,
        )
        self.assertIn("coordinationMigratedToPage:", read_block)
        self.assertIn("deliverablePagesMigratedToProjectSubpages:", read_block)
        # Legacy deliverable pages remain stored across modal edits, but are no longer surfaced.
        self.assertIn("page: existingDeliverable?.page,", read_block)


if __name__ == "__main__":
    unittest.main()
