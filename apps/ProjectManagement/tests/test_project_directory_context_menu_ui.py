import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class ProjectDirectoryContextMenuUiTests(unittest.TestCase):
    def test_project_directory_context_menu_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "let projectDirectoryContextMenu = null;",
            "function openProjectDirectory(project, kind) {",
            "window.pywebview.api.open_project_directory(",
            "function ensureProjectDirectoryContextMenu() {",
            'textContent: "Open Server Directory",',
            'textContent: "Open Local Directory",',
            'data-project-directory-action": "server",',
            'data-project-directory-action": "local",',
            "function attachProjectDirectoryContextMenu(target, project) {",
            'target.addEventListener("contextmenu", (event) => {',
            "showProjectDirectoryContextMenu(event, project);",
            "function attachCardProjectPathOpen(target, project) {",
            'target.dataset.projectPathOpen = "true";',
            'target.addEventListener("click", openPath);',
            'target.addEventListener("keydown", (event) => {',
            "window.pywebview.api.open_path(convertPath(path));",
        ):
            self.assertIn(expected, script)

    def test_list_and_card_project_names_attach_context_menu(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        table_start = script.index("function buildProjectTableRow(")
        table_end = script.index("function renderGroupedProjectDeliverablesCell(", table_start)
        table_block = script[table_start:table_end]

        card_start = script.index("function buildCardProjectMeta(")
        card_end = script.index("function renderCardView(", card_start)
        card_block = script[card_start:card_end]

        self.assertIn("attachProjectDirectoryContextMenu(link, project);", table_block)
        self.assertIn("attachProjectDirectoryContextMenu(titleText, project);", table_block)
        self.assertIn("attachProjectDirectoryContextMenu(nameEl, project);", card_block)
        self.assertIn("attachCardProjectPathOpen(nameEl, project);", card_block)

    def test_project_directory_context_menu_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".project-directory-context-menu {",
            "position: fixed;",
            ".project-directory-context-menu[hidden] {",
            ".project-directory-context-menu__item {",
            ".project-title-text[data-project-directory-menu]",
            ".kanban-card-project-name[data-project-directory-menu]",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
