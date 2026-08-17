import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class ProjectSearchViewModeUiTests(unittest.TestCase):
    def test_project_search_from_card_view_switches_to_list_without_saving_preference(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        handler_start = script.index("function handleProjectSearchInput() {")
        handler_end = script.index("function initEventListeners() {", handler_start)
        handler_block = script[handler_start:handler_end]
        helper_start = script.index("function clearProjectSearchInput() {")
        setter_start = script.index("function setProjectsViewMode(mode, options = {}) {")
        setter_end = script.index("function shiftProjectCardWeek(days) {", setter_start)
        helper_block = script[helper_start:setter_start]
        setter_block = script[setter_start:setter_end]
        listeners_start = script.index("function initEventListeners() {")
        listeners_end = script.index(
            'document.getElementById("notesSearch").addEventListener',
            listeners_start,
        )
        listeners_block = script[listeners_start:listeners_end]

        self.assertIn('projectsViewMode !== "list" && val("search")', handler_block)
        self.assertIn('setProjectsViewMode("list", { persist: false });', handler_block)
        self.assertIn("render();", handler_block)
        self.assertIn("debounce(handleProjectSearchInput, 250)", listeners_block)
        self.assertIn("const shouldPersist = options.persist !== false;", setter_block)
        self.assertIn(
            "if (shouldPersist && userSettings.projectsViewMode !== next) {",
            setter_block,
        )
        self.assertIn("function clearProjectSearchInput() {", helper_block)
        self.assertIn('const searchInput = document.getElementById("search");', helper_block)
        self.assertIn('searchInput.value = "";', helper_block)
        self.assertIn(
            'const didClearProjectSearch = next !== "list" && clearProjectSearchInput();',
            setter_block,
        )
        self.assertIn("if (didClearProjectSearch) render();", setter_block)
        self.assertNotIn(
            'cardBtn.addEventListener("click", () => clearProjectSearchInput())',
            listeners_block,
        )


if __name__ == "__main__":
    unittest.main()
