import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class TimesheetSuggestionsUiTests(unittest.TestCase):
    def test_timesheet_suggestions_collect_project_ids_from_all_current_week_entries(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("const projectIdsWithEntries = new Set();", text)
        self.assertIn(
            'const projectIdKey = String(entry?.projectId || "")',
            text,
        )
        self.assertIn(".trim()", text)
        self.assertIn(".toLowerCase();", text)
        self.assertIn("if (projectIdKey) projectIdsWithEntries.add(projectIdKey);", text)
        self.assertIn("if (!isProjectSummaryEntry(entry)) return;", text)

    def test_timesheet_suggestions_mute_cards_from_project_id_presence(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("hasProjectEntry = false,", text)
        self.assertIn("const showAddedStyle = hasProjectEntry;", text)
        self.assertIn("? projectIdsWithEntries.has(projectIdKey)", text)
        self.assertIn(": entryStatus.count > 0;", text)
        self.assertNotIn("const showAddedStyle = count > 0 && !needsUpdate;", text)

    def test_timesheet_suggestions_keep_needs_update_copy_separate_from_muted_state(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("if (needsUpdate) {", text)
        self.assertIn('textContent: "Click to add another entry"', text)
        self.assertIn("} else if (hasProjectEntry) {", text)
        self.assertIn('textContent: "Added to timesheet"', text)


if __name__ == "__main__":
    unittest.main()
