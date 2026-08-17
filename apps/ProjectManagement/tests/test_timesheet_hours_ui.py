import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class TimesheetHoursUiTests(unittest.TestCase):
    def test_timesheet_hours_use_manual_number_inputs_with_24_hour_cap(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("const MAX_HOURS_PER_DAY = 24;", text)
        self.assertIn("cell.appendChild(createHourInput(entry, day, hours, row));", text)
        self.assertIn('className: "ts-hour-input"', text)
        self.assertIn('type: "number"', text)
        self.assertIn('max: String(MAX_HOURS_PER_DAY)', text)
        self.assertIn('step: "0.1"', text)
        self.assertIn("function normalizeTimesheetHours(value) {", text)
        self.assertIn("function getRemainingTimesheetHoursForDay(day, currentEntry) {", text)
        self.assertIn("function createHourInput(entry, day, hours, row) {", text)
        self.assertIn('if (rawValue === "") {', text)
        self.assertIn("const parsedHours = Number.parseFloat(rawValue);", text)
        self.assertIn("if (!Number.isFinite(parsedHours)) return;", text)
        self.assertNotIn(
            'e.target.value = rawValue === "" ? "" : formatTimesheetHours(newHours);',
            text,
        )
        self.assertNotIn("function createHourDragBar(", text)

    def test_timesheet_hour_input_styles_exist(self):
        text = STYLES_CSS_PATH.read_text(encoding="utf-8")
        self.assertIn(".ts-hour-input {", text)
        self.assertIn("appearance: textfield;", text)
        self.assertIn(".ts-hour-input::-webkit-outer-spin-button,", text)

    def test_timesheet_hours_table_no_longer_has_mileage_column(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertNotIn('<th class="ts-mileage-col">MILES</th>', html)
        self.assertNotIn('id="totalMileage"', html)
        self.assertNotIn("const totals = { week: 0, mileage: 0 };", script)
        self.assertNotIn("totals.mileage += entry.mileage || 0;", script)
        self.assertNotIn('const mileageCell = document.getElementById("totalMileage");', script)


if __name__ == "__main__":
    unittest.main()
