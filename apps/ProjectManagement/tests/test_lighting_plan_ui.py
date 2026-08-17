import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]


class LightingPlanUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.html = (REPO_ROOT / "index.html").read_text(encoding="utf-8")
        cls.script = (REPO_ROOT / "script.js").read_text(encoding="utf-8")
        cls.styles = (REPO_ROOT / "styles.css").read_text(encoding="utf-8")

    def test_fixture_scheduler_contains_phase_one_review_workspace(self):
        for element_id in (
            "lightingPlanImportBtn",
            "lightingPlanExportBtn",
            "lightingPlanMetrics",
            "lightingPlanFixtureCounts",
            "lightingPlanCircuits",
            "lightingPlanWarnings",
        ):
            self.assertIn(f'id="{element_id}"', self.html)
        self.assertIn("Lighting Plan Automation", self.html)
        self.assertIn("LIGHTPLANSCAN", self.html)

    def test_ui_calls_project_scoped_backend_apis(self):
        self.assertIn("async function loadLightingPlanRecord(", self.script)
        self.assertIn("async function importLightingPlanSnapshot()", self.script)
        self.assertIn("async function exportLightingPlanInstructions()", self.script)
        self.assertIn("get_lighting_plan_record(projectId)", self.script)
        self.assertIn("import_lighting_plan_snapshot(\n      projectId,", self.script)
        self.assertIn("export_lighting_plan_instructions(\n      projectId", self.script)
        self.assertIn("await loadLightingPlanRecord(db[index]);", self.script)

    def test_generation_button_is_blocked_by_analysis_errors(self):
        self.assertIn("!analysis.canGenerate", self.script)
        self.assertIn(".lighting-plan-warning.is-error", self.styles)


if __name__ == "__main__":
    unittest.main()
