import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class WorkroomPhaseDefaultsUiTests(unittest.TestCase):
    def test_index_html_omits_removed_survey_tools_and_dialog(self):
        html = (ROOT / "index.html").read_text(encoding="utf-8")

        self.assertNotIn('id="toolUtilityPlanEditor"', html)
        self.assertNotIn('id="toolCreateElectricalSurveyTemplate"', html)
        self.assertNotIn('id="utilityPlanEditorDlg"', html)
        self.assertNotIn('src="utility-plan-editor.js"', html)

    def test_script_js_defaults_to_pre_design_without_survey_state(self):
        script = (ROOT / "script.js").read_text(encoding="utf-8")

        self.assertIn('return WORKROOM_PHASE_VALUES.has(raw) ? raw : "pre_design";', script)
        self.assertIn('workroomPhase: seed.workroomPhase || "pre_design"', script)
        self.assertNotIn("survey_report", script)
        self.assertNotIn("utility_plan", script)
        self.assertNotIn("workroomSurveyPage", script)
        self.assertNotIn("toolUtilityPlanEditor", script)
        self.assertNotIn("toolCreateElectricalSurveyTemplate", script)
        self.assertNotIn("electricalSurvey", script)
        self.assertNotIn("utilityPlanState", script)
        self.assertNotIn("normalizeWorkroomSurveyPage", script)

    def test_backend_assets_and_styles_remove_survey_feature_files(self):
        main_text = (ROOT / "main.py").read_text(encoding="utf-8")
        styles = (ROOT / "styles.css").read_text(encoding="utf-8")

        self.assertFalse((ROOT / "utility-plan-editor.js").exists())
        self.assertFalse((ROOT / "templates" / "Electrical Survey Report Template.doc").exists())
        self.assertNotIn("dialog.utility-plan-dialog", styles)
        self.assertNotIn(".utility-plan-shell", styles)
        self.assertNotIn("def get_survey_report_draft(", main_text)
        self.assertNotIn("def save_survey_report_draft(", main_text)
        self.assertNotIn("def discover_survey_arch_pdf(", main_text)
        self.assertNotIn("def get_utility_plan_pdf_info(", main_text)
        self.assertNotIn("def get_utility_plan_page_preview(", main_text)
        self.assertNotIn("def export_utility_plan_pngs(", main_text)
        self.assertNotIn("SURVEY_REPORT_DB_FILE", main_text)
        self.assertNotIn("electricalSurvey", main_text)


if __name__ == "__main__":
    unittest.main()
