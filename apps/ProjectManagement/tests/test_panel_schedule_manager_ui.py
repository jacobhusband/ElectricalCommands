import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class PanelScheduleManagerUiTests(unittest.TestCase):
    def test_tool_card_and_editor_dialog_exist(self):
        html = (PROJECT_ROOT / "index.html").read_text(encoding="utf-8")
        for marker in (
            'id="toolPanelScheduleManager"',
            'id="panelScheduleManagerDlg"',
            'id="psmProjectSelect"',
            'id="psmWorkbookPath"',
            'id="psmConflictBanner"',
            'id="psmPanelList"',
            'id="psmCircuitRows"',
            'id="psmCircuitTableWrap"',
            'id="psmCircuitPageDownBtn"',
            'id="psmCircuitBottomBtn"',
            'id="psmReloadBtn"',
            'id="psmSaveBtn"',
        ):
            self.assertIn(marker, html)

    def test_bidirectional_sync_and_conflict_ui_are_wired(self):
        script = (PROJECT_ROOT / "script.js").read_text(encoding="utf-8")
        for marker in (
            "const panelScheduleManagerState = {",
            "async function openPanelScheduleManager() {",
            "async function connectPanelScheduleManager(",
            "async function pollPanelScheduleManager() {",
            "async function savePanelScheduleManager() {",
            "function showPanelScheduleManagerConflict(",
            '"data-psm-circuit-index": String(circuitIndex)',
            '"data-psm-circuit-field": field',
            "window.pywebview.api.start_panel_schedule_manager(",
            "window.pywebview.api.poll_panel_schedule_manager(",
            "window.pywebview.api.save_panel_schedule_manager(",
            "window.pywebview.api.reload_panel_schedule_manager(",
        ):
            self.assertIn(marker, script)

    def test_circuits_are_rendered_as_odd_and_even_panel_columns(self):
        html = (PROJECT_ROOT / "index.html").read_text(encoding="utf-8")
        script = (PROJECT_ROOT / "script.js").read_text(encoding="utf-8")
        styles = (PROJECT_ROOT / "styles.css").read_text(encoding="utf-8")

        self.assertIn("Odd circuits", html)
        self.assertIn("Even circuits", html)
        self.assertIn("function getPanelScheduleManagerCircuitPairs(", script)
        self.assertIn("function scrollPanelScheduleManagerCircuits(", script)
        self.assertIn("function handlePanelScheduleManagerCircuitWheel(", script)
        self.assertIn("appendPanelScheduleManagerCircuitNumber(row, pair.left", script)
        self.assertIn("appendPanelScheduleManagerCircuitNumber(row, pair.right", script)
        self.assertIn(".psm-circuit-phase", styles)

    def test_circuit_table_has_a_bounded_vertical_scroll_area(self):
        styles = (PROJECT_ROOT / "styles.css").read_text(encoding="utf-8")

        self.assertRegex(
            styles,
            r"(?s)dialog\.panel-schedule-manager-dialog\s*\{[^}]*"
            r"height:\s*calc\(100vh - 16px\);[^}]*"
            r"max-height:\s*calc\(100vh - 16px\);",
        )
        self.assertNotIn("height: min(940px, 94vh);", styles)
        self.assertRegex(
            styles,
            r"(?s)\.psm-workspace\s*\{[^}]*flex:\s*1 1 0;[^}]*"
            r"height:\s*0;[^}]*max-height:\s*100%;[^}]*min-height:\s*0;[^}]*"
            r"overflow:\s*hidden;",
        )
        self.assertRegex(
            styles,
            r"(?s)\.psm-editor\s*\{[^}]*flex:\s*1 1 0;[^}]*height:\s*100%;[^}]*"
            r"max-height:\s*100%;[^}]*display:\s*grid;[^}]*"
            r"grid-template-rows:\s*auto auto minmax\(0, 1fr\);[^}]*overflow:\s*hidden;",
        )
        self.assertRegex(
            styles,
            r"(?s)\.psm-circuit-section\s*\{[^}]*height:\s*100%;[^}]*"
            r"max-height:\s*100%;[^}]*min-height:\s*0;[^}]*display:\s*flex;[^}]*"
            r"flex-direction:\s*column;[^}]*overflow:\s*hidden;",
        )
        self.assertRegex(
            styles,
            r"(?s)\.psm-table-wrap\s*\{[^}]*flex:\s*1 1 0;[^}]*height:\s*0;[^}]*"
            r"max-height:\s*100%;[^}]*min-height:\s*0;[^}]*overflow-x:\s*auto;[^}]*"
            r"overflow-y:\s*scroll;",
        )

    def test_backend_api_exposes_edit_session_operations(self):
        source = (PROJECT_ROOT / "main.py").read_text(encoding="utf-8")
        for marker in (
            "def start_panel_schedule_manager(",
            "def poll_panel_schedule_manager(",
            "def save_panel_schedule_manager(",
            "def reload_panel_schedule_manager(",
            "def close_panel_schedule_manager(",
        ):
            self.assertIn(marker, source)


if __name__ == "__main__":
    unittest.main()
