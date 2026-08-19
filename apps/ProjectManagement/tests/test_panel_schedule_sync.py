"""
Unit tests for apps/ProjectManagement/panel_schedule_sync.py
"""

import os
import tempfile
import unittest
import openpyxl

from apps.ProjectManagement.database import init_db, upsert_project, get_panel_schedule
from apps.ProjectManagement.panel_schedule_sync import (
    calculate_phase_balance,
    parse_panel_sheet,
    parse_panel_workbook,
    sync_panel_workbook_to_db,
    COL_L_DESC,
    COL_L_TRIP,
    COL_L_POLE,
    COL_L_KVA,
    COL_R_DESC,
    COL_R_TRIP,
    COL_R_POLE,
    COL_R_KVA,
    COL_VOLTAGE
)


class TestPanelScheduleSync(unittest.TestCase):
    def setUp(self):
        self.fd, self.temp_db = tempfile.mkstemp(suffix=".db")
        os.close(self.fd)
        init_db(self.temp_db)

    def tearDown(self):
        if os.path.exists(self.temp_db):
            try:
                os.remove(self.temp_db)
            except OSError:
                pass

    def test_calculate_phase_balance(self):
        # Balanced 3-phase loads: A=2000, B=2000, C=2000
        circuits = [
            {"circuitNumber": 1, "phasePole": "A", "connectedVA": 2000.0},
            {"circuitNumber": 3, "phasePole": "B", "connectedVA": 2000.0},
            {"circuitNumber": 5, "phasePole": "C", "connectedVA": 2000.0}
        ]
        summary = calculate_phase_balance(circuits, phase_count=3)
        self.assertEqual(summary["phaseAConnectedVA"], 2000.0)
        self.assertEqual(summary["phaseBConnectedVA"], 2000.0)
        self.assertEqual(summary["phaseCConnectedVA"], 2000.0)
        self.assertEqual(summary["totalConnectedVA"], 6000.0)
        self.assertEqual(summary["unbalancePercentage"], 0.0)

        # Unbalanced 3-phase loads: A=3000, B=2000, C=1000. Avg = 2000. (3000-1000)/2000 * 100 = 100%
        circuits_unbalanced = [
            {"circuitNumber": 1, "phasePole": "A", "connectedVA": 3000.0},
            {"circuitNumber": 3, "phasePole": "B", "connectedVA": 2000.0},
            {"circuitNumber": 5, "phasePole": "C", "connectedVA": 1000.0}
        ]
        summary_unbal = calculate_phase_balance(circuits_unbalanced, phase_count=3)
        self.assertEqual(summary_unbal["unbalancePercentage"], 100.0)

    def test_parse_and_sync_panel_workbook(self):
        # Create a sample workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "LP-1"

        ws["A3"] = "(E) PANEL 'LP-1'"
        ws[f"{COL_VOLTAGE}2"] = "120/208V"
        ws["G3"] = "225A"
        ws["K2"] = "4W"
        ws["K3"] = "3PH"
        ws["K4"] = "NEMA 1"
        ws["N2"] = "SURFACE"
        ws["N3"] = "10KAIC"

        # Row 8: Ckt 1 (L) and Ckt 2 (R)
        ws[f"{COL_L_DESC}8"] = "Lighting Corridor"
        ws[f"{COL_L_TRIP}8"] = "20"
        ws[f"{COL_L_POLE}8"] = "1"
        ws[f"{COL_L_KVA}8"] = "1.6"

        ws[f"{COL_R_DESC}8"] = "Receptacles Office"
        ws[f"{COL_R_TRIP}8"] = "20"
        ws[f"{COL_R_POLE}8"] = "1"
        ws[f"{COL_R_KVA}8"] = "1.8"

        temp_xlsx = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        temp_xlsx_path = temp_xlsx.name
        temp_xlsx.close()

        try:
            wb.save(temp_xlsx_path)

            proj = upsert_project("24-999", "Sync Test Project", "C:/Projects/24-999", db_path=self.temp_db)
            synced = sync_panel_workbook_to_db(proj["id"], temp_xlsx_path, db_path=self.temp_db)

            self.assertEqual(len(synced), 1)
            panel = synced[0]
            self.assertEqual(panel["panel_name"], "LP-1")
            self.assertEqual(panel["voltage"], "120/208V")
            self.assertEqual(panel["main_bus_amps"], 225)
            self.assertGreater(len(panel["circuits"]), 0)

            # Check circuit 1
            ckt1 = next(c for c in panel["circuits"] if c["circuit_number"] == 1)
            self.assertEqual(ckt1["load_description"], "Lighting Corridor")
            self.assertEqual(ckt1["connected_va"], 1600.0)

            # Check circuit 2
            ckt2 = next(c for c in panel["circuits"] if c["circuit_number"] == 2)
            self.assertEqual(ckt2["load_description"], "Receptacles Office")
            self.assertEqual(ckt2["connected_va"], 1800.0)

        finally:
            if os.path.exists(temp_xlsx_path):
                os.remove(temp_xlsx_path)


if __name__ == "__main__":
    unittest.main()
