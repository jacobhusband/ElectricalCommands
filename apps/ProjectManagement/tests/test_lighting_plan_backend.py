import json
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch


def _ensure_optional_import_stubs():
    try:
        import webview  # noqa: F401
    except Exception:
        module = types.ModuleType("webview")
        module.windows = []
        module.create_window = lambda *args, **kwargs: None
        module.start = lambda *args, **kwargs: None
        sys.modules["webview"] = module

    try:
        from google import genai as _genai  # noqa: F401
    except Exception:
        google = sys.modules.setdefault("google", types.ModuleType("google"))
        google.__path__ = []
        genai = types.ModuleType("google.genai")
        genai_types = types.ModuleType("google.genai.types")
        genai.types = genai_types
        google.genai = genai
        sys.modules["google.genai"] = genai
        sys.modules["google.genai.types"] = genai_types


_ensure_optional_import_stubs()

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import main as main_module  # noqa: E402
from main import Api  # noqa: E402


def _snapshot():
    return {
        "schemaVersion": "1.0.0",
        "drawing": {"path": r"C:\Projects\260001 Example\E1.1.dwg", "fingerprint": "dwg-1"},
        "rooms": [
            {
                "id": "R1",
                "handle": "100",
                "name": "OFFICE 101",
                "roomType": "Office",
                "boundary": [[0, 0], [10, 0], [10, 10], [0, 10]],
            }
        ],
        "fixtures": [
            {
                "id": "F1",
                "handle": "200",
                "blockName": "ACIES-LIGHT",
                "position": [5, 5, 0],
                "attributes": {
                    "MARK": "L1",
                    "PANEL": "LP1",
                    "CIRCUIT": "3",
                    "CONTROL_ZONE": "a",
                },
            }
        ],
    }


class LightingPlanBackendTests(unittest.TestCase):
    def setUp(self):
        self.api = Api.__new__(Api)

    def test_import_analyzes_persists_and_exports_instructions(self):
        with tempfile.TemporaryDirectory(prefix="acies-lighting-plan-") as temp_dir:
            temp_path = Path(temp_dir)
            db_path = temp_path / "lighting_plans.db"
            schedule_db_path = temp_path / "lighting_schedules.db"
            snapshot_path = temp_path / "ACIESLightingPlan.snapshot.json"
            snapshot_path.write_text(json.dumps(_snapshot()), encoding="utf-8")

            with (
                patch.object(main_module, "LIGHTING_PLAN_DB_FILE", str(db_path)),
                patch.object(main_module, "LIGHTING_SCHEDULE_DB_FILE", str(schedule_db_path)),
            ):
                main_module._save_lighting_schedule_record(
                    "260001",
                    {"schedule": {"rows": [{"mark": "L1", "watts": "12"}]}},
                )
                imported = self.api.import_lighting_plan_snapshot("260001", str(snapshot_path))
                loaded = self.api.get_lighting_plan_record("260001")
                exported = self.api.export_lighting_plan_instructions("260001")

            self.assertEqual("success", imported["status"])
            analysis = imported["data"]["analysis"]
            self.assertEqual(1, analysis["summary"]["fixtureCount"])
            self.assertEqual(12.0, analysis["summary"]["totalWatts"])
            self.assertEqual("R1", analysis["fixtures"][0]["roomId"])
            self.assertTrue(loaded["exists"])
            self.assertEqual("success", exported["status"])
            self.assertEqual(1, exported["tagCount"])
            instruction_path = temp_path / "ACIESLightingPlan.instructions.json"
            self.assertEqual(str(instruction_path), exported["path"])
            instructions = json.loads(instruction_path.read_text(encoding="utf-8"))
            self.assertEqual("200", instructions["tags"][0]["sourceHandle"])

    def test_export_is_blocked_when_scan_contains_validation_errors(self):
        with tempfile.TemporaryDirectory(prefix="acies-lighting-plan-errors-") as temp_dir:
            temp_path = Path(temp_dir)
            snapshot = _snapshot()
            snapshot["fixtures"][0]["attributes"].pop("MARK")
            snapshot_path = temp_path / "ACIESLightingPlan.snapshot.json"
            snapshot_path.write_text(json.dumps(snapshot), encoding="utf-8")

            with (
                patch.object(main_module, "LIGHTING_PLAN_DB_FILE", str(temp_path / "plans.db")),
                patch.object(main_module, "LIGHTING_SCHEDULE_DB_FILE", str(temp_path / "schedules.db")),
            ):
                imported = self.api.import_lighting_plan_snapshot("260002", str(snapshot_path))
                exported = self.api.export_lighting_plan_instructions("260002")

            self.assertEqual("success", imported["status"])
            self.assertFalse(imported["data"]["analysis"]["canGenerate"])
            self.assertEqual("error", exported["status"])
            self.assertIn("validation errors", exported["message"])


if __name__ == "__main__":
    unittest.main()
