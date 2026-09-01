import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path


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

APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from main import Api


class TestT24OutputJson(unittest.TestCase):
    def setUp(self):
        self.api = Api()
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        if os.path.exists(self.temp_dir):
            for file in os.listdir(self.temp_dir):
                os.remove(os.path.join(self.temp_dir, file))
            os.rmdir(self.temp_dir)

    def test_read_t24_output_json_with_total(self):
        t24_file = os.path.join(self.temp_dir, "T24Output.json")
        sample_data = [
            {"RoomType": "OFFICE", "SquareFeet": 1250.5},
            {"RoomType": "BREAK ROOM", "SquareFeet": 300.0},
            {"RoomType": "CORRIDOR", "SquareFeet": 450.25},
            {"RoomType": "TOTAL", "SquareFeet": 2000.75}
        ]
        with open(t24_file, "w", encoding="utf-8") as f:
            json.dump(sample_data, f, indent=2)

        res = self.api.read_t24_output_json(t24_file)
        self.assertEqual(res["status"], "success")
        self.assertEqual(len(res["data"]["rows"]), 3)
        self.assertEqual(res["data"]["rows"][0]["roomType"], "OFFICE")
        self.assertEqual(res["data"]["rows"][0]["squareFeet"], 1250.5)
        self.assertEqual(res["data"]["rows"][1]["roomType"], "BREAK ROOM")
        self.assertEqual(res["data"]["rows"][1]["squareFeet"], 300.0)
        self.assertEqual(res["data"]["rows"][2]["roomType"], "CORRIDOR")
        self.assertEqual(res["data"]["rows"][2]["squareFeet"], 450.25)
        self.assertEqual(res["data"]["totalSquareFeet"], 2000.75)

    def test_read_t24_output_json_without_total(self):
        t24_file = os.path.join(self.temp_dir, "T24Output.json")
        sample_data = [
            {"RoomType": "OFFICE", "SquareFeet": 500.0},
            {"RoomType": "RESTROOM", "SquareFeet": 150.0}
        ]
        with open(t24_file, "w", encoding="utf-8") as f:
            json.dump(sample_data, f, indent=2)

        res = self.api.read_t24_output_json(t24_file)
        self.assertEqual(res["status"], "success")
        self.assertEqual(len(res["data"]["rows"]), 2)
    def test_read_arealabel_json_with_total(self):
        area_file = os.path.join(self.temp_dir, "AreaLabel.json")
        sample_data = [
            {"RoomType": "CONFERENCE", "SquareFeet": 420.0},
            {"RoomType": "OFFICE", "SquareFeet": 850.5},
            {"RoomType": "TOTAL", "SquareFeet": 1270.5}
        ]
        with open(area_file, "w", encoding="utf-8") as f:
            json.dump(sample_data, f, indent=2)

        res = self.api.read_t24_output_json(area_file)
        self.assertEqual(res["status"], "success")
        self.assertEqual(len(res["data"]["rows"]), 2)
        self.assertEqual(res["data"]["rows"][0]["roomType"], "CONFERENCE")
        self.assertEqual(res["data"]["rows"][0]["squareFeet"], 420.0)
        self.assertEqual(res["data"]["rows"][1]["roomType"], "OFFICE")
        self.assertEqual(res["data"]["rows"][1]["squareFeet"], 850.5)
        self.assertEqual(res["data"]["totalSquareFeet"], 1270.5)

    def test_read_arealabel_json_with_room_name_key(self):
        area_file = os.path.join(self.temp_dir, "AreaLabel.json")
        sample_data = [
            {"RoomName": "LOBBY", "SquareFeet": 600.0},
            {"RoomName": "STORAGE", "SquareFeet": 100.0}
        ]
        with open(area_file, "w", encoding="utf-8") as f:
            json.dump(sample_data, f, indent=2)

        res = self.api.read_t24_output_json(area_file)
        self.assertEqual(res["status"], "success")
        self.assertEqual(len(res["data"]["rows"]), 2)
        self.assertEqual(res["data"]["rows"][0]["roomType"], "LOBBY")
        self.assertEqual(res["data"]["rows"][0]["squareFeet"], 600.0)
        self.assertEqual(res["data"]["totalSquareFeet"], 700.0)

    def test_read_invalid_filename_rejected(self):
        invalid_file = os.path.join(self.temp_dir, "OtherData.json")
        with open(invalid_file, "w", encoding="utf-8") as f:
            json.dump([{"RoomType": "OFFICE", "SquareFeet": 100.0}], f)

        res = self.api.read_t24_output_json(invalid_file)
        self.assertEqual(res["status"], "error")
        self.assertIn("Expected file name", res["message"])


if __name__ == "__main__":
    unittest.main()
