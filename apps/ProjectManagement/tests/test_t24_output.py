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
        self.assertEqual(res["data"]["totalSquareFeet"], 650.0)


if __name__ == "__main__":
    unittest.main()
