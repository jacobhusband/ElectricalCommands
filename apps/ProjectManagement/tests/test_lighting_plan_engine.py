import unittest

from lighting_plan import (
    LightingPlanValidationError,
    analyze_lighting_plan,
    normalize_snapshot,
    point_in_polygon,
)


def _room(room_id="R1", name="OFFICE 101", size=10, origin=(0, 0)):
    x, y = origin
    return {
        "id": room_id,
        "handle": room_id,
        "name": name,
        "roomType": "Office",
        "boundary": [
            {"x": x, "y": y},
            {"x": x + size, "y": y},
            {"x": x + size, "y": y + size},
            {"x": x, "y": y + size},
        ],
    }


def _fixture(fixture_id="F1", position=(5, 5), **overrides):
    result = {
        "id": fixture_id,
        "handle": fixture_id,
        "blockName": "ACIES-LIGHT",
        "position": {"x": position[0], "y": position[1], "z": 0},
        "attributes": {
            "MARK": "L1",
            "PANEL": "LP1",
            "CIRCUIT": "7",
            "CONTROL_ZONE": "a",
        },
    }
    result.update(overrides)
    return result


class LightingPlanGeometryTests(unittest.TestCase):
    def test_point_in_polygon_includes_edges(self):
        boundary = _room()["boundary"]
        self.assertTrue(point_in_polygon({"x": 5, "y": 5}, boundary))
        self.assertTrue(point_in_polygon({"x": 0, "y": 5}, boundary))
        self.assertFalse(point_in_polygon({"x": 11, "y": 5}, boundary))


class LightingPlanNormalizationTests(unittest.TestCase):
    def test_attribute_aliases_populate_fixture_fields(self):
        normalized = normalize_snapshot(
            {
                "Rooms": [],
                "Fixtures": [
                    {
                        "Handle": "A1",
                        "Position": [1, 2],
                        "Attributes": {
                            "FIXTURE_TYPE": "L2",
                            "PANEL_NAME": "LP2",
                            "CKT": "11",
                            "ZONE": "b",
                            "WATTAGE": "12.5",
                        },
                    }
                ],
            }
        )
        fixture = normalized["fixtures"][0]
        self.assertEqual("L2", fixture["mark"])
        self.assertEqual("LP2", fixture["panel"])
        self.assertEqual("11", fixture["circuit"])
        self.assertEqual("b", fixture["controlZone"])
        self.assertEqual("12.5", fixture["wattsRaw"])

    def test_rejects_non_array_fixtures(self):
        with self.assertRaises(LightingPlanValidationError):
            normalize_snapshot({"rooms": [], "fixtures": {}})


class LightingPlanAnalysisTests(unittest.TestCase):
    def test_assigns_rooms_rolls_up_fixture_counts_and_circuit_watts(self):
        snapshot = {
            "schemaVersion": "1.0.0",
            "drawing": {"path": r"C:\\Projects\\260001\\E1.1.dwg", "fingerprint": "abc"},
            "rooms": [_room()],
            "fixtures": [
                _fixture("F1", (2, 2)),
                _fixture("F2", (8, 8), attributes={"MARK": "L1", "PANEL": "LP1", "CIRCUIT": "7"}),
            ],
        }
        analysis = analyze_lighting_plan(
            snapshot,
            [{"mark": "L1", "description": "4 IN LED", "watts": "11"}],
        )

        self.assertTrue(analysis["canGenerate"])
        self.assertEqual("R1", analysis["fixtures"][0]["roomId"])
        self.assertEqual("OFFICE 101", analysis["fixtures"][1]["roomName"])
        self.assertEqual(
            {
                "mark": "L1",
                "description": "4 IN LED",
                "quantity": 2,
                "totalWatts": 22.0,
                "fixturesWithoutWatts": 0,
            },
            analysis["fixtureCounts"][0],
        )
        self.assertEqual(22.0, analysis["circuits"][0]["totalWatts"])
        self.assertEqual("LIGHTING - OFFICE 101", analysis["circuits"][0]["description"])
        self.assertEqual("L1\\PLP1-7a", analysis["instructions"]["tags"][0]["text"])
        self.assertEqual("fixture:F1:tag:v1", analysis["instructions"]["tags"][0]["generationKey"])

    def test_occurrence_watts_override_schedule_watts(self):
        fixture = _fixture(attributes={
            "MARK": "L1",
            "PANEL": "LP1",
            "CIRCUIT": "7",
            "WATTS": "15",
        })
        analysis = analyze_lighting_plan(
            {"rooms": [_room()], "fixtures": [fixture]},
            [{"mark": "L1", "watts": "11"}],
        )
        self.assertEqual(15.0, analysis["fixtures"][0]["watts"])
        self.assertEqual("fixture", analysis["fixtures"][0]["wattsSource"])

    def test_snapshot_tag_configuration_flows_into_generation_instructions(self):
        analysis = analyze_lighting_plan(
            {
                "config": {
                    "tagLayer": "E-ANNO-LITE-TAGS",
                    "tagOffset": {"x": 2, "y": -1, "z": 0},
                },
                "rooms": [_room()],
                "fixtures": [_fixture(position=(3, 4))],
            },
            [{"mark": "L1", "watts": 10}],
        )
        instruction = analysis["instructions"]
        self.assertEqual("E-ANNO-LITE-TAGS", instruction["tagLayer"])
        self.assertEqual(
            {"x": 5.0, "y": 3.0, "z": 0.0}, instruction["tags"][0]["position"]
        )

    def test_smallest_overlapping_room_is_selected_and_warned(self):
        analysis = analyze_lighting_plan(
            {
                "rooms": [
                    _room("R1", "OPEN OFFICE", 20),
                    _room("R2", "HUDDLE 102", 5, origin=(2, 2)),
                ],
                "fixtures": [_fixture(position=(3, 3))],
            },
            [{"mark": "L1", "watts": 10}],
        )
        self.assertEqual("R2", analysis["fixtures"][0]["roomId"])
        self.assertIn(
            "fixture.overlapping_rooms",
            {item["code"] for item in analysis["warnings"]},
        )

    def test_errors_block_generation_readiness_and_invalid_tags_are_skipped(self):
        analysis = analyze_lighting_plan(
            {
                "rooms": [_room()],
                "fixtures": [
                    {
                        "id": "F1",
                        "handle": "F1",
                        "position": {"x": 2, "y": 2},
                        "attributes": {"PANEL": "LP1", "CIRCUIT": "1"},
                    }
                ],
            },
            [],
        )
        self.assertFalse(analysis["canGenerate"])
        self.assertIn("fixture.missing_mark", {item["code"] for item in analysis["warnings"]})
        self.assertEqual([], analysis["instructions"]["tags"])


if __name__ == "__main__":
    unittest.main()
