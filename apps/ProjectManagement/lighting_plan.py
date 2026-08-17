"""Phase 1 lighting-plan analysis shared by the ACIES desktop app.

The AutoCAD command exports geometry and block attributes.  This module keeps
the deterministic work (normalization, point-in-polygon assignment, fixture
validation, counts, circuit loads, and generation instructions) independent of
AutoCAD so it can be tested without a running CAD session.
"""

from __future__ import annotations

import hashlib
import json
import math
import re
from collections import defaultdict
from copy import deepcopy
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple


SCHEMA_VERSION = "1.0.0"
INSTRUCTION_SCHEMA_VERSION = "1.0.0"
DEFAULT_TAG_LAYER = "E-LITE-TAGS-ACIES"
DEFAULT_TAG_OFFSET = {"x": 1.5, "y": 1.5, "z": 0.0}

ATTRIBUTE_ALIASES = {
    "mark": ("MARK", "TYPE", "FIXTURE_TYPE", "FIXTURETYPE"),
    "panel": ("PANEL", "PANEL_NAME", "PANELNAME"),
    "circuit": ("CIRCUIT", "CKT", "CIRCUIT_NUMBER", "CIRCUITNUMBER"),
    "controlZone": ("CONTROL_ZONE", "CONTROLZONE", "CONTROL", "ZONE"),
    "watts": ("WATTS", "WATTAGE", "LOAD_WATTS", "LOADWATTS"),
}


class LightingPlanValidationError(ValueError):
    """Raised when an imported lighting-plan snapshot has an invalid shape."""


def _text(value: Any) -> str:
    return str(value if value is not None else "").strip()


def _lookup(mapping: Mapping[str, Any], *names: str, default: Any = None) -> Any:
    if not isinstance(mapping, Mapping):
        return default
    lowered = {str(key).lower(): value for key, value in mapping.items()}
    for name in names:
        if name in mapping:
            return mapping[name]
        lowered_name = name.lower()
        if lowered_name in lowered:
            return lowered[lowered_name]
    return default


def _finite_number(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    return result if math.isfinite(result) else None


def _rounded(value: float) -> float:
    return round(float(value), 4)


def _point(value: Any) -> Optional[Dict[str, float]]:
    if isinstance(value, Mapping):
        x = _finite_number(_lookup(value, "x", "X"))
        y = _finite_number(_lookup(value, "y", "Y"))
        z = _finite_number(_lookup(value, "z", "Z", default=0.0))
    elif isinstance(value, Sequence) and not isinstance(value, (str, bytes)):
        x = _finite_number(value[0]) if len(value) > 0 else None
        y = _finite_number(value[1]) if len(value) > 1 else None
        z = _finite_number(value[2]) if len(value) > 2 else 0.0
    else:
        return None
    if x is None or y is None:
        return None
    return {"x": _rounded(x), "y": _rounded(y), "z": _rounded(z or 0.0)}


def _vertices(value: Any) -> List[Dict[str, float]]:
    if not isinstance(value, Sequence) or isinstance(value, (str, bytes)):
        return []
    result = []
    for item in value:
        point = _point(item)
        if point is not None:
            result.append(point)
    if len(result) > 1 and result[0] == result[-1]:
        result.pop()
    return result


def polygon_area(vertices: Sequence[Mapping[str, float]]) -> float:
    if len(vertices) < 3:
        return 0.0
    doubled_area = 0.0
    for index, current in enumerate(vertices):
        following = vertices[(index + 1) % len(vertices)]
        doubled_area += current["x"] * following["y"]
        doubled_area -= following["x"] * current["y"]
    return abs(doubled_area) / 2.0


def _point_on_segment(
    point: Mapping[str, float],
    start: Mapping[str, float],
    end: Mapping[str, float],
    tolerance: float = 1e-8,
) -> bool:
    cross = (
        (point["y"] - start["y"]) * (end["x"] - start["x"])
        - (point["x"] - start["x"]) * (end["y"] - start["y"])
    )
    if abs(cross) > tolerance:
        return False
    dot = (
        (point["x"] - start["x"]) * (end["x"] - start["x"])
        + (point["y"] - start["y"]) * (end["y"] - start["y"])
    )
    if dot < -tolerance:
        return False
    squared_length = (
        (end["x"] - start["x"]) ** 2 + (end["y"] - start["y"]) ** 2
    )
    return dot <= squared_length + tolerance


def point_in_polygon(
    point: Mapping[str, float], vertices: Sequence[Mapping[str, float]]
) -> bool:
    """Return True for points inside or on the edge of a simple polygon."""
    if len(vertices) < 3:
        return False
    inside = False
    x = point["x"]
    y = point["y"]
    for index, start in enumerate(vertices):
        end = vertices[(index + 1) % len(vertices)]
        if _point_on_segment(point, start, end):
            return True
        intersects = (start["y"] > y) != (end["y"] > y)
        if not intersects:
            continue
        intersection_x = (
            (end["x"] - start["x"])
            * (y - start["y"])
            / (end["y"] - start["y"])
            + start["x"]
        )
        if x < intersection_x:
            inside = not inside
    return inside


def _attributes(value: Any) -> Dict[str, str]:
    if not isinstance(value, Mapping):
        return {}
    return {
        _text(key).upper(): _text(item)
        for key, item in value.items()
        if _text(key)
    }


def _attribute_value(attributes: Mapping[str, str], field: str) -> str:
    for alias in ATTRIBUTE_ALIASES[field]:
        value = _text(attributes.get(alias))
        if value:
            return value
    return ""


def normalize_snapshot(snapshot: Any) -> Dict[str, Any]:
    if not isinstance(snapshot, Mapping):
        raise LightingPlanValidationError(
            "Lighting plan snapshot must contain a JSON object."
        )

    drawing_value = _lookup(snapshot, "drawing", "Drawing", default={})
    drawing = drawing_value if isinstance(drawing_value, Mapping) else {}
    rooms_value = _lookup(snapshot, "rooms", "Rooms", default=[])
    fixtures_value = _lookup(snapshot, "fixtures", "Fixtures", default=[])
    if not isinstance(rooms_value, list):
        raise LightingPlanValidationError("Snapshot rooms must be a JSON array.")
    if not isinstance(fixtures_value, list):
        raise LightingPlanValidationError("Snapshot fixtures must be a JSON array.")

    normalized_rooms = []
    for index, raw in enumerate(rooms_value, start=1):
        if not isinstance(raw, Mapping):
            raise LightingPlanValidationError(f"Room #{index} must be a JSON object.")
        boundary = _vertices(_lookup(raw, "boundary", "vertices", "Boundary", "Vertices"))
        handle = _text(_lookup(raw, "handle", "Handle"))
        room_id = _text(_lookup(raw, "id", "roomId", "Id", "RoomId")) or handle
        if not room_id:
            room_id = f"room-{index}"
        reported_area = _finite_number(_lookup(raw, "area", "Area"))
        normalized_rooms.append(
            {
                "id": room_id,
                "handle": handle,
                "name": _text(_lookup(raw, "name", "roomName", "Name", "RoomName")),
                "roomType": _text(_lookup(raw, "roomType", "type", "RoomType", "Type")),
                "layer": _text(_lookup(raw, "layer", "Layer")),
                "boundary": boundary,
                "area": _rounded(reported_area if reported_area is not None else polygon_area(boundary)),
            }
        )

    normalized_fixtures = []
    for index, raw in enumerate(fixtures_value, start=1):
        if not isinstance(raw, Mapping):
            raise LightingPlanValidationError(
                f"Fixture #{index} must be a JSON object."
            )
        attributes = _attributes(_lookup(raw, "attributes", "Attributes", default={}))
        handle = _text(_lookup(raw, "handle", "Handle"))
        fixture_id = _text(_lookup(raw, "id", "fixtureId", "Id", "FixtureId")) or handle
        if not fixture_id:
            fixture_id = f"fixture-{index}"
        normalized_fixtures.append(
            {
                "id": fixture_id,
                "handle": handle,
                "blockName": _text(_lookup(raw, "blockName", "name", "BlockName", "Name")),
                "layer": _text(_lookup(raw, "layer", "Layer")),
                "position": _point(_lookup(raw, "position", "Position")),
                "rotation": _rounded(_finite_number(_lookup(raw, "rotation", "Rotation")) or 0.0),
                "attributes": attributes,
                "mark": _text(_lookup(raw, "mark", "Mark")) or _attribute_value(attributes, "mark"),
                "panel": _text(_lookup(raw, "panel", "Panel")) or _attribute_value(attributes, "panel"),
                "circuit": _text(_lookup(raw, "circuit", "Circuit")) or _attribute_value(attributes, "circuit"),
                "controlZone": _text(_lookup(raw, "controlZone", "ControlZone")) or _attribute_value(attributes, "controlZone"),
                "wattsRaw": _lookup(raw, "watts", "wattage", "Watts", "Wattage", default=None)
                if _lookup(raw, "watts", "wattage", "Watts", "Wattage", default=None) not in (None, "")
                else _attribute_value(attributes, "watts"),
                "roomId": _text(_lookup(raw, "roomId", "RoomId")),
            }
        )

    config_value = _lookup(snapshot, "config", "Config", default={})
    return {
        "schemaVersion": _text(_lookup(snapshot, "schemaVersion", "SchemaVersion")) or SCHEMA_VERSION,
        "drawing": {
            "path": _text(_lookup(drawing, "path", "Path")),
            "name": _text(_lookup(drawing, "name", "Name")),
            "fingerprint": _text(_lookup(drawing, "fingerprint", "Fingerprint")),
            "units": _text(_lookup(drawing, "units", "Units")),
            "scannedAtUtc": _text(_lookup(drawing, "scannedAtUtc", "ScannedAtUtc")),
        },
        "config": deepcopy(config_value) if isinstance(config_value, Mapping) else {},
        "rooms": normalized_rooms,
        "fixtures": normalized_fixtures,
    }


def _normalize_schedule(schedule_rows: Any) -> Dict[str, Dict[str, Any]]:
    rows = schedule_rows if isinstance(schedule_rows, list) else []
    result: Dict[str, Dict[str, Any]] = {}
    for raw in rows:
        if not isinstance(raw, Mapping):
            continue
        mark = _text(_lookup(raw, "mark", "Mark")).upper()
        if not mark:
            continue
        watts_raw = _lookup(raw, "watts", "Watts", "wattage", "Wattage")
        result[mark] = {
            "mark": _text(_lookup(raw, "mark", "Mark")),
            "description": _text(_lookup(raw, "description", "Description")),
            "watts": _finite_number(watts_raw),
            "wattsRaw": watts_raw,
        }
    return result


def _natural_key(value: str) -> List[Any]:
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", value)]


def _fingerprint_payload(snapshot: Mapping[str, Any]) -> str:
    encoded = json.dumps(snapshot, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _tag_text(fixture: Mapping[str, Any]) -> str:
    mark = _text(fixture.get("mark"))
    panel = _text(fixture.get("panel"))
    circuit = _text(fixture.get("circuit"))
    control = _text(fixture.get("controlZone"))
    power = "-".join(part for part in (panel, circuit) if part)
    circuit_line = f"{power}{control}" if power else control
    return "\\P".join(part for part in (mark, circuit_line) if part)


def create_generation_instructions(
    analysis: Mapping[str, Any],
    *,
    tag_layer: str = DEFAULT_TAG_LAYER,
    tag_offset: Optional[Mapping[str, Any]] = None,
) -> Dict[str, Any]:
    offset = _point(tag_offset or DEFAULT_TAG_OFFSET) or dict(DEFAULT_TAG_OFFSET)
    instructions = []
    for fixture in analysis.get("fixtures", []):
        position = fixture.get("position")
        handle = _text(fixture.get("handle"))
        mark = _text(fixture.get("mark"))
        text = _tag_text(fixture)
        if not isinstance(position, Mapping) or not handle or not mark or not text:
            continue
        instructions.append(
            {
                "generationKey": f"fixture:{handle}:tag:v1",
                "sourceHandle": handle,
                "fixtureId": fixture.get("id"),
                "mark": fixture.get("mark"),
                "panel": fixture.get("panel"),
                "circuit": fixture.get("circuit"),
                "controlZone": fixture.get("controlZone"),
                "roomId": fixture.get("roomId"),
                "roomName": fixture.get("roomName"),
                "text": text,
                "layer": _text(tag_layer) or DEFAULT_TAG_LAYER,
                "position": {
                    "x": _rounded(position["x"] + offset["x"]),
                    "y": _rounded(position["y"] + offset["y"]),
                    "z": _rounded(position.get("z", 0.0) + offset.get("z", 0.0)),
                },
            }
        )

    drawing = analysis.get("drawing", {})
    return {
        "schemaVersion": INSTRUCTION_SCHEMA_VERSION,
        "kind": "acies-lighting-plan-generation",
        "sourceFingerprint": analysis.get("sourceFingerprint", ""),
        "drawingPath": drawing.get("path", ""),
        "tagLayer": _text(tag_layer) or DEFAULT_TAG_LAYER,
        "tags": instructions,
    }


def analyze_lighting_plan(
    snapshot: Any,
    schedule_rows: Any = None,
    *,
    tag_layer: Optional[str] = None,
    tag_offset: Optional[Mapping[str, Any]] = None,
) -> Dict[str, Any]:
    normalized = normalize_snapshot(snapshot)
    snapshot_config = normalized.get("config", {})
    effective_tag_layer = (
        _text(tag_layer)
        or _text(_lookup(snapshot_config, "tagLayer", "TagLayer"))
        or DEFAULT_TAG_LAYER
    )
    effective_tag_offset = tag_offset
    if effective_tag_offset is None:
        configured_offset = _lookup(snapshot_config, "tagOffset", "TagOffset")
        if isinstance(configured_offset, Mapping):
            effective_tag_offset = configured_offset
    schedule = _normalize_schedule(schedule_rows)
    warnings: List[Dict[str, Any]] = []

    def warn(
        code: str,
        message: str,
        *,
        severity: str = "warning",
        entity_type: str = "drawing",
        entity_id: str = "",
    ) -> None:
        warnings.append(
            {
                "code": code,
                "severity": severity,
                "message": message,
                "entityType": entity_type,
                "entityId": entity_id,
            }
        )

    if normalized["schemaVersion"] != SCHEMA_VERSION:
        warn(
            "snapshot.schema_version",
            f"Snapshot schema {normalized['schemaVersion']} is not the supported {SCHEMA_VERSION}; continuing with best effort.",
        )

    valid_rooms = []
    seen_room_ids = set()
    for room in normalized["rooms"]:
        room_id = room["id"]
        if room_id in seen_room_ids:
            warn(
                "room.duplicate_id",
                f"Room ID {room_id} appears more than once.",
                severity="error",
                entity_type="room",
                entity_id=room_id,
            )
        seen_room_ids.add(room_id)
        if len(room["boundary"]) < 3 or room["area"] <= 0:
            warn(
                "room.invalid_boundary",
                f"Room {room.get('name') or room_id} does not have a usable closed boundary.",
                severity="error",
                entity_type="room",
                entity_id=room_id,
            )
            continue
        valid_rooms.append(room)

    if not normalized["rooms"]:
        warn("drawing.no_rooms", "No room boundaries were found in the scan.")
    if not normalized["fixtures"]:
        warn("drawing.no_fixtures", "No light fixture blocks were found in the scan.")

    seen_fixture_ids = set()
    analyzed_fixtures = []
    for fixture in normalized["fixtures"]:
        fixture = deepcopy(fixture)
        fixture_id = fixture["id"]
        if fixture_id in seen_fixture_ids:
            warn(
                "fixture.duplicate_id",
                f"Fixture ID {fixture_id} appears more than once.",
                severity="error",
                entity_type="fixture",
                entity_id=fixture_id,
            )
        seen_fixture_ids.add(fixture_id)

        position = fixture.get("position")
        containing_rooms = []
        if position is None:
            warn(
                "fixture.missing_position",
                "Fixture has no valid insertion position.",
                severity="error",
                entity_type="fixture",
                entity_id=fixture_id,
            )
        else:
            containing_rooms = [
                room for room in valid_rooms if point_in_polygon(position, room["boundary"])
            ]

        selected_room = min(containing_rooms, key=lambda room: room["area"]) if containing_rooms else None
        if len(containing_rooms) > 1:
            warn(
                "fixture.overlapping_rooms",
                f"Fixture falls inside {len(containing_rooms)} room boundaries; the smallest room was selected.",
                entity_type="fixture",
                entity_id=fixture_id,
            )
        if selected_room is None:
            fixture["roomId"] = ""
            fixture["roomName"] = ""
            fixture["roomType"] = ""
            if position is not None:
                warn(
                    "fixture.unassigned_room",
                    "Fixture is not inside a valid room boundary.",
                    entity_type="fixture",
                    entity_id=fixture_id,
                )
        else:
            fixture["roomId"] = selected_room["id"]
            fixture["roomName"] = selected_room["name"]
            fixture["roomType"] = selected_room["roomType"]

        mark_key = fixture["mark"].upper()
        schedule_item = schedule.get(mark_key)
        if not mark_key:
            warn(
                "fixture.missing_mark",
                "Fixture has no schedule mark.",
                severity="error",
                entity_type="fixture",
                entity_id=fixture_id,
            )
        elif schedule_item is None:
            warn(
                "fixture.unknown_mark",
                f"Fixture mark {fixture['mark']} is not present in the project fixture schedule.",
                entity_type="fixture",
                entity_id=fixture_id,
            )

        occurrence_watts = _finite_number(fixture.get("wattsRaw"))
        raw_watts_supplied = fixture.get("wattsRaw") not in (None, "")
        if raw_watts_supplied and occurrence_watts is None:
            warn(
                "fixture.invalid_watts",
                f"Fixture wattage {fixture.get('wattsRaw')!r} is not numeric.",
                severity="error",
                entity_type="fixture",
                entity_id=fixture_id,
            )
        fixture["watts"] = occurrence_watts
        fixture["wattsSource"] = "fixture" if occurrence_watts is not None else ""
        if fixture["watts"] is None and schedule_item is not None:
            fixture["watts"] = schedule_item["watts"]
            fixture["wattsSource"] = "schedule" if schedule_item["watts"] is not None else ""
        if fixture["watts"] is None:
            warn(
                "fixture.missing_watts",
                f"Fixture {fixture['mark'] or fixture_id} has no numeric wattage.",
                entity_type="fixture",
                entity_id=fixture_id,
            )
        else:
            fixture["watts"] = _rounded(fixture["watts"])
            if fixture["watts"] < 0:
                warn(
                    "fixture.negative_watts",
                    "Fixture wattage cannot be negative.",
                    severity="error",
                    entity_type="fixture",
                    entity_id=fixture_id,
                )

        if not fixture["panel"]:
            warn(
                "fixture.missing_panel",
                "Fixture has no panel assignment.",
                entity_type="fixture",
                entity_id=fixture_id,
            )
        if not fixture["circuit"]:
            warn(
                "fixture.missing_circuit",
                "Fixture has no circuit assignment.",
                entity_type="fixture",
                entity_id=fixture_id,
            )
        analyzed_fixtures.append(fixture)

    fixture_groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    circuit_groups: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
    for fixture in analyzed_fixtures:
        fixture_groups[fixture["mark"] or "(UNMARKED)"].append(fixture)
        circuit_groups[(fixture["panel"], fixture["circuit"])].append(fixture)

    fixture_counts = []
    for mark in sorted(fixture_groups, key=_natural_key):
        fixtures = fixture_groups[mark]
        schedule_item = schedule.get(mark.upper(), {})
        known_watts = [item["watts"] for item in fixtures if item["watts"] is not None]
        fixture_counts.append(
            {
                "mark": mark,
                "description": schedule_item.get("description", ""),
                "quantity": len(fixtures),
                "totalWatts": _rounded(sum(known_watts)),
                "fixturesWithoutWatts": len(fixtures) - len(known_watts),
            }
        )

    circuits = []
    for (panel, circuit), fixtures in sorted(
        circuit_groups.items(), key=lambda item: (_natural_key(item[0][0]), _natural_key(item[0][1]))
    ):
        room_names = sorted(
            {item["roomName"] for item in fixtures if item["roomName"]},
            key=_natural_key,
        )
        known_watts = [item["watts"] for item in fixtures if item["watts"] is not None]
        circuits.append(
            {
                "panel": panel,
                "circuit": circuit,
                "fixtureCount": len(fixtures),
                "totalWatts": _rounded(sum(known_watts)),
                "fixturesWithoutWatts": len(fixtures) - len(known_watts),
                "rooms": room_names,
                "description": f"LIGHTING - {', '.join(room_names)}" if room_names else "LIGHTING",
            }
        )

    source_fingerprint = normalized["drawing"].get("fingerprint") or _fingerprint_payload(normalized)
    result = {
        "schemaVersion": SCHEMA_VERSION,
        "sourceFingerprint": source_fingerprint,
        "drawing": normalized["drawing"],
        "rooms": normalized["rooms"],
        "fixtures": analyzed_fixtures,
        "fixtureCounts": fixture_counts,
        "circuits": circuits,
        "warnings": warnings,
        "summary": {
            "roomCount": len(normalized["rooms"]),
            "fixtureCount": len(analyzed_fixtures),
            "fixtureTypeCount": len(fixture_counts),
            "circuitCount": len(circuits),
            "totalWatts": _rounded(
                sum(item["watts"] for item in analyzed_fixtures if item["watts"] is not None)
            ),
            "errorCount": sum(item["severity"] == "error" for item in warnings),
            "warningCount": sum(item["severity"] == "warning" for item in warnings),
        },
    }
    result["canGenerate"] = result["summary"]["errorCount"] == 0
    result["instructions"] = create_generation_instructions(
        result, tag_layer=effective_tag_layer, tag_offset=effective_tag_offset
    )
    return result
