# Lighting Plan Automation — Phase 1

Phase 1 is a review-first exchange between the `ElectricalCommands.LFSCommands`
AutoCAD bundle and the ACIES Light Fixture Scheduler.

## Workflow

1. Open and save the target electrical DWG.
2. Run `LIGHTPLANSCAN` (alias `LPSCAN`) in AutoCAD.
3. The first scan creates `ACIESLightingPlan.config.json` beside the DWG. If the
   scan misses or over-selects entities, edit its wildcard layer/block mappings
   and scan again.
4. AutoCAD writes `ACIESLightingPlan.snapshot.json` beside the DWG.
5. In ACIES, open **Light Fixture Scheduler**, select the project, and choose
   **Import CAD Scan**.
6. Review room/fixture/circuit totals and validation findings. Correct hard
   errors in the drawing or fixture schedule, rescan, and import again.
7. Choose **Prepare CAD Tags**. ACIES writes
   `ACIESLightingPlan.instructions.json` beside the scan.
8. Run `LIGHTPLANAPPLY` (alias `LPAPPLY`) in AutoCAD.

`LIGHTPLANAPPLY` identifies its own MText through the
`ACIES_LIGHTING_PLAN` registered application and a stable generation key. A
repeat run updates existing generated tags and removes only stale
ACIES-generated tags; it does not touch manually created text.

## Input expectations

- Fixtures are model-space block references matched by configured layer or
  block-name wildcard. Attributed blocks can also qualify when they contain a
  fixture mark plus panel, circuit, or wattage data.
- Rooms are closed model-space polylines on configured room-boundary layers.
- Room names and room types are taken from the nearest configured DBText/MText
  contained by each boundary.
- Default fixture attribute aliases include:

  | Field | Default aliases |
  | --- | --- |
  | Mark | `MARK`, `TYPE`, `FIXTURE_TYPE`, `FIXTURETYPE` |
  | Panel | `PANEL`, `PANEL_NAME`, `PANELNAME` |
  | Circuit | `CIRCUIT`, `CKT`, `CIRCUIT_NUMBER`, `CIRCUITNUMBER` |
  | Control zone | `CONTROL_ZONE`, `CONTROLZONE`, `CONTROL`, `ZONE` |
  | Wattage | `WATTS`, `WATTAGE`, `LOAD_WATTS`, `LOADWATTS` |

An occurrence wattage overrides its fixture-schedule wattage. Otherwise, ACIES
uses the wattage from the project fixture schedule.

## Validation behavior

Errors block tag preparation. Phase 1 errors include duplicate identifiers,
invalid room boundaries, missing fixture positions, missing fixture marks,
invalid wattage, and negative wattage.

Warnings remain reviewable but do not block preparation. They include unknown
fixture schedule marks, unassigned or overlapping rooms, missing wattage, and
missing panel/circuit assignments.

The instruction file records the DWG path and fingerprint. AutoCAD rejects an
instruction file prepared for a different or changed drawing and asks for a new
scan.

## Phase 1 boundaries

This phase generates fixture mark and power/control text tags, fixture counts,
circuit connected loads, and suggested circuit descriptions. It does not yet
place switches or sensors, calculate daylight zones, edit native panel schedule
tables, or make Title 24 compliance determinations.
