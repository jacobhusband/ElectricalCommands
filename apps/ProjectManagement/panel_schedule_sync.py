"""
ACIES Engineering - Panel Schedule Parser & Sync Engine
Handles extraction of panel schedule workbooks (.xlsx), JSON schema compliance,
NEC phase load balancing calculations, diagnostics generation, and SQLite synchronization.
"""

import os
import re
import json
import math
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple

import openpyxl
from apps.ProjectManagement.database import upsert_panel_schedule, get_panel_schedule

# Column mapping conforming to ACIES Panel Schedule Excel Template
COL_L_NOTE = "B"
COL_L_TYPE = "C"
COL_L_POLE = "D"
COL_L_TRIP = "E"
COL_L_DESC = "F"
COL_L_KVA = "I"

COL_VOLTAGE = "G"

COL_R_KVA = "K"
COL_R_DESC = "L"
COL_R_POLE = "O"
COL_R_TRIP = "P"
COL_R_TYPE = "Q"
COL_R_NOTE = "R"

START_ROW = 8
MAX_ROW = 28


def clean_cell_str(val: Any) -> str:
    """Helper to clean and normalize cell string values."""
    if val is None:
        return ""
    return str(val).strip()


def parse_numeric(val: Any, default: float = 0.0) -> float:
    """Safely extracts numeric float values from cell strings or numbers."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    text = str(val).replace(",", "").strip()
    match = re.search(r"[-+]?\d+(?:\.\d+)?", text)
    if match:
        try:
            return float(match.group(0))
        except ValueError:
            return default
    return default


def parse_integer(val: Any, default: int = 0) -> int:
    """Safely extracts integer values."""
    return int(parse_numeric(val, float(default)))


def map_load_type(type_code: str, description: str) -> str:
    """Maps shorthand load type codes or descriptions to canonical schema enum."""
    code = type_code.upper().strip()
    desc = description.upper().strip()
    
    if any(x in desc for x in ["SPARE"]):
        return "SPARE"
    if any(x in desc for x in ["SPACE"]):
        return "SPACE"
    if code == "L" or "LIGHT" in desc or "LTG" in desc:
        return "LIGHTING_CONTINUOUS"
    if code in ("R", "G") or "RECEP" in desc or "PLUG" in desc or "OUTLET" in desc:
        return "RECEPTACLE_NON_CONTINUOUS"
    if code == "M" or "MOTOR" in desc or "PUMP" in desc or "FAN" in desc:
        return "MOTOR"
    if code in ("H", "AC", "HVAC") or any(x in desc for x in ["A/C", "AC", "HVAC", "AHU", "RTU", "CONDENSER"]):
        return "HVAC_CONTINUOUS"
    if code == "K" or "KITCHEN" in desc or "HOOD" in desc or "OVEN" in desc:
        return "KITCHEN_EQUIPMENT"
    if "HEAT" in desc or "HEATER" in desc:
        return "ELECTRIC_HEATING"
    return "RECEPTACLE_NON_CONTINUOUS"


def calculate_phase_balance(circuits: List[Dict[str, Any]], phase_count: int = 3) -> Dict[str, Any]:
    """
    Calculates connected VA per phase (A, B, C) and phase unbalance percentage per standard formula:
    Unbalance % = (Max(VA_A, VA_B, VA_C) - Min(VA_A, VA_B, VA_C)) / Avg(VA_A, VA_B, VA_C) * 100
    """
    phase_a = 0.0
    phase_b = 0.0
    phase_c = 0.0

    for ckt in circuits:
        p = ckt.get("phasePole", "A").upper()
        va = float(ckt.get("connectedVA", 0.0) or 0.0)
        if p == "A":
            phase_a += va
        elif p == "B":
            phase_b += va
        elif p == "C":
            phase_c += va

    total_connected = phase_a + phase_b + phase_c
    
    if phase_count == 3:
        phases = [phase_a, phase_b, phase_c]
        avg_phase = (phase_a + phase_b + phase_c) / 3.0 if total_connected > 0 else 0.0
        if avg_phase > 0:
            unbalance = ((max(phases) - min(phases)) / avg_phase) * 100.0
        else:
            unbalance = 0.0
    else:
        phases = [phase_a, phase_b]
        avg_phase = (phase_a + phase_b) / 2.0 if total_connected > 0 else 0.0
        if avg_phase > 0:
            unbalance = ((max(phases) - min(phases)) / avg_phase) * 100.0
        else:
            unbalance = 0.0

    return {
        "phaseAConnectedVA": round(phase_a, 2),
        "phaseBConnectedVA": round(phase_b, 2),
        "phaseCConnectedVA": round(phase_c, 2) if phase_count == 3 else 0.0,
        "totalConnectedVA": round(total_connected, 2),
        "totalConnectedAmps": 0.0,
        "phaseADemandVA": round(phase_a, 2),
        "phaseBDemandVA": round(phase_b, 2),
        "phaseCDemandVA": round(phase_c, 2) if phase_count == 3 else 0.0,
        "totalDemandVA": round(total_connected, 2),
        "totalDemandAmps": 0.0,
        "unbalancePercentage": round(unbalance, 2)
    }


def parse_panel_sheet(ws: openpyxl.worksheet.worksheet.Worksheet, workbook_path: str = "") -> Dict[str, Any]:
    """
    Parses a single Excel worksheet into a canonical panel schedule dictionary conforming
    to shared/schemas/panel_schedule.schema.json.
    """
    # 1. Extract Header Information
    title_cell = clean_cell_str(ws["A3"].value)
    # Parse panel name from title e.g. "(E) PANEL 'LP-1'" or "PANEL LP-1" or sheet title
    panel_name_match = re.search(r"PANEL\s*['\"]?([^'\"]+)['\"]?", title_cell, re.IGNORECASE)
    if panel_name_match:
        panel_name = panel_name_match.group(1).strip()
    else:
        panel_name = ws.title.strip()

    raw_voltage = clean_cell_str(ws[f"{COL_VOLTAGE}2"].value)
    voltage = "120/208V"
    if "480" in raw_voltage or "277" in raw_voltage:
        voltage = "277/480V"
    elif "240" in raw_voltage:
        voltage = "120/240V"

    bus_amps_raw = ws["G3"].value
    bus_amps = parse_integer(bus_amps_raw, default=225)
    
    wire_raw = ws["K2"].value
    wire_count = parse_integer(wire_raw, default=4)
    if wire_count not in (3, 4):
        wire_count = 4

    phase_raw = ws["K3"].value
    phase_count = parse_integer(phase_raw, default=3)
    if phase_count not in (1, 3):
        phase_count = 3

    enclosure = clean_cell_str(ws["K4"].value) or "NEMA 1"
    mounting = clean_cell_str(ws["N2"].value) or "SURFACE"
    
    aic_raw = ws["N3"].value
    aic_rating = parse_integer(aic_raw, default=10000)
    if aic_rating < 1000:
        aic_rating = aic_rating * 1000  # handle e.g. 10 -> 10000

    diagnostics = []

    # 2. Extract Circuits
    circuits = []
    
    # Phase pole sequence helper
    # For 3-phase: Row 0 -> A, Row 1 -> B, Row 2 -> C, Row 3 -> A...
    # For 1-phase: Row 0 -> A, Row 1 -> B, Row 2 -> A...
    phases_seq = ["A", "B", "C"] if phase_count == 3 else ["A", "B"]

    for row_idx in range(START_ROW, MAX_ROW + 1):
        rel_idx = row_idx - START_ROW
        phase_pole = phases_seq[rel_idx % len(phases_seq)]

        # --- Left Column (Odd Circuit) ---
        odd_ckt_num = 2 * rel_idx + 1
        l_desc = clean_cell_str(ws[f"{COL_L_DESC}{row_idx}"].value)
        l_trip = parse_integer(ws[f"{COL_L_TRIP}{row_idx}"].value, default=20)
        l_pole = parse_integer(ws[f"{COL_L_POLE}{row_idx}"].value, default=1)
        l_type = clean_cell_str(ws[f"{COL_L_TYPE}{row_idx}"].value)
        l_kva = parse_numeric(ws[f"{COL_L_KVA}{row_idx}"].value, default=0.0)
        
        # Convert kVA to VA (multiply by 1000 if in kVA format e.g. 1.6 -> 1600)
        l_va = l_kva * 1000.0 if 0 < l_kva < 50 else l_kva

        if l_desc and l_desc != "---":
            load_type = map_load_type(l_type, l_desc)
            circuits.append({
                "circuitNumber": odd_ckt_num,
                "phasePole": phase_pole,
                "loadDescription": l_desc,
                "loadType": load_type,
                "breakerAmps": l_trip if l_trip > 0 else 20,
                "poles": l_pole if l_pole in (1, 2, 3) else 1,
                "wireGauge": "#12 AWG",
                "conduitSize": '3/4" C',
                "connectedVA": round(l_va, 2),
                "demandFactorPercent": 100.0,
                "demandVA": round(l_va, 2),
                "roomOrZone": ""
            })
        elif not l_desc or "SPARE" in l_desc or "SPACE" in l_desc:
            circuits.append({
                "circuitNumber": odd_ckt_num,
                "phasePole": phase_pole,
                "loadDescription": l_desc or "SPARE",
                "loadType": "SPARE",
                "breakerAmps": 20,
                "poles": 1,
                "wireGauge": "#12 AWG",
                "conduitSize": '3/4" C',
                "connectedVA": 0.0,
                "demandFactorPercent": 100.0,
                "demandVA": 0.0,
                "roomOrZone": ""
            })

        # --- Right Column (Even Circuit) ---
        even_ckt_num = 2 * rel_idx + 2
        r_desc = clean_cell_str(ws[f"{COL_R_DESC}{row_idx}"].value)
        r_trip = parse_integer(ws[f"{COL_R_TRIP}{row_idx}"].value, default=20)
        r_pole = parse_integer(ws[f"{COL_R_POLE}{row_idx}"].value, default=1)
        r_type = clean_cell_str(ws[f"{COL_R_TYPE}{row_idx}"].value)
        r_kva = parse_numeric(ws[f"{COL_R_KVA}{row_idx}"].value, default=0.0)
        r_va = r_kva * 1000.0 if 0 < r_kva < 50 else r_kva

        if r_desc and r_desc != "---":
            load_type = map_load_type(r_type, r_desc)
            circuits.append({
                "circuitNumber": even_ckt_num,
                "phasePole": phase_pole,
                "loadDescription": r_desc,
                "loadType": load_type,
                "breakerAmps": r_trip if r_trip > 0 else 20,
                "poles": r_pole if r_pole in (1, 2, 3) else 1,
                "wireGauge": "#12 AWG",
                "conduitSize": '3/4" C',
                "connectedVA": round(r_va, 2),
                "demandFactorPercent": 100.0,
                "demandVA": round(r_va, 2),
                "roomOrZone": ""
            })
        elif not r_desc or "SPARE" in r_desc or "SPACE" in r_desc:
            circuits.append({
                "circuitNumber": even_ckt_num,
                "phasePole": phase_pole,
                "loadDescription": r_desc or "SPARE",
                "loadType": "SPARE",
                "breakerAmps": 20,
                "poles": 1,
                "wireGauge": "#12 AWG",
                "conduitSize": '3/4" C',
                "connectedVA": 0.0,
                "demandFactorPercent": 100.0,
                "demandVA": 0.0,
                "roomOrZone": ""
            })

    # Sort circuits by circuit number
    circuits.sort(key=lambda x: x["circuitNumber"])

    # 3. Calculate Phase Balance & Diagnostics
    load_summary = calculate_phase_balance(circuits, phase_count)
    
    validation_status = "VALID"
    if load_summary["unbalancePercentage"] > 10.0:
        validation_status = "WARNINGS"
        diagnostics.append(f"High phase unbalance detected: {load_summary['unbalancePercentage']}% (Target <= 5.0%)")
    elif load_summary["unbalancePercentage"] > 5.0:
        diagnostics.append(f"Phase unbalance exceeds 5.0%: {load_summary['unbalancePercentage']}%")

    if not panel_name or panel_name.upper() == "TEMPLATE":
        validation_status = "WARNINGS"
        diagnostics.append("Panel name is undefined or default template name.")

    return {
        "panelName": panel_name,
        "voltage": voltage,
        "phase": phase_count,
        "wire": wire_count,
        "mainBusRatingAmps": bus_amps,
        "mainType": "MCB",
        "mainBreakerAmps": bus_amps,
        "enclosureNema": enclosure,
        "shortCircuitCurrentRatingAIC": aic_rating,
        "location": "ELECTRICAL ROOM",
        "excelWorkbookPath": workbook_path,
        "validationStatus": validation_status,
        "diagnostics": diagnostics,
        "loadSummary": load_summary,
        "circuits": circuits
    }


def parse_panel_workbook(workbook_path: str) -> List[Dict[str, Any]]:
    """
    Opens an Excel panel schedule workbook (.xlsx) and parses all panel sheets (ignoring TEMPLATE).
    """
    if not os.path.exists(workbook_path):
        raise FileNotFoundError(f"Panel workbook not found at: {workbook_path}")

    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    panels = []

    for sheet_name in wb.sheetnames:
        if sheet_name.upper() in ("TEMPLATE", "SUMMARY", "CONFIG"):
            continue
        ws = wb[sheet_name]
        try:
            panel_data = parse_panel_sheet(ws, workbook_path=workbook_path)
            panels.append(panel_data)
        except Exception as ex:
            panels.append({
                "panelName": sheet_name,
                "validationStatus": "ERRORS",
                "diagnostics": [f"Failed to parse sheet '{sheet_name}': {str(ex)}"],
                "voltage": "120/208V",
                "phase": 3,
                "wire": 4,
                "mainBusRatingAmps": 225,
                "mainType": "MCB",
                "circuits": []
            })
    return panels


def sync_panel_workbook_to_db(project_id: str, workbook_path: str, db_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Parses a panel workbook and upserts all discovered panels into the SQLite database.
    """
    panels = parse_panel_workbook(workbook_path)
    synced_records = []
    for p in panels:
        if p.get("panelName") and p.get("circuits"):
            record = upsert_panel_schedule(project_id, p, db_path=db_path)
            synced_records.append(record)
    return synced_records
