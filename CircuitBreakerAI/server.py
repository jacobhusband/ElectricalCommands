import os
import shutil
import uuid
import base64
import time
import math
import re
import random
import json
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse, FileResponse, StreamingResponse
from pydantic import BaseModel, Field
import openpyxl
from google.genai import types
from google import genai
from PIL import Image

app = FastAPI()

# ==========================================
# 1. CONFIGURATION & STATE
# ==========================================

_SERVER_DIR = Path(__file__).resolve().parent

UPLOAD_DIR = _SERVER_DIR / "temp_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# ElectricalPanels Config
API_KEY = os.environ.get("GEMINI_API_KEY", os.environ.get("GOOGLE_API_KEY", ""))
TEMPLATE_FILE = str(_SERVER_DIR / "ElectricalPanels" / "Template.xlsx")
OUTPUT_FILE = str(_SERVER_DIR / "ElectricalPanels" / "Filled_Panel_Schedules.xlsx")

# Rate Limiting
MAX_REQUESTS = 2
TIME_WINDOW = 61
api_call_timestamps = []

# Session tracking - reset output file on first panel of each session
session_initialized = False

# Excel Columns
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


# Initialize Gemini client
client = genai.Client(
    api_key=API_KEY,
    http_options=types.HttpOptions(timeout=600000),
)

# Serve static files
app.mount("/static", StaticFiles(directory=str(_SERVER_DIR / "static")), name="static")


# ==========================================
# 2. DATA MODELS
# ==========================================

class CircuitItem(BaseModel):
    circuit_number: int = Field(..., description="The first circuit number of the breaker")
    description: str = Field(..., description="Load description")
    breaker_amps: str = Field(..., description="Amperage (e.g., '20')")
    poles: int = Field(1, description="Number of poles (1, 2, or 3)")
    load_type: str = Field(..., description="Code: 'C', 'D', 'G', 'K', 'M'")
    kva: str = Field("", description="Visible load in kVA from an existing directory schedule, or blank if not shown")
    phase_kva: List[str] = Field([], description="Visible kVA values for each occupied circuit row, ordered top-to-bottom")

class PanelData(BaseModel):
    panel_name: str = Field(..., description="Panel Name")
    voltage: str = Field(..., description="Voltage")
    bus_rating: str = Field(..., description="Bus Rating")
    aic_rating: str = Field("", description="Panel AIC rating when visibly shown, or blank if not shown")
    phase: str = Field(..., description="Phase")
    wire: str = Field(..., description="Wire")
    mounting: str = Field(..., description="Mounting")
    enclosure: str = Field(..., description="Enclosure")
    circuits: List[CircuitItem] = Field(..., description="List of detected breakers")


# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================

def enforce_rate_limit():
    global api_call_timestamps
    now = time.time()
    api_call_timestamps = [t for t in api_call_timestamps if now - t < TIME_WINDOW]
    if len(api_call_timestamps) >= MAX_REQUESTS:
        earliest_call = api_call_timestamps[0]
        wait_time = (earliest_call + TIME_WINDOW) - now
        if wait_time > 0:
            print(f"   [Rate Limit] Waiting {wait_time:.1f} seconds...")
            time.sleep(wait_time)
    api_call_timestamps.append(time.time())

def clean_text(text):
    if text is None:
        return ""
    return str(text).upper().strip()

def calculate_estimated_load(amps_str, description, load_type):
    """
    Calculate KVA load based on circuit description type.
    Uses rule-based assignment for different equipment types.
    """
    desc = clean_text(description)

    # Skip spare/unused circuits
    if any(x in desc for x in ["SPARE", "SPACE", "UNUSED"]):
        return 0.00

    # Get breaker capacity in KVA (amps * 120V / 1000)
    try:
        numeric_amps = float("".join(filter(str.isdigit, str(amps_str))))
        breaker_kva = (numeric_amps * 120) / 1000
    except:
        breaker_kva = 0

    # A/C units - 70% of breaker capacity
    if any(x in desc for x in ["A/C", "AC", "AIR COND", "CONDENSER", "HVAC", "COOLING", "COMPRESSOR"]):
        return round(breaker_kva * 0.70, 2)

    # Water heaters - 60% of breaker capacity
    if any(x in desc for x in ["WATER HEATER", "WH", "HWH", "HOT WATER"]):
        return round(breaker_kva * 0.60, 2)

    # Fans - 50% of breaker capacity
    if any(x in desc for x in ["FAN", "EXHAUST", "VENTILAT", "VENT FAN"]):
        return round(breaker_kva * 0.50, 2)

    # ATMs - fixed 0.6 KVA
    if "ATM" in desc:
        return 0.60

    # Exterior pole lights - random 1.0, 1.1, or 1.2
    if any(x in desc for x in ["POLE LIGHT", "POLE LT", "PARKING LOT", "LOT LIGHT"]):
        return random.choice([1.0, 1.1, 1.2])

    # Exterior lights (not pole) - random 0.6 to 1.0
    if any(x in desc for x in ["EXT LIGHT", "EXTERIOR LIGHT", "OUTSIDE LIGHT", "OUTDOOR LIGHT",
                                "EXT LT", "EXTERIOR LT", "SIGN", "FACADE", "BUILDING LIGHT"]):
        return random.choice([0.6, 0.7, 0.8, 0.9, 1.0])

    # Interior lights - weighted by room size indicators
    if any(x in desc for x in ["LIGHT", "LT", "LTG", "LIGHTING", "LAMP"]):
        # Higher values for larger-sounding rooms
        large_room_keywords = ["LOBBY", "HALL", "CONFERENCE", "MEETING", "AUDITORIUM", "GYM",
                               "WAREHOUSE", "SHOWROOM", "DINING", "MAIN", "OPEN", "COMMON"]
        medium_room_keywords = ["OFFICE", "BREAK", "STORAGE", "KITCHEN", "LUNCH"]

        if any(kw in desc for kw in large_room_keywords):
            return random.choice([0.7, 0.8, 0.9])
        elif any(kw in desc for kw in medium_room_keywords):
            return random.choice([0.5, 0.6, 0.7])
        else:
            return random.choice([0.3, 0.4, 0.5])

    # Plugs/Receptacles - weighted by room type
    if any(x in desc for x in ["PLUG", "RECEP", "DUPLEX", "OUTLET", "REC", "RCPT"]):
        # Higher values for rooms with more receptacles
        high_recep_keywords = ["KITCHEN", "BREAK", "LAB", "WORKSTATION", "COMPUTER", "SERVER",
                               "OFFICE", "CONF", "MEETING", "NURSE", "MEDICAL"]
        medium_recep_keywords = ["STORAGE", "UTILITY", "MECH", "CLOSET", "REST", "BATH"]

        if any(kw in desc for kw in high_recep_keywords):
            return random.choice([0.72, 0.9])
        elif any(kw in desc for kw in medium_recep_keywords):
            return random.choice([0.36, 0.54])
        else:
            return random.choice([0.36, 0.54, 0.72, 0.9])

    # Default fallback - use old formula
    l_type = clean_text(load_type)
    factor = 0.50 if l_type in ['C', 'G'] else 0.80
    return round((numeric_amps * factor * 120) / 1000, 2)

def normalize_extracted_kva(raw_kva):
    if raw_kva is None or isinstance(raw_kva, bool):
        return ""
    if isinstance(raw_kva, (int, float)):
        value = float(raw_kva)
    else:
        text = str(raw_kva).strip()
        if not text:
            return ""
        upper = text.upper()
        if upper in {"N/A", "NA", "NONE", "UNKNOWN", "UNREADABLE", "BLANK", "-", "--", "---"}:
            return ""
        match = re.search(r"[-+]?\d+(?:\.\d+)?", upper.replace(",", ""))
        if not match:
            return ""
        try:
            value = float(match.group(0))
        except Exception:
            return ""
        if "VA" in upper and "KVA" not in upper and value >= 10:
            value = value / 1000
    if not math.isfinite(value) or value < 0:
        return ""
    return round(value, 3)

def normalize_extracted_phase_kva(circuit):
    raw_values = getattr(circuit, "phase_kva", None)
    if isinstance(raw_values, (list, tuple)):
        values = [normalize_extracted_kva(value) for value in raw_values]
        if any(value != "" for value in values):
            return values
    return [normalize_extracted_kva(getattr(circuit, "kva", ""))]

def fix_nema_type(raw_type):
    t = str(raw_type).upper()
    return "NEMA 3R" if any(x in t for x in ["3R", "OUT", "EXT", "WEATHER"]) else "NEMA 1"

def resolve_ditto_marks(circuits: List[CircuitItem]) -> List[CircuitItem]:
    ditto_pattern = re.compile(r'^[\""\u2018\u2019\u201c\u201d\.]+$|^(SAME|DO)$', re.IGNORECASE)
    odds = sorted([c for c in circuits if c.circuit_number % 2 != 0], key=lambda x: x.circuit_number)
    evens = sorted([c for c in circuits if c.circuit_number % 2 == 0], key=lambda x: x.circuit_number)

    def process_column(col_list):
        last_desc = ""
        for ckt in col_list:
            if ditto_pattern.match(ckt.description.strip()):
                if last_desc:
                    ckt.description = last_desc
            elif ckt.description.strip() and ckt.description.strip() != "---":
                last_desc = ckt.description
        return col_list
    return process_column(odds) + process_column(evens)

def update_excel_workbook(panel_data: PanelData, folder_name: str, use_extracted_kva=False):
    global session_initialized

    # On first panel of the session, start fresh from template
    # Also reset if the output file was cleaned up externally.
    if (not session_initialized) or (not os.path.exists(OUTPUT_FILE)):
        # Delete existing output file if it exists
        if os.path.exists(OUTPUT_FILE):
            os.remove(OUTPUT_FILE)
            print(f"   [Session Reset] Deleted existing output file: {OUTPUT_FILE}")

        # Copy template to output location
        shutil.copy(TEMPLATE_FILE, OUTPUT_FILE)
        print(f"   [Session Reset] Created fresh output from template")
        session_initialized = True

    # Load the output workbook
    wb = openpyxl.load_workbook(OUTPUT_FILE)

    # Ensure template sheet exists to copy from
    if "TEMPLATE" not in wb.sheetnames:
        raise ValueError("TEMPLATE sheet not found in workbook")

    source = wb["TEMPLATE"]
    target = wb.copy_worksheet(source)
    target.sheet_view.showGridLines = False

    # 1. Update Sheet Name and Header
    safe_name = folder_name.upper().replace("/", "-").replace("\\", "-")[:31]
    target.title = safe_name
    target["A3"] = f"(E) PANEL '{safe_name}'"

    # Header Info
    target[f"{COL_VOLTAGE}2"] = clean_text(panel_data.voltage)
    target["G3"] = clean_text(panel_data.bus_rating)
    target["K2"] = clean_text(panel_data.wire)
    target["K3"] = clean_text(panel_data.phase)
    target["K4"] = fix_nema_type(panel_data.enclosure)
    target["N2"] = clean_text(panel_data.mounting)
    aic_rating = clean_text(getattr(panel_data, "aic_rating", ""))
    if aic_rating:
        target["N3"] = aic_rating

    # Circuits
    START_ROW = 8
    MAX_ROW = 28
    occupied_slots = set()

    cleaned_circuits = resolve_ditto_marks(panel_data.circuits)
    sorted_ckts = sorted(cleaned_circuits, key=lambda x: x.circuit_number)

    for ckt in sorted_ckts:
        c_num = ckt.circuit_number
        if c_num > 42: continue

        row_idx = START_ROW + math.ceil(c_num / 2) - 1
        if row_idx > MAX_ROW: continue

        is_odd = (c_num % 2 != 0)
        side = "L" if is_odd else "R"

        if (side, row_idx) in occupied_slots: continue

        if is_odd:
            c_note, c_type, c_pole, c_trip, c_desc, c_kva = COL_L_NOTE, COL_L_TYPE, COL_L_POLE, COL_L_TRIP, COL_L_DESC, COL_L_KVA
        else:
            c_note, c_type, c_pole, c_trip, c_desc, c_kva = COL_R_NOTE, COL_R_TYPE, COL_R_POLE, COL_R_TRIP, COL_R_DESC, COL_R_KVA

        desc_text = clean_text(ckt.description)
        is_spare = any(x in desc_text for x in ["SPARE", "SPACE", "UNUSED"])

        if is_spare:
            kva_val = ""
            phase_kva_values = []
        elif use_extracted_kva:
            phase_kva_values = normalize_extracted_phase_kva(ckt)
            kva_val = phase_kva_values[0] if phase_kva_values else ""
        else:
            kva_val = calculate_estimated_load(ckt.breaker_amps, ckt.description, ckt.load_type)
            phase_kva_values = []
        note_val = "" if is_spare else "1"
        type_val = "" if is_spare else clean_text(ckt.load_type)

        target[f"{c_desc}{row_idx}"] = desc_text
        target[f"{c_pole}{row_idx}"] = ckt.poles
        target[f"{c_trip}{row_idx}"] = clean_text(ckt.breaker_amps)
        target[f"{c_kva}{row_idx}"] = kva_val
        target[f"{c_note}{row_idx}"] = note_val
        target[f"{c_type}{row_idx}"] = type_val

        occupied_slots.add((side, row_idx))

        if ckt.poles > 1:
            for i in range(1, ckt.poles):
                ext_row = row_idx + i
                if ext_row <= MAX_ROW:
                    target[f"{c_desc}{ext_row}"] = "---"
                    target[f"{c_pole}{ext_row}"] = "-"
                    target[f"{c_trip}{ext_row}"] = "-"
                    if use_extracted_kva:
                        ext_kva_val = phase_kva_values[i] if i < len(phase_kva_values) else ""
                    else:
                        ext_kva_val = kva_val
                    target[f"{c_kva}{ext_row}"] = ext_kva_val
                    target[f"{c_type}{ext_row}"] = type_val
                    target[f"{c_note}{ext_row}"] = note_val
                    occupied_slots.add((side, ext_row))

    wb.save(OUTPUT_FILE)
    return True


# ==========================================
# 4. API ENDPOINTS
# ==========================================

@app.get("/")
async def read_index():
    return FileResponse(str(_SERVER_DIR / "static" / "index.html"))

@app.post("/api/analyze-panel")
async def analyze_panel_endpoint(
    panel_name: str = Form(...),
    input_mode: str = Form("field_photos"),
    breaker_images: List[UploadFile] = File(default=[]),
    directory_images: List[UploadFile] = File(default=[])
):
    try:
        # Save uploaded files
        saved_breaker_paths = []
        saved_directory_paths = []

        # Helper to save files
        def save_uploads(files):
            paths = []
            for file in files:
                path = UPLOAD_DIR / f"{uuid.uuid4()}_{file.filename}"
                with path.open("wb") as buffer:
                    shutil.copyfileobj(file.file, buffer)
                paths.append(str(path))
            return paths

        saved_breaker_paths = save_uploads(breaker_images)
        saved_directory_paths = save_uploads(directory_images)

        # Validation based on input mode
        if input_mode == "existing_directory":
            if not saved_directory_paths:
                raise HTTPException(status_code=400, detail="At least one directory photo is required for existing directory mode.")
        else:
            if not saved_breaker_paths or not saved_directory_paths:
                raise HTTPException(status_code=400, detail="At least one breaker photo and at least one directory photo are required.")

        # Prepare images for Gemini
        gemini_images = []
        for p in saved_breaker_paths:
            gemini_images.append(Image.open(p))
        for p in saved_directory_paths:
            gemini_images.append(Image.open(p))

        num_breaker_imgs = len(saved_breaker_paths)
        num_dir_imgs = len(saved_directory_paths)

        # Gemini Prompt
        if input_mode == "existing_directory":
            prompt = f"""
            Analyze these electrical panel schedule directory document photos for Panel: {panel_name}.

            You are provided with {num_dir_imgs} images of the existing panel directory document. This document contains the circuit information (circuit numbers, load descriptions, breaker ratings/amperages, poles, and any visible kVA/load values) needed to recreate the panel schedule.

            IMAGE INTERPRETATION RULES
            - All directory images belong to the SAME panel directory document.
            - Directory coverage may be split across multiple photos, such as circuits 1-42 on one image and 43-84 on another.
            - Merge partial and overlapping views into one complete understanding of the panel.
            - Reconcile overlapping information by using visible circuit numbers, breaker positions, and the clearest label text.

            TASK 1: HEADER
            - Extract Voltage, Bus Rating, Wire, Phase, Mounting, Enclosure, and panel AIC rating from any headers, tables, or title block details.
            - Set the JSON field "aic_rating" to the visible AIC value (examples: "10 KAIC", "22,000 AIC").
            - If AIC is blank, hidden, unreadable, or not present, set "aic_rating" to "".
            - Do not confuse AIC rating with Bus Rating, breaker amperage, voltage, or main requirement.

            TASK 2: CIRCUITS & POLES
            - Identify every circuit entry in the directory document.
            - CRITICAL: Determine the 'poles' (1, 2, or 3) for each active circuit breaker.
                - Look at how descriptions or circuit lines are grouped or span multiple spaces to represent multi-pole breakers.
                - A 2-pole breaker has tied circuit spaces (e.g., descriptions spanning circuits 1 and 3, or a single breaker amperage serving a 2-circuit load).
                - Provide the circuit_number as the TOP-most circuit number the breaker occupies.
            - Extract Amperage (e.g. 20A, 30A, 50A), Load Description, and the visible kVA/load values from the directory table.
            - For each circuit, set the JSON field "phase_kva" to an ordered array of visible kVA decimals for the occupied circuit rows, top-to-bottom (examples: ["0.72"] for 1-pole, ["1.20", "1.10"] for 2-pole, ["2.00", "2.10", "2.05"] for 3-pole).
            - If a value is shown in VA, convert it to kVA. If an occupied row's kVA/load value is blank, hidden, unreadable, or not present, use "" for that row in "phase_kva".
            - Do not add, total, or sum phase loads together. Do not repeat a 2-pole or 3-pole total on every occupied row.
            - Keep the legacy JSON field "kva" blank unless only one single visible total is available and there are no row-by-row phase values.
            - Do not estimate kVA and do not infer it from breaker amperage, voltage, load type, or description.
            - Resolve ditto marks (") if seen in the directory.
            - If the same circuit appears in overlapping photos, combine the evidence and keep one final record.

            TASK 3: LOAD TYPES
            - LIGHTING -> 'C', RECEPTACLES -> 'G', MOTORS/HVAC -> 'M', KITCHEN -> 'K', DEDICATED -> 'D'
            """.strip()
        else:
            prompt = f"""
            Analyze these electrical panel photos for Panel: {panel_name}.

            You are provided with {num_breaker_imgs} images of the CIRCUIT BREAKERS (first {num_breaker_imgs} images)
            and {num_dir_imgs} images of the CIRCUIT DIRECTORY (last {num_dir_imgs} images).

            IMAGE INTERPRETATION RULES
            - All breaker images belong to the SAME panel.
            - All directory images belong to the SAME panel.
            - Breaker coverage may be split across multiple photos of the same panel, such as upper half, middle, and bottom half images.
            - Directory coverage may be split across multiple photos, such as circuits 1-42 on one image and 43-84 on another.
            - Merge partial and overlapping views into one complete understanding of the panel.
            - Do not treat split photos as separate panels.
            - Reconcile overlapping information by using visible circuit numbers, breaker positions, and the clearest label text.
            - Breaker images come first. Directory images come last.
            - Do not assume each image shows the entire panel or a complete directory by itself.

            TASK 1: HEADER
            - Extract Voltage, Bus Rating, Wire, Phase, Mounting, Enclosure.
            - Look at the directory images or labels on the panel across all images.
            - Extract the panel AIC rating into JSON field "aic_rating" when visibly shown (examples: "10 KAIC", "22,000 AIC").
            - If AIC is blank, hidden, unreadable, or not present, set "aic_rating" to "".
            - Do not confuse AIC rating with Bus Rating, breaker amperage, voltage, or main requirement.

            TASK 2: CIRCUITS & POLES
            - Identify every breaker visible across the Breaker Images.
            - Use visible breaker positions and circuit numbering to merge split sections from multiple photos.
            - CRITICAL: Determine the 'poles' (1, 2, or 3).
                - A 1-pole breaker takes up 1 circuit space.
                - A 2-pole breaker has a tied handle and takes up 2 vertical circuit spaces (e.g. 1 & 3).
                - A 3-pole breaker has a tied handle and takes up 3 vertical circuit spaces (e.g. 1, 3, & 5).
            - Provide the circuit_number as the TOP-most circuit number the breaker occupies.
            - Extract Amperage and Load Description from the labels (Breaker Images) or circuit directory (Directory Images).
            - Resolve ditto marks (") if seen in the directory.
            - If the same circuit appears in overlapping photos, combine the evidence and keep one final record.

            TASK 3: LOAD TYPES
            - LIGHTING -> 'C', RECEPTACLES -> 'G', MOTORS/HVAC -> 'M', KITCHEN -> 'K', DEDICATED -> 'D'
            """.strip()

        enforce_rate_limit()

        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=[prompt, *gemini_images],
            config=types.GenerateContentConfig(
                response_mime_type="application/json", response_schema=PanelData),
        )

        panel_data = response.parsed
        panel_data.panel_name = panel_name # Ensure name matches

        # Update Excel
        update_excel_workbook(
            panel_data,
            panel_name,
            use_extracted_kva=input_mode == "existing_directory",
        )

        return {
            "status": "success",
            "message": f"Panel '{panel_name}' added to schedule.",
            "data": panel_data.model_dump()
        }

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/analyze-panel-stream")
async def analyze_panel_stream_endpoint(
    panel_name: str = Form(...),
    input_mode: str = Form("field_photos"),
    breaker_images: List[UploadFile] = File(default=[]),
    directory_images: List[UploadFile] = File(default=[])
):
    """SSE endpoint that streams status updates, then final Gemini results."""

    async def generate_events():
        try:
            # Save uploaded files
            saved_breaker_paths = []
            saved_directory_paths = []

            def save_uploads(files):
                paths = []
                for file in files:
                    path = UPLOAD_DIR / f"{uuid.uuid4()}_{file.filename}"
                    with path.open("wb") as buffer:
                        shutil.copyfileobj(file.file, buffer)
                    paths.append(str(path))
                return paths

            saved_breaker_paths = save_uploads(breaker_images)
            saved_directory_paths = save_uploads(directory_images)

            # Validation based on input mode
            if input_mode == "existing_directory":
                if not saved_directory_paths:
                    yield f"data: {json.dumps({'event': 'error', 'detail': 'At least one directory photo is required for existing directory mode.'})}\n\n"
                    return
            else:
                if not saved_breaker_paths or not saved_directory_paths:
                    yield f"data: {json.dumps({'event': 'error', 'detail': 'At least one breaker photo and at least one directory photo are required.'})}\n\n"
                    return

            # Prepare images for Gemini
            gemini_images = []
            for p in saved_breaker_paths:
                gemini_images.append(Image.open(p))
            for p in saved_directory_paths:
                gemini_images.append(Image.open(p))

            num_breaker_imgs = len(saved_breaker_paths)
            num_dir_imgs = len(saved_directory_paths)

            # Signal Gemini analysis starting
            yield f"data: {json.dumps({'event': 'gemini_started', 'total_images': num_breaker_imgs + num_dir_imgs})}\n\n"

            # Gemini Prompt
            if input_mode == "existing_directory":
                prompt = f"""
                Analyze these electrical panel schedule directory document photos for Panel: {panel_name}.

                You are provided with {num_dir_imgs} images of the existing panel directory document. This document contains the circuit information (circuit numbers, load descriptions, breaker ratings/amperages, poles, and any visible kVA/load values) needed to recreate the panel schedule.

                IMAGE INTERPRETATION RULES
                - All directory images belong to the SAME panel directory document.
                - Directory coverage may be split across multiple photos, such as circuits 1-42 on one image and 43-84 on another.
                - Merge partial and overlapping views into one complete understanding of the panel.
                - Reconcile overlapping information by using visible circuit numbers, breaker positions, and the clearest label text.

                TASK 1: HEADER
                - Extract Voltage, Bus Rating, Wire, Phase, Mounting, Enclosure, and panel AIC rating from any headers, tables, or title block details.
                - Set the JSON field "aic_rating" to the visible AIC value (examples: "10 KAIC", "22,000 AIC").
                - If AIC is blank, hidden, unreadable, or not present, set "aic_rating" to "".
                - Do not confuse AIC rating with Bus Rating, breaker amperage, voltage, or main requirement.

                TASK 2: CIRCUITS & POLES
                - Identify every circuit entry in the directory document.
                - CRITICAL: Determine the 'poles' (1, 2, or 3) for each active circuit breaker.
                    - Look at how descriptions or circuit lines are grouped or span multiple spaces to represent multi-pole breakers.
                    - A 2-pole breaker has tied circuit spaces (e.g., descriptions spanning circuits 1 and 3, or a single breaker amperage serving a 2-circuit load).
                    - Provide the circuit_number as the TOP-most circuit number the breaker occupies.
                - Extract Amperage (e.g. 20A, 30A, 50A), Load Description, and the visible kVA/load values from the directory table.
                - For each circuit, set the JSON field "phase_kva" to an ordered array of visible kVA decimals for the occupied circuit rows, top-to-bottom (examples: ["0.72"] for 1-pole, ["1.20", "1.10"] for 2-pole, ["2.00", "2.10", "2.05"] for 3-pole).
                - If a value is shown in VA, convert it to kVA. If an occupied row's kVA/load value is blank, hidden, unreadable, or not present, use "" for that row in "phase_kva".
                - Do not add, total, or sum phase loads together. Do not repeat a 2-pole or 3-pole total on every occupied row.
                - Keep the legacy JSON field "kva" blank unless only one single visible total is available and there are no row-by-row phase values.
                - Do not estimate kVA and do not infer it from breaker amperage, voltage, load type, or description.
                - Resolve ditto marks (") if seen in the directory.
                - If the same circuit appears in overlapping photos, combine the evidence and keep one final record.

                TASK 3: LOAD TYPES
                - LIGHTING -> 'C', RECEPTACLES -> 'G', MOTORS/HVAC -> 'M', KITCHEN -> 'K', DEDICATED -> 'D'
                """.strip()
            else:
                prompt = f"""
                Analyze these electrical panel photos for Panel: {panel_name}.

                You are provided with {num_breaker_imgs} images of the CIRCUIT BREAKERS (first {num_breaker_imgs} images)
                and {num_dir_imgs} images of the CIRCUIT DIRECTORY (last {num_dir_imgs} images).

                IMAGE INTERPRETATION RULES
                - All breaker images belong to the SAME panel.
                - All directory images belong to the SAME panel.
                - Breaker coverage may be split across multiple photos of the same panel, such as upper half, middle, and bottom half images.
                - Directory coverage may be split across multiple photos, such as circuits 1-42 on one image and 43-84 on another.
                - Merge partial and overlapping views into one complete understanding of the panel.
                - Do not treat split photos as separate panels.
                - Reconcile overlapping information by using visible circuit numbers, breaker positions, and the clearest label text.
                - Breaker images come first. Directory images come last.
                - Do not assume each image shows the entire panel or a complete directory by itself.

                TASK 1: HEADER
                - Extract Voltage, Bus Rating, Wire, Phase, Mounting, Enclosure.
                - Look at the directory images or labels on the panel across all images.
                - Extract the panel AIC rating into JSON field "aic_rating" when visibly shown (examples: "10 KAIC", "22,000 AIC").
                - If AIC is blank, hidden, unreadable, or not present, set "aic_rating" to "".
                - Do not confuse AIC rating with Bus Rating, breaker amperage, voltage, or main requirement.

                TASK 2: CIRCUITS & POLES
                - Identify every breaker visible across the Breaker Images.
                - Use visible breaker positions and circuit numbering to merge split sections from multiple photos.
                - CRITICAL: Determine the 'poles' (1, 2, or 3).
                    - A 1-pole breaker takes up 1 circuit space.
                    - A 2-pole breaker has a tied handle and takes up 2 vertical circuit spaces (e.g. 1 & 3).
                    - A 3-pole breaker has a tied handle and takes up 3 vertical circuit spaces (e.g. 1, 3, & 5).
                - Provide the circuit_number as the TOP-most circuit number the breaker occupies.
                - Extract Amperage and Load Description from the labels (Breaker Images) or circuit directory (Directory Images).
                - Resolve ditto marks (") if seen in the directory.
                - If the same circuit appears in overlapping photos, combine the evidence and keep one final record.

                TASK 3: LOAD TYPES
                - LIGHTING -> 'C', RECEPTACLES -> 'G', MOTORS/HVAC -> 'M', KITCHEN -> 'K', DEDICATED -> 'D'
                """.strip()

            enforce_rate_limit()

            response = client.models.generate_content(
                model="gemini-3.6-flash",
                contents=[prompt, *gemini_images],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json", response_schema=PanelData),
            )

            panel_data = response.parsed
            panel_data.panel_name = panel_name

            # Update Excel
            update_excel_workbook(
                panel_data,
                panel_name,
                use_extracted_kva=input_mode == "existing_directory",
            )

            # Send final result
            msg_payload = {
                'event': 'complete',
                'data': panel_data.model_dump(),
                'message': f"Panel '{panel_name}' added to schedule."
            }
            yield f"data: {json.dumps(msg_payload)}\n\n"

        except Exception as e:
            import traceback
            traceback.print_exc()
            yield f"data: {json.dumps({'event': 'error', 'detail': str(e)})}\n\n"

    return StreamingResponse(
        generate_events(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )

@app.get("/api/download-schedule")
async def download_schedule():
    if os.path.exists(OUTPUT_FILE):
        return FileResponse(OUTPUT_FILE, filename="Filled_Panel_Schedules.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    raise HTTPException(status_code=404, detail="No schedule file found.")

if __name__ == "__main__":
    import argparse
    import uvicorn
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()
    uvicorn.run(app, host="127.0.0.1", port=args.port)
