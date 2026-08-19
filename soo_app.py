"""
Turnover Package Tool - Desktop Application
Defense Fire Protection

Generates Sequence of Operations (SOO) and auto-fills CAN/ULC-S537 Verification/Inspection
reports from Autocall fire alarm panel CSV exports. Assembles a complete customer
turnover package as an interactive HTML document with tabbed navigation.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import json
import base64
import shutil
from collections import defaultdict, OrderedDict
from datetime import datetime

from fpdf import FPDF
from PIL import Image, ImageTk
from pypdf import PdfWriter, PdfReader

# ─── CONFIG / PERSISTENCE ───────────────────────────────────────────────────

APP_TITLE = "Turnover Package Tool"
APP_VERSION = "1.0.0"
GITHUB_REPO = "KGBAhart1/dfp-turnover-tool"
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".soo_extractor")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")

DEFAULT_CONFIG = {
    "company_name": "Defense Fire Protection",
    "logo_path": "",
    "last_directory": "",
    "technician_name": "",
    "technician_cert": "",
    "vi_template_path": "",
    "include_point_list": True,
    "include_initiating": True,
    "include_supervisory": True,
    "include_multi_criteria": True,
    "include_isolators": True,
    "include_outputs": True,
    "include_general_alarm": True,
    "include_trouble": True,
    "include_supervisory_seq": True,
    "include_reset_silence": True,
    "include_location_summary": True,
}


def load_config():
    os.makedirs(CONFIG_DIR, exist_ok=True)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                saved = json.load(f)
            return {**DEFAULT_CONFIG, **saved}
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)


# ─── AUTO-UPDATE CHECK ──────────────────────────────────────────────────────

def _parse_version(v):
    """'v1.2.3' or '1.2.3' -> (1, 2, 3), for simple tuple comparison."""
    v = v.strip().lstrip("vV")
    parts = []
    for p in v.split("."):
        m = re.match(r"\d+", p)
        parts.append(int(m.group()) if m else 0)
    return tuple(parts) or (0,)


def check_for_update(timeout=4):
    """Queries the GitHub Releases API for a newer version.
    Returns {"version": str, "url": str, "download_url": str|None} if a newer
    release exists, otherwise None. Never raises — network/API failures are
    treated as 'no update available'.
    """
    import urllib.request
    import urllib.error

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    try:
        req = urllib.request.Request(api_url, headers={"Accept": "application/vnd.github+json"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None

    tag = data.get("tag_name", "")
    if not tag:
        return None

    try:
        if _parse_version(tag) <= _parse_version(APP_VERSION):
            return None
    except Exception:
        return None

    download_url = None
    for asset in data.get("assets", []):
        name = asset.get("name", "")
        if name.lower().endswith(".exe"):
            download_url = asset.get("browser_download_url")
            break

    return {
        "version": tag,
        "url": data.get("html_url", f"https://github.com/{GITHUB_REPO}/releases/latest"),
        "download_url": download_url,
    }


# ─── DEVICE TYPE DEFINITIONS ────────────────────────────────────────────────

DEVICE_TYPES = {
    "PHOTO":    {"desc": "Photoelectric Smoke Detector",              "cat": "SMOKE",   "fn": "ALARM",   "vi_code": "S"},
    "MPHOTO":   {"desc": "Multi-Criteria Photoelectric Smoke Sensor", "cat": "SMOKE",   "fn": "ALARM",   "vi_code": "S"},
    "ION":      {"desc": "Ionization Smoke Detector",                 "cat": "SMOKE",   "fn": "ALARM",   "vi_code": "S"},
    "LASER":    {"desc": "Laser Smoke Detector",                      "cat": "SMOKE",   "fn": "ALARM",   "vi_code": "S"},
    "HEAT":     {"desc": "Heat Detector",                             "cat": "HEAT",    "fn": "ALARM",   "vi_code": "HT"},
    "MHEAT":    {"desc": "Multi-Criteria Heat Sensor",                "cat": "HEAT",    "fn": "ALARM",   "vi_code": "HT"},
    "MAHPO":    {"desc": "Multi-Criteria Heat/Photo Sensor",          "cat": "HEAT",    "fn": "ALARM",   "vi_code": "HT"},
    "MCOP":     {"desc": "Multi-Criteria CO/Photo Base",              "cat": "MULTI",   "fn": "SUPV",    "vi_code": "SFD"},
    "MCOGAS":   {"desc": "Multi-Criteria CO Gas Sensor",              "cat": "CO",      "fn": "SUPV",    "vi_code": "CO2"},
    "MAOFF":    {"desc": "Multi-Criteria Auxiliary Off",               "cat": "MULTI",   "fn": "SUPV",    "vi_code": "SFD"},
    "ADRPUL":   {"desc": "Addressable Manual Pull Station",           "cat": "PULL",    "fn": "ALARM",   "vi_code": "M"},
    "PULL":     {"desc": "Manual Pull Station",                       "cat": "PULL",    "fn": "ALARM",   "vi_code": "M"},
    "IDNISO2":  {"desc": "Addressable Isolator Module",               "cat": "ISO",     "fn": "MONITOR", "vi_code": "EM"},
    "IDNISO":   {"desc": "Addressable Isolator Module",               "cat": "ISO",     "fn": "MONITOR", "vi_code": "EM"},
    "IDNMON":   {"desc": "Addressable Monitor Module",                "cat": "MON",     "fn": "MONITOR", "vi_code": "SFD"},
    "MON":      {"desc": "Monitor Module",                            "cat": "MON",     "fn": "MONITOR", "vi_code": "SFD"},
    "RIAM":     {"desc": "Addressable Relay Module",                  "cat": "RELAY",   "fn": "OUTPUT",  "vi_code": "AD"},
    "RELAY":    {"desc": "Relay Module",                               "cat": "RELAY",   "fn": "OUTPUT",  "vi_code": "AD"},
    "NAC":      {"desc": "Notification Appliance Circuit",            "cat": "NAC",     "fn": "OUTPUT",  "vi_code": "NAC"},
    "TRUEALERT":{"desc": "TrueAlert Addressable Notification Appliance","cat": "NAC",   "fn": "OUTPUT",  "vi_code": "H/V"},
    "AV":       {"desc": "TrueAlert Addressable Horn/Strobe (Audio-Visual)","cat": "NAC","fn": "NOTIFY", "vi_code": "H/V"},
    "VO":       {"desc": "TrueAlert Addressable Strobe (Visual Only)","cat": "NAC",     "fn": "NOTIFY",  "vi_code": "V"},
    "AO":       {"desc": "TrueAlert Addressable Horn (Audio Only)",   "cat": "NAC",     "fn": "NOTIFY",  "vi_code": "H"},
    "TRUEALERT":{"desc": "TrueAlert Addressable Notification Appliance","cat": "NAC",   "fn": "NOTIFY",  "vi_code": "H/V"},
    "WFLOW":    {"desc": "Waterflow Switch",                          "cat": "WFLOW",   "fn": "ALARM",   "vi_code": "FD"},
    "TAMP":     {"desc": "Tamper Switch",                             "cat": "TAMPER",  "fn": "SUPV",    "vi_code": "SS"},
    "ISO":      {"desc": "Addressable Short Circuit Isolator",        "cat": "ISO",     "fn": "MONITOR", "vi_code": "EM"},
    "IAM":      {"desc": "Addressable Monitor Module",                "cat": "MON",     "fn": "MONITOR", "vi_code": "SFD"},
    "EOL":      {"desc": "End-of-Line Resistor",                      "cat": "EOL",     "fn": "SUPV",    "vi_code": "EOL"},
}

SIGNAL_LABELS = {
    "SMOKE": "Smoke Detection", "SUPDET": "Supervisory Detection",
    "COGASSUPV": "CO Gas Supervisory", "FAHPO": "Heat/Photo",
    "UTILITY": "Utility/Supervisory", "PULL": "Manual Fire Alarm",
    "ISO": "Isolator Module", "DUCT": "Duct Smoke Detection",
    "RRELAY": "Relay Output", "ALARM": "Fire Alarm",
    "WATERFLOW": "Waterflow Alarm", "TAMPER": "Tamper Supervisory",
}

CATEGORY_LABELS = {
    "SMOKE": "Smoke Detectors", "HEAT": "Heat Detectors",
    "MULTI": "Multi-Criteria Bases", "CO": "CO Detectors",
    "PULL": "Pull Stations", "ISO": "Isolator Modules",
    "MON": "Monitor Modules", "RELAY": "Relay Modules",
    "NAC": "Notification Appliances", "WFLOW": "Waterflow Devices",
    "TAMPER": "Tamper Devices",
}


# ─── S537 VI CHECKLIST DATA ─────────────────────────────────────────────────
# Each section: (sheet, title, yes_col, no_col, na_col, [(row, id, text, default), ...])
# Defaults: "Y"=YES, "N"=NO, "A"=N/A
# Sections with na_col=None only have YES/NO columns

VI_SECTIONS = [
    ("28", "C28.1 Verification Report", 10, 11, None, [
        (11, "A", "System provides single-stage operation", "Y"),
        (13, "B", "System provides two-stage operation", "N"),
        (15, "C", "Entire fire alarm system verified per CAN/ULC-S537-19", "Y"),
        (18, "D", "Partial verification for partial occupancy", "N"),
        (21, "E", "Components modified/replaced from different manufacturer, compatible", "N"),
        (23, "F", "Partial verification - system replaced in stages", "N"),
        (25, "G", "Verification of modified portion per Section 27", "N"),
        (28, "H", "Installed per design and CAN/ULC-S524", "Y"),
        (34, "I", "Fire alarm system documentation on site (per Section 29)", "Y"),
        (36, "J", "SOO specified in design confirmed, documentation on site", "Y"),
        (43, "K", "Fire alarm system is fully functional", "Y"),
        (48, "M", "Copy of report given to owner/representative", "Y"),
        (52, "N", "Fire alarm system has deficiencies noted in this report", "N"),
    ]),
    ("29", "C29 Documentation", 8, 9, 10, [
        (12, "A", "Instructions for resetting system and silencing alarm signals", "Y"),
        (14, "B", "Instructions for silencing trouble signal and action to be taken", "Y"),
        (17, "C", "Description of function of each operating control and indicator on FACP", "Y"),
        (20, "D", "Description of area/zone protected by each alarm detection circuit", "Y"),
        (24, "E", "Description of alarm signal operation", "Y"),
        (26, "F", "Description of ancillary equipment controlled by fire alarm system", "Y"),
        (28, "G", "Smoke control system sequence of operation (if applicable)", "A"),
        (34, "H", "Building diagrams indicating type/location of smoke control equipment", "Y"),
    ]),
    ("29_sub", "C29 Documentation (H sub-items)", 8, 9, 10, [
        (38, "H-i", "Sequence of Operation (see 5.9)", "Y"),
        (39, "H-ii", "Operating Instructions (see 5.9)", "Y"),
        (40, "H-iii", "Description of each type of field device", "Y"),
        (41, "H-iv", "Details of input to programmed output functions", "Y"),
        (43, "H-v", "Connection to fire signal receiving centre", "Y"),
        (45, "H-vi", "Previous Verification Report(s) if applicable", "Y"),
        (46, "H-vii", "As-Built drawings of building fire alarm system", "Y"),
        (48, "H-viii", "Copy of site-specific software (if applicable)", "A"),
    ]),
    ("30", "C30 Field Device Tests", 8, 9, 10, [
        (8, "A", "Correct field termination, conductor type and wire gauge per CSA C22.1 Sec 32", "Y"),
        (11, "B", "Correct field termination per manufacturer's installation instructions", "Y"),
        (14, "C", "Correct circuit polarities", "Y"),
        (16, "D", "Open circuit fault on conventional device circuit causes trouble signal", "Y"),
        (19, "E", "Removal of any active/supporting field device causes trouble signal", "Y"),
        (22, "F", "One contact and one non-contact device tested for operation and annunciation", "Y"),
        (26, "G", "Class A circuits tested for alarm on each side of open circuit fault", "A"),
        (31, "H", "Ground fault indications at electrically furthest device, no status change", "Y"),
        (35, "I", "Furthest field device receives rated power per manufacturer specs", "Y"),
        (39, "J", "Replaceable over-current devices correct rating", "Y"),
        (41, "K", "Multi-strand optical fibre not dedicated - system functions despite impairment", "A"),
        (46, "L", "Circuit fault tolerance - OPEN CIRCUIT confirmed", "Y"),
        (51, "M", "Circuit fault tolerance - SHORT CIRCUIT confirmed", "Y"),
        (62, "N", "Circuit fault tolerance - GROUND FAULT confirmed", "Y"),
        (66, "O", "Suite fault isolators - SHORT CIRCUIT confirmed", "A"),
        (71, "P", "Alarm condition confirmed on source side of shorted suite isolation module", "A"),
        (76, "Q", "Voice comm non-life-safety message fault detection maintained", "A"),
    ]),
    ("31", "C31 Operation Test - Circuit Fault Tolerance", 8, 9, 10, [
        (16, "A", "Each abnormal condition per Table 3.1 tested for each DCL", "Y"),
        (20, "B", "Alarm/trouble received under single ground fault on each DCL conductor", "Y"),
        (24, "C", "Class A DCL tested for alarm on each side of single open circuit fault", "Y"),
        (29, "D", "Wire-to-wire short on DCL without fault isolation - trouble and alarm confirmed", "Y"),
        (34, "E", "Fault isolators in DCL - short on isolated side, annunciation and source-side operation", "Y"),
        (41, "F", "Fault isolation between CU/transponders - fault confirmed, outside operation confirmed", "Y"),
    ]),
    ("32.1", "C32.1 Control Unit Inspection", 8, 9, 10, [
        (15, "A", "Input circuit designations correctly identified", "Y"),
        (18, "B", "Output circuit designations correctly identified", "Y"),
        (21, "C", "Correct designations for common control functions and indicators", "Y"),
        (23, "D", "Plug-in components and modules securely in place", "Y"),
        (25, "E", "Plug-in cables securely in place", "Y"),
        (33, "G", "Control unit clean and free of dust and dirt", "Y"),
        (35, "H", "Fuses in accordance with manufacturer's specification", "Y"),
        (37, "I", "Control unit lock functional", "Y"),
        (39, "J", "Termination points from wiring to field devices secure", "Y"),
        (41, "K", "Power disconnects per CSA C22.1 Canadian Electrical Code", "Y"),
        (45, "L", "Field wiring entry points per manufacturer's installation instructions", "Y"),
        (49, "M", "Main power supply feed wiring per manufacturer's specs", "Y"),
        (52, "N", "Stand alone capability serves same area for input and output circuits", "Y"),
        (60, "O", "Stand alone capability has silence, reset, trouble silence, degraded mode", "Y"),
        (65, "P", "Operating and maintenance instructions furnished", "Y"),
        (68, "Q", "Visual indicators comply with Table 3 Colour Code", "Y"),
    ]),
    ("32.2", "C32.2 Control Unit Testing", 8, 9, 10, [
        (81, "A", "Power ON visual indicator operates", "Y"),
        (83, "B", "Common visual trouble signal operates", "Y"),
        (85, "C", "Common audible trouble signal operates", "Y"),
        (87, "D", "Trouble signal silence switch operates", "Y"),
        (89, "E", "Main power supply failure trouble signal operates", "Y"),
        (91, "F", "Ground fault tested - positive and negative indication initiates trouble", "Y"),
        (94, "G", "Alert signal operates", "N"),
        (96, "H", "Alarm signal operates", "Y"),
        (98, "I", "Automatic transfer from alert to alarm operates", "N"),
        (101, "J", "Manual transfer from alert to alarm operates", "N"),
        (103, "K", "Automatic transfer cancel (acknowledge) on two-stage system", "N"),
        (106, "L", "Alarm signal silence inhibit function operates", "N"),
        (108, "M", "Alarm signal manual silence operates", "Y"),
        (110, "N", "Alarm signal silence visual indication operates", "Y"),
        (112, "O", "Silenced signals reinitiate upon subsequent alarm", "Y"),
        (116, "P", "Alarm signal silence automatic cut-out timer", "N"),
        (124, "Q", "Audible and visual alarm signals programmed per design", "Y"),
        (127, "R", "Input circuit alarm/supervisory operation operates", "Y"),
        (130, "S", "Input circuit supervision fault causes trouble", "Y"),
        (132, "T", "Output circuit alarm indicators operate", "Y"),
        (135, "U", "Output circuit supervision fault causes trouble", "Y"),
        (137, "V", "Visual indicator test (lamp test) operates", "Y"),
        (139, "W", "Coded signal sequences operate required number of times", "Y"),
        (142, "X", "Coded signal sequences not interrupted by subsequent alarms", "Y"),
        (144, "Y", "Ancillary device control circuit rated for intended purpose", "Y"),
        (146, "Z", "Ancillary device bypass results in trouble signal", "Y"),
        (148, "AA", "Input to output circuit operation per design and specification", "Y"),
        (152, "BB", "Fire alarm system reset operates", "Y"),
        (154, "CC", "Main power to emergency power transfer operates", "Y"),
        (156, "DD", "Control unit bonded to ground", "Y"),
        (158, "EE", "Status change confirmation (smoke detectors) verified", "N"),
    ]),
    ("32.3", "C32.3 Voice Communication Test", 8, 9, 10, [
        (167, "A", "Power ON indicator operates", "A"),
        (169, "B", "Common visual trouble signal operates", "A"),
        (171, "C", "Common audible trouble signal operates", "A"),
        (173, "D", "Trouble signal silence switch operates", "A"),
        (175, "E", "All-call voice paging operates", "A"),
        (177, "F", "Selective voice paging output circuits operate", "A"),
        (180, "G", "Selective voice paging trouble operation", "A"),
        (183, "H", "Microphone and press-to-talk operates", "A"),
        (185, "I", "Voice paging does not interfere with alert/alarm signal", "A"),
        (188, "J", "All-call voice paging operates on emergency power", "A"),
        (190, "K", "Amplifier failure auto-transfers to backup", "A"),
        (193, "L", "Emergency telephone call-in operation", "A"),
        (196, "M", "Emergency telephone two-way voice communication", "A"),
        (199, "N", "Emergency telephone trouble operation", "A"),
        (202, "O", "Emergency telephone verbal communication operates", "A"),
        (204, "P", "Emergency telephone operable/in-use tone", "A"),
        (206, "Q", "Voice comm bus fault detection maintained during standby paging", "A"),
    ]),
    ("32.5", "C32.5 Required System Response Times", 8, 9, 10, [
        (219, "A", "Manual fire alarm zone signals within 5 seconds", "Y"),
        (223, "B", "Signals operated within 10 seconds, subsequent input within 10 seconds", "Y"),
        (227, "C", "Remote connection operated within 10 seconds", "N"),
        (229, "D", "Releasing device start of sequence within 10 seconds", "N"),
        (231, "E", "Required annunciation within 10 seconds", "Y"),
        (234, "F", "Central alarm and control facility within 10 seconds", "Y"),
        (237, "G", "Ancillary circuits within 10 seconds, subsequent within 30 seconds", "Y"),
        (240, "H", "Trouble signal within 90 seconds", "Y"),
        (243, "I", "Water flow devices within 10 seconds", "Y"),
    ]),
    ("32.6", "C32.6 Large Scale Network Systems", 8, 9, 10, [
        (253, "A", "CU/transponders serve same area for input and output circuits", "A"),
        (256, "B", "Stand alone capability with silence, reset, trouble silence", "A"),
        (261, "C", "Single fault on network results in trouble and continued alarm", "A"),
        (266, "D", "DCL failure - each CU/transponder receives alarm and provides output", "A"),
        (272, "E", "DCL failure in two locations creating two segments - each functional", "A"),
    ]),
    ("32.7", "C32.7 Power Supply Inspection", 8, 9, 10, [
        (298, "A", "Conforms with CAN/ULC-S524 and CSA C22.1 Section 32", "Y"),
        (303, "B", "Fused per manufacturer's marked rating", "Y"),
        (305, "C", "Primary supply equipped with identified disconnect means", "Y"),
        (307, "D", "Adequate to meet system requirements", "Y"),
        (309, "E", "Ancillary device power from separate source", "N"),
        (313, "F", "Ancillary device power from CU designed to provide it", "N"),
        (316, "G", "Ancillary devices powered from CU are recorded", "N"),
    ]),
    ("32.8", "C32.8 Emergency Power Supply", 8, 9, 10, [
        (329, "A", "Correct battery type as recommended by manufacturer", "Y"),
        (331, "B", "Correct battery rating per battery calculations", "Y"),
        (345, "G", "Inspected for physical damage", "Y"),
        (347, "H", "Terminals cleaned and lubricated", "Y"),
        (349, "I", "Terminals clamped tightly", "Y"),
        (351, "J", "Correct electrolyte level", "N"),
        (353, "K", "Specific gravity within manufacturer's specs", "N"),
        (355, "L", "Electrolyte leakage", "N"),
        (357, "M", "Adequately ventilated", "Y"),
        (361, "O", "Disconnection causes trouble signal", "Y"),
        (380, "S", "Battery voltage not less than 85% of rating after tests", "Y"),
    ]),
    ("32.8_gen", "C32.8 Emergency Generator", 8, 9, 10, [
        (385, "GenA", "Generator provides power to AC circuit serving fire alarm", "N"),
        (387, "GenB", "Generator trouble results in trouble signal at annunciator", "N"),
        (391, "GenC", "Generator run results in trouble signal at annunciator", "N"),
        (395, "GenD", "Testing coordinated with generator tests", "N"),
        (397, "GenE", "Low fuel level results in trouble signal", "N"),
    ]),
    ("32.9", "C32.9 Annunciator Inspection", 8, 9, 10, [
        (416, "A", "Power ON visual indicator operates", "Y"),
        (418, "B", "Individual alarm/supervisory zones clearly indicated", "Y"),
        (420, "C", "Zone designation labels properly identified", "Y"),
        (422, "D", "Device labels correspond with actual field location", "Y"),
        (424, "E", "Common trouble signal operates", "Y"),
        (426, "F", "Visual indicator test (lamp test) operates", "Y"),
        (429, "G", "Input wiring from CU supervised and correct type/gauge", "Y"),
        (431, "H", "Alarm signal silence visual indicator operates", "Y"),
        (433, "I", "Ancillary function switches operate per design", "N"),
        (436, "J", "Ancillary function visual indicators operate", "N"),
        (438, "K", "Manual activation of alarm signal operates", "Y"),
        (441, "L", "Displays visible in installed location", "Y"),
        (443, "M", "Operates on emergency power", "Y"),
        (445, "N", "Visual indicators comply with Table 3 Colour Code", "Y"),
        (447, "O", "Multi-line sequential display operates per C5.9", "Y"),
    ]),
    ("32.10", "C32.10 Annunciators / Sequential Displays", 8, 9, 10, [
        (459, "A", "Power ON visual indicator operates", "Y"),
        (461, "B", "Zone designation labels properly identified", "Y"),
        (463, "C", "Individual alarm/supervisory indications properly identified", "Y"),
        (465, "D", "Device labels correspond with actual field location", "Y"),
        (467, "E", "Common trouble signal operates", "Y"),
        (469, "F", "Visual indicator test (lamp test) operates", "Y"),
        (472, "G", "Input wiring from CU supervised", "Y"),
        (474, "H", "Alarm signal silence visual indicator operates", "Y"),
        (476, "I", "Ancillary function switches operate per design", "N"),
        (479, "J", "Ancillary function visual indicators operate", "N"),
        (481, "K", "Manual activation of alarm signal operates", "Y"),
        (484, "L", "Displays visible in installed location", "Y"),
    ]),
    ("32.11", "C32.11 Remote Trouble Signal Unit", 8, 9, 10, [
        (493, "A", "Input wiring from CU supervised", "N"),
        (495, "B", "Visual trouble signal operates", "N"),
        (497, "C", "Audible trouble signal operates", "N"),
        (499, "D", "Audible trouble signal silence operates", "N"),
    ]),
    ("32.12", "C32.12 Printer Test", 8, 9, 10, [
        (508, "A", "Operates per design and specification", "N"),
        (510, "B", "Zone of each alarm initiating device correctly printed", "N"),
        (512, "C", "Rated voltage is present", "N"),
    ]),
]

VI_SECTIONS_32_14 = ("32.14", "C32.14 Interconnection to Fire Signal Receiving Centre", 8, 9, 10, [
    (7, "A", "FSRC transmitter integral to fire alarm control unit", "A"),
    (10, "B", "Interconnection between FACU and separate FSRC transmitter provided", "A"),
    (13, "C", "Demarcation terminal box with min 12 terminals installed", "A"),
    (18, "D", "Demarcation terminal box in same room as FACU", "A"),
    (21, "E", "Demarcation terminal box labeled Fire Alarm Demarcation", "A"),
    (24, "F", "Conductors between FACU and demarcation box comply with S524 Section 3.4", "A"),
    (29, "G", "Receipt of alarm transmission to FSRC", "A"),
    (31, "H", "Tested and confirmed operation of trouble relay", "A"),
    (33, "I", "Supervisory transmission to FSRC received", "A"),
    (36, "J", "Trouble transmission to FSRC received", "A"),
    (39, "K", "Disabling FSRC transmitter results in trouble signal at CU and to FSRC", "A"),
    (44, "L", "Disconnecting FSRC transmitter transmits trouble to FSRC", "A"),
])


# ─── PANEL OPERATING INSTRUCTIONS ───────────────────────────────────────────

OPERATING_INSTRUCTIONS_DIR = os.path.join(os.path.dirname(__file__), "OPERATING INSTRUCTIONS")

PANEL_SERIES_OPTIONS = ["4007ES", "4010ES", "4100ES"]

PANEL_SERIES_FILES = {
    "4007ES": "4007ES.pdf",
    "4010ES": "4010ES.pdf",
    "4100ES": "4100ES.pdf",
}

# Touch screen display instructions only apply to 4100ES and 4010ES, not 4007ES
TOUCHSCREEN_FILE = "4100ES & 4010ES Touch Screen Display.pdf"
TOUCHSCREEN_CAPABLE_SERIES = {"4010ES", "4100ES"}


def detect_panel_series(panel_model_text):
    """Guess the panel series (4007ES/4010ES/4100ES) from a detected/typed panel model string."""
    text = (panel_model_text or "").upper()
    if "4100" in text:
        return "4100ES"
    if "4010" in text:
        return "4010ES"
    return "4007ES"


def get_operating_instruction_paths(panel_series, include_touchscreen):
    """Return list of (label, full_path) for the operating instruction PDFs to bundle."""
    paths = []
    base_file = PANEL_SERIES_FILES.get(panel_series)
    if base_file:
        base_path = os.path.join(OPERATING_INSTRUCTIONS_DIR, base_file)
        if os.path.exists(base_path):
            paths.append((f"{panel_series} Operating Instructions", base_path))
    if include_touchscreen and panel_series in TOUCHSCREEN_CAPABLE_SERIES:
        ts_path = os.path.join(OPERATING_INSTRUCTIONS_DIR, TOUCHSCREEN_FILE)
        if os.path.exists(ts_path):
            paths.append(("Touch Screen Display Instructions", ts_path))
    return paths


# ─── CSV PARSING ────────────────────────────────────────────────────────────

CITY_RELAY_LABELS = {
    "CITY FIRE": "Fire Alarm",
    "CITY ALARM": "Fire Alarm",
    "CITY SUPV": "Supervisory",
    "CITY SUPERVISORY": "Supervisory",
    "CITY TBL": "Trouble",
    "CITY TROUBLE": "Trouble",
    "CITY AC FAIL": "AC Power Failure",
    "CITY PRI2": "Priority 2",
}


def _classify_city_relay(label):
    up = label.upper()
    for key, kind in CITY_RELAY_LABELS.items():
        if key in up:
            return kind
    return "Other"


def parse_csv(filepath):
    devices = []
    nac_circuits = []      # SIG3-SIG7 conventional NAC circuits, SIG900+ addressable zones
    city_relays = []       # 1-9 to 1-13 monitoring/city tie relays
    bypass_switches = []   # FUI/master bypass switches (3-x, 0-x with BYPASS in label)
    panel_model = "Autocall 4007"

    with open(filepath, "r", encoding="utf-8-sig") as f:
        lines = f.readlines()

    if len(lines) >= 2:
        line2 = lines[1].strip()
        if "4010" in line2:
            panel_model = "Autocall 4010"
        elif "4100" in line2:
            panel_model = "Autocall 4100"
        elif "4007H" in line2:
            panel_model = "Autocall 4007-H"
        elif "4007" in line2:
            panel_model = "Autocall 4007"

    for line in lines:
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split(",")]
        if len(parts) < 4:
            continue
        addr = parts[0].strip()
        dtype = parts[1].strip()
        sig = parts[2].strip()
        loc = parts[3].strip()
        if not addr or not dtype:
            continue

        if dtype in DEVICE_TYPES:
            info = DEVICE_TYPES[dtype]
            fn = info["fn"]
            # For flexible monitor modules, resolve function from signal type
            if dtype in ("IAM", "MON", "IDNMON", "RIAM", "RELAY"):
                sig_up = sig.upper()
                if any(kw in sig_up for kw in ["FIRE", "SMOKE", "HEAT", "ALARM", "WATERFLOW", "FLOW", "SUPPRESSION"]):
                    fn = "ALARM"
                elif any(kw in sig_up for kw in ["SUPV", "TAMPER", "SUPERVISORY", "VALVE", "PRESSURE"]):
                    fn = "SUPV"
                elif any(kw in sig_up for kw in ["RELAY", "OUTPUT", "CONTROL"]):
                    fn = "OUTPUT"
            devices.append({
                "address": addr, "device_type": dtype,
                "signal_type": sig, "location": loc,
                "desc": info["desc"], "cat": info["cat"], "fn": fn,
                "vi_code": info["vi_code"],
            })
            # Addressable TrueAlert on SIG circuits (e.g. SIG4-1) — register the parent circuit
            sig_sub = re.match(r"^(SIG\d+)-\d+$", addr, re.IGNORECASE)
            if sig_sub and info["cat"] == "NAC":
                parent = sig_sub.group(1).upper()
                if not any(c["id"] == parent for c in nac_circuits):
                    nac_circuits.append({
                        "id": parent, "label": f"TrueAlert Addressable Circuit {parent}",
                        "signal_type": "SQALERT", "addressable": True,
                    })
            continue

        # NAC/signal circuit definitions (SIGB type) - not individual devices
        if dtype == "SIGB" and addr.upper().startswith("SIG"):
            try:
                num = int(addr.upper().replace("SIG", ""))
            except ValueError:
                num = 0
            is_addressable = num >= 900
            nac_circuits.append({
                "id": addr, "label": loc, "signal_type": sig,
                "addressable": is_addressable,
            })
            continue

        # City/monitoring tie relays (SPECOUT type, "1-" address prefix)
        if dtype == "SPECOUT" and re.match(r"^1-\d+$", addr):
            if loc:
                city_relays.append({
                    "address": addr, "label": loc,
                    "kind": _classify_city_relay(loc),
                })
            continue

        # Bypass switches (PBSWITCH type with BYPASS in label)
        if dtype == "PBSWITCH" and "BYPASS" in loc.upper():
            bypass_switches.append({"address": addr, "label": loc})
            continue

    if not panel_model.startswith("Autocall 40"):
        if len(devices) > 254:
            panel_model = "Autocall 4010"

    extras = {
        "nac_circuits": nac_circuits,
        "city_relays": city_relays,
        "bypass_switches": bypass_switches,
    }
    return devices, panel_model, extras


def parse_custom_controls(pdf_path):
    """Parse Autocall Custom Control Detail PDF. Returns dict with system_type and equations."""
    try:
        reader = PdfReader(pdf_path)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        return {"error": str(e), "equations": [], "system_type": "single_stage", "features": []}

    equations = []
    current_eq = None
    section = None

    for line in text.split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        eq_match = re.match(r'Equation\s*:\s*(\d+)\s+Label\s*:\s*(.+)', stripped)
        if eq_match:
            if current_eq:
                equations.append(current_eq)
            current_eq = {"num": int(eq_match.group(1)), "label": eq_match.group(2).strip(),
                          "inputs": [], "outputs": [], "comments": ""}
            section = None
            continue
        if current_eq is None:
            continue
        if stripped.startswith("INPUTS:"):
            section = "inputs"
        elif stripped.startswith("OUTPUTS:"):
            section = "outputs"
        elif stripped == "END:":
            section = None
        elif stripped.startswith("COMMENTS:"):
            section = "comments"
        elif section in ("inputs", "outputs") and "|" in stripped:
            parts = [p.strip() for p in stripped.split("|")]
            if len(parts) >= 4:
                current_eq[section].append({"addr": parts[0], "type": parts[1],
                                             "subtype": parts[2], "desc": parts[3]})

    if current_eq:
        equations.append(current_eq)

    all_labels = " ".join(eq["label"].upper() for eq in equations)
    all_descs = " ".join(
        entry["desc"].upper()
        for eq in equations
        for lst in [eq["inputs"], eq["outputs"]]
        for entry in lst
    )
    # Also scan the raw program section headers/labels from the PDF
    prog_labels = " ".join(re.findall(r'Program\s*:\s*\d+\s+Label\s*:\s*(.+)', text, re.IGNORECASE)).upper()
    combined = all_labels + " " + all_descs + " " + prog_labels

    features = []
    system_type = "single_stage"

    if any(kw in combined for kw in ["2STAGE", "2 STAGE", "TWOSTAGE", "TWO STAGE",
                                      "STAGE1", "STAGE 1", "STAGE2", "STAGE 2",
                                      "ALERT TONE", "EVACUATION", "PRE-SIGNAL", "PRE SIGNAL",
                                      "2-STAGE", "TWO-STAGE"]):
        system_type = "two_stage"
        features.append("two_stage")

    if any(kw in combined for kw in ["AUTOSILENCE", "AUTO SILENCE", "AUTO SIL",
                                      "AUTOSIL", "SILENCE TIMER", "TIMED SILENCE"]):
        features.append("auto_silence")
        if system_type == "single_stage":
            system_type = "auto_silence"

    if any(kw in combined for kw in ["GROUND", "EARTH GND", "EARTH GROUND"]):
        features.append("ground_fault")

    if any(kw in combined for kw in ["BYPASS", "BYPASSED"]):
        features.append("bypass_logic")

    return {"system_type": system_type, "equations": equations, "features": features}


def parse_report_pdf(filepath):
    """Parse a 4100U Panel Programmer PDF Report. Returns same-shape data as parse_csv + cc_data."""
    reader = PdfReader(filepath)
    page_texts = [p.extract_text() or "" for p in reader.pages]
    full_text = "\n".join(page_texts)

    # ── General Info ────────────────────────────────────────────────────────
    # Job name from the "Job Name : XXXXX  Rev :" footer that appears on every page
    m = re.search(r'Job Name\s*:\s*(\S+)\s+Rev\s*:', full_text)
    job_name = m.group(1).strip() if m else ""

    # Full description (customer/property name) — strip the trailing weekday+date from footer
    m = re.search(
        r'Job Description\s*:\s*(.+?)\s+(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)',
        full_text)
    description = m.group(1).strip() if m else ""

    # Panel type — look for known Autocall model patterns in the full text
    model_m = re.search(r'\b(40(?:07(?:ES|H)?|09|10(?:ES|ADV)?|20|27|100(?:U|ES)?|120|127))\b',
                        full_text)
    panel_type = f"Autocall {model_m.group(1)}" if model_m else ""

    # Battery standby hours from General Info page
    page2 = page_texts[1] if len(page_texts) > 1 else ""
    batt_m = re.search(r'Hrs Standby Battery\s*:\s*\n+(\d+)', page2)
    battery_hrs = batt_m.group(1) if batt_m else "24"

    # ── Network Report Info ────────────────────────────────────────────────────
    node_num = None
    network_name = ""
    ext_points = []   # [{node, address, device_type, signal_type, label}]

    for text in page_texts:
        if "Network Report" not in text[:300]:
            continue
        m = re.search(r'Network Node\s*:\s*(\d+)', text)
        if m:
            node_num = int(m.group(1))
        m = re.search(r'Network Name\s*:\s*(\S+)', text)
        if m:
            network_name = m.group(1).strip()
        # EXT lines: "3571 EXT 1:M1-1-0 ADRPUL PULL PULL NODE 1"
        for line in text.split('\n'):
            em = re.match(r'^\d+\s+EXT\s+(\d+):(\S+)\s+(\S+)\s+(\S+)(?:\s+(.+))?$', line.strip())
            if em:
                ext_points.append({
                    "node":        int(em.group(1)),
                    "address":     em.group(2),
                    "device_type": em.group(3).upper(),
                    "signal_type": em.group(4).upper(),
                    "label":       (em.group(5) or "").strip(),
                })

    # ── Device Parsing from System Points Summary ────────────────────────────
    SKIP_DTYPES = {
        "PBSWITCH", "LED", "LEDRED", "LEDYELLOW", "LEDGREEN", "PIEZO",
        "SPECOUT", "MKYLED5", "MKYLED5CRY", "MKYLED5CGY", "MFLASH8",
        "MTBL6", "LIST", "MCODE7",
    }
    SKIP_PTYPES = {"SWITCH", "LED", "LEDRED", "LEDYELLOW", "LEDGREEN",
                   "PIEZO", "SPECOUT", "LCD", "SERV", "AUXPWR"}
    PDF_TYPE_MAP = {
        "PHOTO": "PHOTO", "MPHOTO": "MPHOTO", "ION": "ION", "LASER": "LASER",
        "HEAT": "HEAT", "MHEAT": "MHEAT", "MAHPO": "MAHPO",
        "ADRPUL": "ADRPUL", "PULL": "PULL",
        "IAM": "IAM", "IDNMON": "IDNMON", "MON": "MON",
        "RIAM": "RIAM", "RELAY": "RELAY",
        "ISO": "ISO", "IDNISO": "IDNISO", "IDNISO2": "IDNISO2",
        "IDNETISO": "ISO",  # IDNet isolator module
        "AV": "AV", "VO": "VO", "AO": "AO",
        "WFLOW": "WFLOW", "TAMP": "TAMP",
    }

    devices = []
    nac_circuits = []
    city_relays = []
    bypass_switches = []
    seen_nac_ids = set()
    EXIT_SECTIONS = (
        "Digital Pseudo Points", "Analog Pseudo Points", "Point List",
        "Custom Control", "Card Summary", "Zone/Relay",
    )
    in_points = False

    for text in page_texts:
        head = text.strip()[:200]
        if "System Points Summary" in head:
            in_points = True
        elif in_points and any(s in head for s in EXIT_SECTIONS):
            in_points = False
        if not in_points:
            continue

        for line in text.split('\n'):
            line = line.strip()
            if not line:
                continue
            if any(s in line for s in ("Job Name :", "Job Description :", "HW Ref",
                                        "System Points Summary", "Point Name", "LED Blink",
                                        "Device Type", "Custom Label", "Disable")):
                continue
            if re.match(r'^Page\s+\d+\s*$', line):
                continue

            # Parse: HW_Ref  Point_Name  Device_Type  Point_Type  [Label...]
            m = re.match(r'^(\S+)\s+(\S+)\s+(\S+)\s+(\S+)(?:\s+(.+))?$', line)
            if not m:
                continue
            hw_ref, point_name, dtype_raw, sig_raw = (
                m.group(1), m.group(2), m.group(3).upper(), m.group(4).upper())
            label = (m.group(5) or "").strip()

            # Skip LIST pseudo entries (L0, L1, ...)
            if re.match(r'^L\d+$', point_name):
                continue

            # City relays: MCODE7 RELAY with "City" in label
            if dtype_raw == "MCODE7" and sig_raw == "RELAY" and "CITY" in label.upper():
                city_relays.append({
                    "address": point_name, "label": label,
                    "kind": _classify_city_relay(label),
                })
                continue

            # Bypass switches
            if dtype_raw == "PBSWITCH" and "BYPASS" in label.upper():
                bypass_switches.append({"address": point_name, "label": label})
                continue

            if dtype_raw in SKIP_DTYPES or sig_raw in SKIP_PTYPES:
                continue

            # NAC / signal circuits: SIGB and MSIGB with SQALERT or SIGNAL
            if dtype_raw in ("SIGB", "MSIGB") and sig_raw in ("SQALERT", "SIGNAL"):
                nac_id = point_name.upper()
                if nac_id not in seen_nac_ids:
                    seen_nac_ids.add(nac_id)
                    nac_circuits.append({
                        "id": nac_id,
                        "label": label or nac_id,
                        "signal_type": sig_raw,
                        "addressable": dtype_raw == "MSIGB",
                    })
                continue

            # Skip 4009A4 TrueAlert expander parent cards (MSIGB sub-circuits handled above)
            if dtype_raw == "4009A4":
                continue

            # Map to DEVICE_TYPES key
            dtype_key = PDF_TYPE_MAP.get(dtype_raw)
            if not dtype_key or dtype_key not in DEVICE_TYPES:
                continue

            info = DEVICE_TYPES[dtype_key]
            fn = info.get("fn", "ALARM")

            # Dynamic fn resolution for monitor/relay modules
            if dtype_raw in ("IAM", "MON", "IDNMON"):
                sig_up = sig_raw
                if any(kw in sig_up for kw in ["FIRE", "SMOKE", "HEAT", "ALARM", "WATERFLOW", "FLOW"]):
                    fn = "ALARM"
                elif any(kw in sig_up for kw in ["SUPV", "TAMPER", "LATSUPV", "VALVE", "PRESSURE"]):
                    fn = "SUPV"
                elif any(kw in sig_up for kw in ["RELAY", "OUTPUT", "RRELAY"]):
                    fn = "OUTPUT"

            devices.append({
                "address": point_name, "device_type": dtype_key,
                "signal_type": sig_raw, "location": label, "fn": fn,
                "desc": info.get("desc", dtype_raw),
                "cat": info.get("cat", "INIT"),
            })

    # ── Custom Controls (reuse existing parser on same PDF) ──────────────────
    cc_data = None
    try:
        cc_data = parse_custom_controls(filepath)
    except Exception:
        pass

    return {
        "job_name": job_name,
        "description": description,
        "panel_type": panel_type,
        "battery_hrs": battery_hrs,
        "devices": devices,
        "nac_circuits": nac_circuits,
        "city_relays": city_relays,
        "bypass_switches": bypass_switches,
        "cc_data": cc_data,
        "node_num": node_num,
        "network_name": network_name,
        "ext_points": ext_points,
    }


def get_multi_criteria_groups(devices):
    groups = OrderedDict()
    for dev in devices:
        m = re.match(r"^(M\d+-\d+)", dev["address"])
        if m:
            base = m.group(1)
            if base not in groups:
                groups[base] = []
            groups[base].append(dev)
    return {k: v for k, v in groups.items() if len(v) > 1 and any(d["device_type"] == "MCOP" for d in v)}


# ─── PDF BASE CLASS ─────────────────────────────────────────────────────────

DARK = (33, 37, 41)
ACCENT = (0, 82, 155)
LIGHT_BG = (240, 244, 248)
WHITE = (255, 255, 255)
GRAY = (120, 120, 120)
DFP_ORANGE = (255, 112, 2)


class BasePDF(FPDF):
    def __init__(self, logo_path=None, company_name="", job_name="", doc_title="", **kwargs):
        super().__init__(**kwargs)
        self.logo_path = logo_path
        self.company_name = company_name
        self.job_name = job_name
        self.doc_title = doc_title
        self.set_auto_page_break(auto=True, margin=20)

    def header(self):
        if self.page_no() == 1:
            return
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(120, 120, 120)
        header_text = f"{self.company_name}  |  {self.job_name}  |  {self.doc_title}"
        self.cell(0, 8, header_text, new_x="LMARGIN", new_y="NEXT", align="C")
        self.set_draw_color(200, 200, 200)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(4)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")


# ─── PDF HELPERS ─────────────────────────────────────────────────────────────

def add_cover_logo(pdf, logo_path, y_start):
    if logo_path and os.path.exists(logo_path):
        try:
            img = Image.open(logo_path)
            w, h = img.size
            max_w, max_h = 70, 35
            ratio = min(max_w / w, max_h / h)
            disp_w = w * ratio
            disp_h = h * ratio
            x_logo = (210 - disp_w) / 2
            pdf.image(logo_path, x=x_logo, y=y_start, w=disp_w, h=disp_h)
            return y_start + disp_h + 10
        except Exception:
            pass
    return y_start


def pdf_section_title(pdf, text):
    pdf.ln(6)
    pdf.set_fill_color(*ACCENT)
    pdf.set_text_color(*WHITE)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 9, f"  {text}", new_x="LMARGIN", new_y="NEXT", fill=True)
    pdf.set_text_color(*DARK)
    pdf.ln(4)


def pdf_subsection_title(pdf, text):
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(*ACCENT)
    pdf.cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")
    pdf.set_draw_color(*ACCENT)
    pdf.line(15, pdf.get_y(), 120, pdf.get_y())
    pdf.set_text_color(*DARK)
    pdf.ln(3)


def pdf_body_text(pdf, text, bold=False):
    pdf.set_font("Helvetica", "B" if bold else "", 9)
    pdf.set_text_color(*DARK)
    pdf.multi_cell(0, 5, text, new_x="LMARGIN", new_y="NEXT")


def pdf_bullet(pdf, text, indent=20):
    pdf.set_x(indent)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(*DARK)
    pdf.multi_cell(0, 5, f"-  {text}", new_x="LMARGIN", new_y="NEXT")


def pdf_numbered(pdf, num, text, indent=20):
    pdf.set_x(indent)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*ACCENT)
    pdf.cell(8, 5, f"{num}.", new_x="END")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(*DARK)
    pdf.multi_cell(0, 5, f" {text}", new_x="LMARGIN", new_y="NEXT")


def pdf_table_header(pdf, cols):
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(*ACCENT)
    pdf.set_text_color(*WHITE)
    for label, w in cols:
        pdf.cell(w, 6, label, border=0, fill=True)
    pdf.ln()
    pdf.set_text_color(*DARK)


def pdf_table_row(pdf, vals, cols, alt=False):
    pdf.set_fill_color(*(LIGHT_BG if alt else WHITE))
    pdf.set_font("Helvetica", "", 8)
    for (_, w), val in zip(cols, vals):
        pdf.cell(w, 5, val, border=0, fill=True)
    pdf.ln()


def pdf_check_space(pdf, needed=40):
    if pdf.get_y() > (pdf.h - pdf.b_margin - needed):
        pdf.add_page()


# ─── SOO PDF GENERATION ─────────────────────────────────────────────────────

def generate_soo_pdf(devices, config, output_path, job_name, panel_model, notes="",
                      nac_circuits=None, bypass_switches=None):
    nac_circuits = nac_circuits or []
    bypass_switches = bypass_switches or []
    pdf = BasePDF(
        logo_path=config.get("logo_path"),
        company_name=config.get("company_name", ""),
        job_name=job_name, doc_title="Sequence of Operations",
        orientation="P", unit="mm", format="Letter",
    )
    pdf.alias_nb_pages()
    pdf.set_margins(15, 15, 15)

    # Cover page
    pdf.add_page()
    y_start = add_cover_logo(pdf, config.get("logo_path", ""), 40)
    pdf.set_y(y_start)
    pdf.set_font("Helvetica", "B", 28)
    pdf.set_text_color(*ACCENT)
    pdf.cell(0, 14, "SEQUENCE OF OPERATIONS", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 14)
    pdf.set_text_color(*DARK)
    pdf.cell(0, 10, "Fire Alarm System", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(8)
    pdf.set_draw_color(*ACCENT)
    pdf.set_line_width(0.8)
    pdf.line(60, pdf.get_y(), 150, pdf.get_y())
    pdf.ln(10)

    for label, value in [("Job Name", job_name), ("Panel Type", panel_model),
                          ("Company", config.get("company_name", "")),
                          ("Date", datetime.now().strftime("%B %d, %Y")),
                          ("Total Devices", str(len(devices)))]:
        pdf.set_font("Helvetica", "", 11)
        pdf.set_text_color(*GRAY)
        pdf.cell(55, 8, label + ":", align="R")
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(*DARK)
        pdf.cell(0, 8, "  " + value, new_x="LMARGIN", new_y="NEXT")

    if notes.strip():
        pdf.ln(10)
        pdf.set_font("Helvetica", "I", 10)
        pdf.set_text_color(*GRAY)
        pdf.multi_cell(0, 6, f"Notes: {notes.strip()}", new_x="LMARGIN", new_y="NEXT", align="C")

    by_cat = defaultdict(list)
    by_loc = defaultdict(list)
    for dev in devices:
        by_cat[dev["cat"]].append(dev)
        by_loc[dev["location"]].append(dev)
    mc_groups = get_multi_criteria_groups(devices)

    # System Overview
    pdf.add_page()
    pdf_section_title(pdf, "1. SYSTEM OVERVIEW")
    pdf_body_text(pdf, f"This document describes the sequence of operations for the {job_name} fire alarm system utilizing a {panel_model} fire alarm control panel.")
    pdf.ln(4)
    pdf_subsection_title(pdf, "Device Summary")
    cols = [("Category", 80), ("Count", 25)]
    pdf_table_header(pdf, cols)
    cat_order = ["SMOKE", "HEAT", "MULTI", "CO", "PULL", "ISO", "MON", "RELAY", "NAC", "WFLOW", "TAMPER", "OTHER"]
    row_idx = 0
    for cat in cat_order:
        if cat in by_cat:
            pdf_table_row(pdf, [CATEGORY_LABELS.get(cat, cat), str(len(by_cat[cat]))], cols, alt=(row_idx % 2 == 1))
            row_idx += 1
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(*ACCENT)
    pdf.set_text_color(*WHITE)
    pdf.cell(80, 6, "TOTAL", fill=True)
    pdf.cell(25, 6, str(len(devices)), fill=True)
    pdf.ln()
    pdf.set_text_color(*DARK)

    # Device Point List
    if config.get("include_point_list", True):
        pdf.add_page()
        pdf_section_title(pdf, "2. DEVICE POINT LIST")
        cols = [("Address", 28), ("Type", 22), ("Signal", 28), ("Location", 60), ("Description", 45)]
        pdf_table_header(pdf, cols)
        for i, dev in enumerate(devices):
            pdf_check_space(pdf, 8)
            if pdf.get_y() > pdf.h - 25:
                pdf.add_page()
                pdf_table_header(pdf, cols)
            pdf_table_row(pdf, [dev["address"], dev["device_type"], dev["signal_type"],
                 dev["location"][:30], dev["desc"][:25]], cols, alt=(i % 2 == 1))

    sec_num = 3

    # Initiating Devices
    if config.get("include_initiating", True):
        pdf.add_page()
        pdf_section_title(pdf, f"{sec_num}. INITIATING DEVICES - ALARM SEQUENCE")
        smoke_devs = by_cat.get("SMOKE", [])
        if smoke_devs:
            area = [d for d in smoke_devs if d["signal_type"] == "SMOKE"]
            duct = [d for d in smoke_devs if d["signal_type"] == "DUCT"]
            supdet = [d for d in smoke_devs if d["signal_type"] == "SUPDET"]
            if area:
                pdf_subsection_title(pdf, f"{sec_num}.1  Area Smoke Detectors")
                cols = [("Address", 30), ("Location", 80)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(area):
                    pdf_check_space(pdf, 8)
                    pdf_table_row(pdf, [d["address"], d["location"]], cols, alt=(i % 2 == 1))
                pdf.ln(4)
                pdf_body_text(pdf, "Activation Sequence:", bold=True)
                for i, s in enumerate(["Smoke detector senses products of combustion",
                    "Detector transmits ALARM signal to Autocall FACP",
                    "FACP annunciates FIRE ALARM condition",
                    "All NACs activate - audible EVACUATION tone and visual strobes",
                    "FACP transmits alarm signal to central monitoring station",
                    "Alarm LED illuminates; LCD displays device address, type, and location"], 1):
                    pdf_numbered(pdf, i, s)
            if duct:
                pdf_check_space(pdf, 60)
                pdf_subsection_title(pdf, f"{sec_num}.2  Duct Smoke Detectors")
                cols = [("Address", 30), ("Location", 80)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(duct):
                    pdf_table_row(pdf, [d["address"], d["location"]], cols, alt=(i % 2 == 1))
                pdf.ln(4)
                pdf_body_text(pdf, "Activation Sequence:", bold=True)
                for i, s in enumerate(["Duct smoke detector senses smoke in HVAC ductwork",
                    "Detector transmits ALARM signal to Autocall FACP",
                    "FACP annunciates FIRE ALARM condition",
                    "HVAC system shutdown initiated via relay output",
                    "All NACs activate - EVACUATION tone and strobes",
                    "FACP transmits alarm signal to central monitoring station"], 1):
                    pdf_numbered(pdf, i, s)
            if supdet:
                pdf_check_space(pdf, 40)
                pdf_subsection_title(pdf, f"{sec_num}.3  Supervisory Detection Sensors")
                cols = [("Address", 30), ("Location", 80)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(supdet):
                    pdf_table_row(pdf, [d["address"], d["location"]], cols, alt=(i % 2 == 1))
                pdf.ln(4)
                pdf_body_text(pdf, "Activation Sequence:", bold=True)
                for i, s in enumerate(["Multi-criteria smoke sensor detects products of combustion",
                    "Sensor transmits signal to multi-criteria base",
                    "FACP evaluates multi-criteria algorithm",
                    "If ALARM threshold met: full FIRE ALARM sequence initiated"], 1):
                    pdf_numbered(pdf, i, s)

        pull_devs = by_cat.get("PULL", [])
        if pull_devs:
            pdf_check_space(pdf, 60)
            pdf_subsection_title(pdf, f"{sec_num}.4  Manual Pull Stations")
            cols = [("Address", 30), ("Location", 80)]
            pdf_table_header(pdf, cols)
            for i, d in enumerate(pull_devs):
                pdf_table_row(pdf, [d["address"], d["location"]], cols, alt=(i % 2 == 1))
            pdf.ln(4)
            pdf_body_text(pdf, "Activation Sequence:", bold=True)
            for i, s in enumerate(["Occupant activates manual pull station",
                "Pull station transmits ALARM signal to Autocall FACP",
                "FACP annunciates FIRE ALARM condition",
                "All NACs activate - EVACUATION tone and strobes",
                "FACP transmits alarm signal to central monitoring station"], 1):
                pdf_numbered(pdf, i, s)

        heat_devs = by_cat.get("HEAT", [])
        if heat_devs:
            pdf_check_space(pdf, 60)
            pdf_subsection_title(pdf, f"{sec_num}.5  Heat Detectors / Multi-Criteria Heat Sensors")
            cols = [("Address", 30), ("Signal", 30), ("Location", 60)]
            pdf_table_header(pdf, cols)
            for i, d in enumerate(heat_devs):
                pdf_table_row(pdf, [d["address"], d["signal_type"], d["location"]], cols, alt=(i % 2 == 1))
            pdf.ln(4)
            pdf_body_text(pdf, "Activation Sequence:", bold=True)
            for i, s in enumerate(["Heat sensor detects rate-of-rise or fixed temperature threshold",
                "Sensor transmits signal to multi-criteria base (or direct alarm)",
                "FACP evaluates alarm condition",
                "If ALARM threshold met: full FIRE ALARM sequence initiated",
                "All NACs activate - EVACUATION tone and strobes",
                "FACP transmits alarm signal to central monitoring station"], 1):
                pdf_numbered(pdf, i, s)

        wflow_devs = by_cat.get("WFLOW", [])
        if wflow_devs:
            pdf_check_space(pdf, 60)
            pdf_subsection_title(pdf, f"{sec_num}.6  Waterflow Devices")
            cols = [("Address", 30), ("Location", 80)]
            pdf_table_header(pdf, cols)
            for i, d in enumerate(wflow_devs):
                pdf_table_row(pdf, [d["address"], d["location"]], cols, alt=(i % 2 == 1))
            pdf.ln(4)
            pdf_body_text(pdf, "Activation Sequence:", bold=True)
            for i, s in enumerate(["Waterflow switch detects flow in sprinkler piping",
                "Monitor module transmits ALARM signal to Autocall FACP",
                "FACP annunciates WATERFLOW ALARM condition",
                "All NACs activate - EVACUATION tone and strobes",
                "FACP transmits alarm signal to central monitoring station"], 1):
                pdf_numbered(pdf, i, s)
        sec_num += 1

    # Supervisory Devices
    if config.get("include_supervisory", True):
        co_devs = by_cat.get("CO", [])
        multi_devs = by_cat.get("MULTI", [])
        tamper_devs = by_cat.get("TAMPER", [])
        if co_devs or multi_devs or tamper_devs:
            pdf.add_page()
            pdf_section_title(pdf, f"{sec_num}. SUPERVISORY DEVICES")
            if co_devs:
                pdf_subsection_title(pdf, f"{sec_num}.1  Carbon Monoxide (CO) Gas Detection")
                cols = [("Address", 30), ("Location", 80)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(co_devs):
                    pdf_table_row(pdf, [d["address"], d["location"]], cols, alt=(i % 2 == 1))
                pdf.ln(4)
                pdf_body_text(pdf, "Activation Sequence:", bold=True)
                for i, s in enumerate(["CO gas sensor detects carbon monoxide above threshold",
                    "Sensor transmits SUPERVISORY signal to Autocall FACP",
                    "FACP annunciates CO GAS SUPERVISORY condition",
                    "FACP transmits supervisory signal to central monitoring station",
                    "Supervisory LED illuminates; LCD displays CO level and location"], 1):
                    pdf_numbered(pdf, i, s)
            if multi_devs:
                pdf_check_space(pdf, 40)
                pdf_subsection_title(pdf, f"{sec_num}.2  Multi-Criteria Detector Bases")
                cols = [("Address", 30), ("Type", 30), ("Location", 60)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(multi_devs):
                    pdf_table_row(pdf, [d["address"], d["device_type"], d["location"]], cols, alt=(i % 2 == 1))
                pdf.ln(3)
                pdf_body_text(pdf, "Multi-criteria detector bases (MCOP) serve as the communication platform for the individual sensing elements (smoke, CO, heat).")
            if tamper_devs:
                pdf_check_space(pdf, 40)
                pdf_subsection_title(pdf, f"{sec_num}.3  Tamper Switches")
                cols = [("Address", 30), ("Location", 80)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(tamper_devs):
                    pdf_table_row(pdf, [d["address"], d["location"]], cols, alt=(i % 2 == 1))
                pdf.ln(4)
                pdf_body_text(pdf, "Activation Sequence:", bold=True)
                for i, s in enumerate(["Tamper switch detects valve closure or abnormal position",
                    "Monitor module transmits SUPERVISORY signal to FACP",
                    "FACP annunciates SUPERVISORY condition",
                    "FACP transmits supervisory signal to central monitoring station"], 1):
                    pdf_numbered(pdf, i, s)
            sec_num += 1

    # Multi-Criteria Groups
    if config.get("include_multi_criteria", True) and mc_groups:
        pdf.add_page()
        pdf_section_title(pdf, f"{sec_num}. MULTI-CRITERIA DETECTOR GROUPS")
        pdf_body_text(pdf, "The following locations have multi-criteria detection combining smoke, CO gas, and heat sensing in a single detector assembly.")
        pdf.ln(3)
        for gi, (base_addr, group) in enumerate(mc_groups.items(), 1):
            pdf_check_space(pdf, 55)
            pdf_subsection_title(pdf, f"Group {gi}: {group[0]['location']}")
            cols = [("Address", 28), ("Description", 75), ("Signal", 40)]
            pdf_table_header(pdf, cols)
            for i, d in enumerate(group):
                pdf_table_row(pdf, [d["address"], d["desc"], SIGNAL_LABELS.get(d["signal_type"], d["signal_type"])], cols, alt=(i % 2 == 1))
            pdf.ln(3)
            pdf_body_text(pdf, "Sequence:", bold=True)
            pdf_bullet(pdf, "SMOKE Detection -> FIRE ALARM -> NACs activate -> Signal to monitoring")
            pdf_bullet(pdf, "CO Gas Detection -> SUPERVISORY -> Signal to monitoring")
            pdf_bullet(pdf, "Heat Detection -> FIRE ALARM -> NACs activate -> Signal to monitoring")
            pdf.ln(2)
        sec_num += 1

    # Isolator / Monitor Modules
    if config.get("include_isolators", True):
        iso_devs = by_cat.get("ISO", [])
        mon_devs = by_cat.get("MON", [])
        if iso_devs or mon_devs:
            pdf.add_page()
            pdf_section_title(pdf, f"{sec_num}. MONITOR / ISOLATOR MODULES")
            if iso_devs:
                pdf_subsection_title(pdf, f"{sec_num}.1  Isolator Modules")
                cols = [("Address", 30), ("Location", 80)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(iso_devs):
                    pdf_table_row(pdf, [d["address"], d["location"]], cols, alt=(i % 2 == 1))
                pdf.ln(4)
                pdf_body_text(pdf, "Function:", bold=True)
                pdf_body_text(pdf, "Isolator modules protect the SLC from short circuit faults. If a short occurs, the isolator isolates the faulted segment while maintaining communication with devices on unfaulted segments.")
                pdf.ln(3)
                pdf_body_text(pdf, "Activation Sequence:", bold=True)
                for i, s in enumerate(["Short circuit fault occurs on SLC wiring",
                    "Isolator module detects fault and isolates the faulted segment",
                    "FACP annunciates TROUBLE condition for isolated devices",
                    "Devices on unfaulted segments continue normal operation",
                    "When fault is cleared, isolator automatically restores the segment"], 1):
                    pdf_numbered(pdf, i, s)
            if mon_devs:
                pdf_check_space(pdf, 40)
                pdf_subsection_title(pdf, f"{sec_num}.2  Monitor Modules")
                cols = [("Address", 30), ("Signal", 30), ("Location", 60)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(mon_devs):
                    pdf_table_row(pdf, [d["address"], d["signal_type"], d["location"]], cols, alt=(i % 2 == 1))
            sec_num += 1

    # Output Devices
    if config.get("include_outputs", True):
        relay_devs = by_cat.get("RELAY", [])
        nac_devs = by_cat.get("NAC", [])
        if relay_devs or nac_devs or nac_circuits or bypass_switches:
            pdf_check_space(pdf, 60)
            pdf_section_title(pdf, f"{sec_num}. OUTPUT DEVICES")
            if relay_devs:
                pdf_subsection_title(pdf, f"{sec_num}.1  Relay Modules")
                cols = [("Address", 30), ("Signal", 30), ("Location", 60)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(relay_devs):
                    pdf_table_row(pdf, [d["address"], d["signal_type"], d["location"]], cols, alt=(i % 2 == 1))
                pdf.ln(4)
                for d in relay_devs:
                    pdf_body_text(pdf, f"{d['address']} - {d['location']}:", bold=True)
                    if any(kw in d["location"].upper() for kw in ["FURNACE", "HVAC", "SHUTDOWN", "AHU", "FAN"]):
                        for i, s in enumerate([f"Upon ANY FIRE ALARM condition on the system",
                            f"FACP activates relay module {d['address']}",
                            "Relay contact changes state (N.O. closes / N.C. opens)",
                            f"Connected equipment ({d['location']}) is shut down",
                            "Relay remains active until system is RESET"], 1):
                            pdf_numbered(pdf, i, s)
                    else:
                        pdf_body_text(pdf, f"  Relay activates upon alarm condition per programmed equations.")
                    pdf.ln(2)
            if nac_devs:
                pdf_check_space(pdf, 40)
                pdf_subsection_title(pdf, f"{sec_num}.2  Addressable Notification Appliances (TrueAlert)")
                cols = [("Address", 30), ("Location", 80)]
                pdf_table_header(pdf, cols)
                for i, d in enumerate(nac_devs):
                    pdf_table_row(pdf, [d["address"], d["location"]], cols, alt=(i % 2 == 1))

            conv_circuits = [c for c in nac_circuits if not c["addressable"]]
            if conv_circuits:
                pdf_check_space(pdf, 40)
                sub_n = "3" if nac_devs else "2"
                pdf_subsection_title(pdf, f"{sec_num}.{sub_n}  Conventional NAC Circuits")
                pdf_body_text(pdf, "These circuits carry conventional (non-addressable) notification appliances "
                                    "wired in series - horns, strobes, and horn/strobe combinations.")
                pdf.ln(2)
                cols = [("Circuit", 25), ("Description", 110)]
                pdf_table_header(pdf, cols)
                for i, c in enumerate(conv_circuits):
                    pdf_table_row(pdf, [c["id"], c["label"]], cols, alt=(i % 2 == 1))

            if bypass_switches:
                pdf_check_space(pdf, 30)
                pdf.ln(3)
                pdf_body_text(pdf, "Bypass Switches Programmed in This System:", bold=True)
                for b in bypass_switches:
                    pdf_bullet(pdf, f"{b['address']}: {b['label']}")
            sec_num += 1

    # General Alarm
    if config.get("include_general_alarm", True):
        pdf.add_page()
        pdf_section_title(pdf, f"{sec_num}. GENERAL ALARM SEQUENCE")
        pdf_body_text(pdf, "Upon activation of ANY alarm-initiating device:", bold=True)
        pdf.ln(2)
        for i, s in enumerate(["INITIATING DEVICE activates (smoke, heat, pull station, waterflow)",
            "Signal transmitted to Autocall FACP via SLC",
            "FACP processes signal and enters FIRE ALARM condition",
            "Panel: alarm LED, LCD displays address/type/location, piezo sounds",
            "NACs activate: audible EVACUATION tone, visual strobes",
            "Relay outputs activate per programming",
            "Alarm signal transmitted to central monitoring station",
            "Remote annunciator mirrors FACP display"], 1):
            pdf_numbered(pdf, i, s)
        relay_devs = by_cat.get("RELAY", [])
        if relay_devs:
            pdf.ln(3)
            pdf_body_text(pdf, "Relay outputs upon General Alarm:", bold=True)
            for d in relay_devs:
                pdf_bullet(pdf, f"{d['address']}: {d['location']}")
        sec_num += 1

    # Trouble
    if config.get("include_trouble", True):
        pdf_check_space(pdf, 60)
        pdf_section_title(pdf, f"{sec_num}. TROUBLE SEQUENCE")
        pdf_body_text(pdf, "Upon detection of ANY trouble condition:", bold=True)
        pdf.ln(2)
        for i, s in enumerate(["FACP detects fault (open circuit, ground fault, device removal, etc.)",
            "Trouble LED illuminates (amber), LCD displays trouble type/location, piezo sounds intermittently",
            "Trouble signal transmitted to central monitoring station",
            "System remains fully operational for all non-faulted circuits"], 1):
            pdf_numbered(pdf, i, s)
        sec_num += 1

    # Supervisory Sequence
    if config.get("include_supervisory_seq", True):
        pdf_check_space(pdf, 60)
        pdf_section_title(pdf, f"{sec_num}. SUPERVISORY SEQUENCE")
        pdf_body_text(pdf, "Upon activation of ANY supervisory device:", bold=True)
        pdf.ln(2)
        for i, s in enumerate(["Supervisory device activates (tamper, CO sensor, valve position, etc.)",
            "Signal transmitted to FACP via SLC",
            "FACP annunciates SUPERVISORY condition",
            "Supervisory LED, LCD displays device/location, piezo sounds",
            "Supervisory signal transmitted to central monitoring station",
            "Notification appliances DO NOT activate for supervisory conditions"], 1):
            pdf_numbered(pdf, i, s)
        sec_num += 1

    # Reset & Silence
    if config.get("include_reset_silence", True):
        pdf_check_space(pdf, 80)
        pdf_section_title(pdf, f"{sec_num}. SYSTEM RESET & ALARM SILENCE")
        pdf_subsection_title(pdf, "System Reset")
        pdf_body_text(pdf, "To reset the system after an alarm or trouble condition:", bold=True)
        pdf.ln(2)
        for i, s in enumerate(["Investigate and correct the cause",
            "Press SYSTEM RESET on the FACP",
            "All alarm/trouble LEDs extinguish",
            "NACs silence and deactivate",
            "Relay outputs return to normal state",
            "System returns to NORMAL operating condition",
            "Reset signal transmitted to monitoring station"], 1):
            pdf_numbered(pdf, i, s)
        pdf.ln(4)
        pdf_subsection_title(pdf, "Alarm Silence")
        pdf_body_text(pdf, "To silence audible notification appliances:", bold=True)
        pdf.ln(2)
        for i, s in enumerate(["Press ALARM SILENCE on the FACP",
            "Audible NACs silence (horns/speakers stop sounding)",
            "Visual NACs remain active (strobes continue flashing)",
            "Alarm condition remains annunciated on FACP display",
            "Subsequent NEW alarm will re-activate audible NACs"], 1):
            pdf_numbered(pdf, i, s)
        sec_num += 1

    # Location Summary
    if config.get("include_location_summary", True):
        pdf.add_page()
        pdf_section_title(pdf, f"{sec_num}. DEVICE LOCATION SUMMARY")
        sorted_locs = sorted(by_loc.keys())
        cols = [("Address", 28), ("Device Type", 50), ("Signal", 35)]
        for loc in sorted_locs:
            pdf_check_space(pdf, 30)
            pdf_subsection_title(pdf, loc)
            pdf_table_header(pdf, cols)
            for i, d in enumerate(by_loc[loc]):
                pdf_check_space(pdf, 8)
                pdf_table_row(pdf, [d["address"], d["desc"][:28], d["signal_type"]], cols, alt=(i % 2 == 1))

    # Footer
    pdf.add_page()
    pdf.ln(40)
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(*ACCENT)
    pdf.cell(0, 10, "END OF SEQUENCE OF OPERATIONS", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(6)
    pdf.set_draw_color(*ACCENT)
    pdf.line(60, pdf.get_y(), 150, pdf.get_y())
    pdf.ln(8)
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(*GRAY)
    pdf.cell(0, 8, f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", new_x="LMARGIN", new_y="NEXT", align="C")
    company = config.get("company_name", "")
    if company:
        pdf.cell(0, 8, company, new_x="LMARGIN", new_y="NEXT", align="C")

    pdf.output(output_path)
    return output_path


# ─── VI REPORT AUTO-FILL ────────────────────────────────────────────────────

def generate_vi_report(output_path, devices, vi_data, panel_model, checklist_overrides=None):
    """Generate a complete CAN/ULC-S537 Verification Report. Every field filled."""
    import openpyxl
    from openpyxl.cell.cell import MergedCell

    template_path = os.path.join(os.path.dirname(__file__), "vi_template_blank.xlsm")
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"VI template not found at {template_path}.\n"
            "Place vi_template_blank.xlsm in the same folder as this app."
        )

    if checklist_overrides is None:
        checklist_overrides = {}

    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path, keep_vba=True)

    def safe_set(ws, row, col, value):
        cell = ws.cell(row=row, column=col)
        if not isinstance(cell, MergedCell):
            cell.value = value

    def write_checklist(ws, section_def):
        """Write all checklist items for a section, using overrides or defaults."""
        sheet_key, _title, yes_col, no_col, na_col, items = section_def
        for row_num, item_id, _text, default in items:
            override_key = f"{sheet_key}_{item_id}"
            value = checklist_overrides.get(override_key, default)
            # Clear all columns first
            safe_set(ws, row_num, yes_col, "")
            safe_set(ws, row_num, no_col, "")
            if na_col:
                safe_set(ws, row_num, na_col, "")
            # Set the chosen column
            if value == "Y":
                safe_set(ws, row_num, yes_col, "P")
            elif value == "N":
                safe_set(ws, row_num, no_col, "P")
            elif value == "A" and na_col:
                safe_set(ws, row_num, na_col, "P")

    # ── Write all checklist sections ──
    for sec in VI_SECTIONS:
        sheet_key = sec[0]
        # Map section keys to actual sheet names
        if sheet_key in ("28",):
            write_checklist(wb['28'], sec)
        elif sheet_key in ("29", "29_sub"):
            write_checklist(wb['29'], sec)
        elif sheet_key == "30":
            write_checklist(wb['30'], sec)
        elif sheet_key == "31":
            write_checklist(wb['31'], sec)
        elif sheet_key.startswith("32."):
            write_checklist(wb['32.1-32.12'], sec)

    # Write 32.14 separately (own sheet)
    write_checklist(wb['32.14'], VI_SECTIONS_32_14)

    # ── Sheet 28: Header fields ──
    ws28 = wb['28']
    safe_set(ws28, 2, 3, vi_data.get("building_name", ""))
    safe_set(ws28, 3, 3, vi_data.get("building_address", ""))
    safe_set(ws28, 4, 3, vi_data.get("date", ""))
    model_short = vi_data.get("model_number", "")
    if not model_short and panel_model:
        model_short = panel_model.replace("Autocall ", "A0").replace("-H", "H")
    safe_set(ws28, 5, 9, model_short)
    safe_set(ws28, 6, 3, vi_data.get("building_permit", ""))
    safe_set(ws28, 6, 9, vi_data.get("electrical_permit", ""))
    # Edition fields
    safe_set(ws28, 30, 10, vi_data.get("s524_edition", "2019"))
    safe_set(ws28, 31, 10, vi_data.get("building_code_edition", "2023"))
    safe_set(ws28, 32, 10, vi_data.get("cec_edition", "2024"))
    # Signatures
    safe_set(ws28, 55, 1, vi_data.get("tech_name", ""))
    safe_set(ws28, 55, 5, vi_data.get("tech_company", "Defense Fire Protection Inc."))
    safe_set(ws28, 55, 9, vi_data.get("tech_phone", "403-807-5822"))
    safe_set(ws28, 59, 5, vi_data.get("tech_cert", ""))
    safe_set(ws28, 50, 8, vi_data.get("owner_rep", ""))
    safe_set(ws28, 62, 1, vi_data.get("assist_tech_name", ""))
    safe_set(ws28, 62, 5, vi_data.get("assist_company", ""))
    safe_set(ws28, 62, 9, vi_data.get("assist_phone", ""))
    safe_set(ws28, 66, 5, vi_data.get("assist_cert", ""))
    safe_set(ws28, 69, 1, vi_data.get("designer_name", ""))
    safe_set(ws28, 69, 5, vi_data.get("designer_company", ""))
    # Comments
    safe_set(ws28, 45, 2, vi_data.get("comments", ""))

    # ── Sheet 29: Documentation location ──
    ws29 = wb['29']
    safe_set(ws29, 50, 7, vi_data.get("doc_location", "PANEL INSTRUCTIONS IN FACP"))

    # ── Sheet 32.1-32.12: All variable fields ──
    ws32 = wb['32.1-32.12']
    panel_loc = vi_data.get("panel_location", "")
    panel_id = vi_data.get("panel_id", "FACP")
    for r in [9, 75, 292, 322, 410, 453]:
        safe_set(ws32, r, 5, panel_loc)
    for r in [11, 77, 293, 323, 412, 455]:
        safe_set(ws32, r, 5, panel_id)
    safe_set(ws32, 294, 5, vi_data.get("disconnect_location", ""))
    safe_set(ws32, 295, 5, vi_data.get("breaker_id", ""))
    safe_set(ws32, 27, 8, vi_data.get("date", ""))
    safe_set(ws32, 29, 9, vi_data.get("firmware_rev", ""))
    safe_set(ws32, 31, 9, vi_data.get("firmware_ver", ""))
    # Battery
    safe_set(ws32, 325, 5, vi_data.get("battery_type", ""))
    batt_qty = vi_data.get("battery_qty", "")
    safe_set(ws32, 325, 10, int(batt_qty) if batt_qty else None)
    for field, row, col in [
        ("voltage_on", 334, 8), ("voltage_off", 337, 6), ("current_off", 337, 8),
        ("voltage_alarm", 340, 6), ("current_alarm", 340, 8), ("charging_current", 343, 8),
        ("voltage_after_test", 378, 8),
    ]:
        val = vi_data.get(field, "")
        safe_set(ws32, row, col, float(val) if val else None)
    safe_set(ws32, 359, 8, vi_data.get("battery_date", ""))
    batt_cap = vi_data.get("battery_capacity", "")
    safe_set(ws32, 376, 8, float(batt_cap) if batt_cap else None)

    # Battery test type checkboxes
    batt_test = vi_data.get("battery_test_type", "24h")
    safe_set(ws32, 364, 8, "P" if batt_test == "24h" else "")
    safe_set(ws32, 366, 8, "P" if batt_test == "silent" else "")
    safe_set(ws32, 368, 8, "P" if batt_test == "accelerated" else "")

    # Alarm operation time
    safe_set(ws32, 326, 6, "P" if vi_data.get("alarm_time", "30") == "30" else "")
    safe_set(ws32, 326, 8, "P" if vi_data.get("alarm_time", "30") == "60" else "")
    safe_set(ws32, 326, 10, "P" if vi_data.get("alarm_time", "30") == "120" else "")

    # ── Sheet 31: Operation Test ──
    ws31 = wb['31']
    safe_set(ws31, 6, 5, panel_loc)
    safe_set(ws31, 8, 5, panel_id)
    safe_set(ws31, 10, 5, vi_data.get("dcl_id", "M1"))

    # ── Sheet 33.1: Device Legend (model numbers) ──
    ws331 = wb['33.1']
    for code, model_key, row in [
        ("M", "model_pull", 10), ("HT", "model_heat", 12),
        ("S", "model_smoke", 13), ("V", "model_strobe", 28),
        ("H/V", "model_hornstrobe", 37), ("EOL", "model_eol", 33),
        ("CO2", "model_co", 45),
    ]:
        val = vi_data.get(model_key, "")
        if val:
            safe_set(ws331, row, 9, val)

    # ── Sheet 33.2: Individual Device Record ──
    ws332 = wb['33.2']
    vi_devices = [d for d in devices if d["device_type"] not in ("MCOP", "MAOFF")]
    for i, dev in enumerate(vi_devices):
        row = 10 + i
        vi_code = DEVICE_TYPES.get(dev["device_type"], {}).get("vi_code", dev["device_type"])
        safe_set(ws332, row, 1, "")
        safe_set(ws332, row, 2, f" {dev['location']}")
        safe_set(ws332, row, 3, f" {vi_code}")
        safe_set(ws332, row, 5, f" {dev['address']}")
        safe_set(ws332, row, 6, "P")
        safe_set(ws332, row, 7, "P")
        safe_set(ws332, row, 9, "P")
        safe_set(ws332, row, 10, "P")
        safe_set(ws332, row, 11, "P")

    # ── Sheet 32.4: Non-Life Safety Message Circuit ──
    if '32.4' in wb.sheetnames:
        ws324 = wb['32.4']
        safe_set(ws324, 5, 1, vi_data.get("speakers_installed", "NO SPEAKERS"))

    # ── Sheet 32.13: Ancillary Devices ──
    ws3213 = wb['32.13']
    anc_row = 4
    # Auto-fill from CSV relay devices
    relay_devs = [d for d in devices if d["cat"] == "RELAY"]
    for dev in relay_devs:
        safe_set(ws3213, anc_row, 1, dev["location"])
        safe_set(ws3213, anc_row, 5, "X")
        safe_set(ws3213, anc_row, 7, "VISUAL")
        anc_row += 1
    # Add manual ancillary devices from UI
    for anc in vi_data.get("ancillary_devices", []):
        safe_set(ws3213, anc_row, 1, anc.get("device", ""))
        powered = anc.get("powered_by", "")
        safe_set(ws3213, anc_row, 4, powered)
        confirmed = anc.get("confirmed", "Yes")
        safe_set(ws3213, anc_row, 5, "X" if confirmed == "Yes" else "")
        safe_set(ws3213, anc_row, 6, "X" if confirmed == "No" else "")
        safe_set(ws3213, anc_row, 7, anc.get("method", "VISUAL"))
        anc_row += 1

    # ── Sheet 33.4: Circuit Fault Tolerance ──
    ws334 = wb['33.4']
    entries_334 = vi_data.get("fault_tolerance_tests", [])
    if entries_334:
        for i, ft in enumerate(entries_334):
            row = 6 + i
            safe_set(ws334, row, 1, ft.get("fault_loc", ""))
            safe_set(ws334, row, 2, ft.get("short", "X"))
            safe_set(ws334, row, 3, ft.get("open", "X"))
            safe_set(ws334, row, 4, ft.get("ground", "X"))
            safe_set(ws334, row, 5, ft.get("isolation", ""))
            safe_set(ws334, row, 6, ft.get("nonfault_dev", ""))
            safe_set(ws334, row, 7, ft.get("result", "P"))
    else:
        # Auto-generate from device areas if no manual entries
        pull_devs = [d for d in devices if d["cat"] == "PULL"]
        if pull_devs:
            test_row = 6
            seen_areas = set()
            for dev in devices:
                area = dev["location"].split()[0] if dev["location"] else ""
                if area and area not in seen_areas and test_row < 20:
                    seen_areas.add(area)
                    safe_set(ws334, test_row, 1, dev["location"])
                    safe_set(ws334, test_row, 2, "X")
                    safe_set(ws334, test_row, 3, "X")
                    safe_set(ws334, test_row, 4, "X")
                    safe_set(ws334, test_row, 7, "P")
                    test_row += 1

    # ── Remarks ──
    wsR = wb['Remarks']
    remarks = vi_data.get("remarks", "")
    if remarks:
        safe_set(wsR, 2, 1, remarks)

    # ── Sheet 33.2: Manual (non-addressable) devices ──
    manual_devices = vi_data.get("manual_devices", [])
    manual_start_row = 10 + len(vi_devices)
    for i, md in enumerate(manual_devices):
        row = manual_start_row + i
        safe_set(ws332, row, 1, f" {md.get('floor', '')}")
        safe_set(ws332, row, 2, f" {md.get('location', '')}")
        safe_set(ws332, row, 3, f" {md.get('type', '')}")
        safe_set(ws332, row, 5, f" {md.get('circuit', '')}")
        safe_set(ws332, row, 6, "")
        safe_set(ws332, row, 7, md.get("installed", "P") or "")
        safe_set(ws332, row, 9, md.get("operation", "P") or "")
        safe_set(ws332, row, 10, "")
        safe_set(ws332, row, 11, md.get("supervision", "P") or "")

    # ── Sheet 28: FSRC connection flag + name/phone ──
    fsrc_connected = vi_data.get("fsrc_connected", "No") == "Yes"
    safe_set(ws28, 39, 10, "P" if fsrc_connected else "")
    safe_set(ws28, 39, 11, "" if fsrc_connected else "P")
    safe_set(ws28, 40, 3, vi_data.get("fsrc_name", ""))
    safe_set(ws28, 41, 3, vi_data.get("fsrc_phone", ""))

    # ── Sheet 32.14: "No Interconnection" flag + FSRC name/phone ──
    ws3214 = wb['32.14']
    safe_set(ws3214, 5, 7, "" if fsrc_connected else "P")
    safe_set(ws3214, 48, 7, vi_data.get("fsrc_name", ""))
    safe_set(ws3214, 49, 8, vi_data.get("fsrc_phone", ""))

    # ── Sheet 33.3: Subsequent Alarm Activation Test ──
    if '33.3' in wb.sheetnames:
        ws333 = wb['33.3']
        alarm_tests = vi_data.get("alarm_tests_333", [])
        for i, at in enumerate(alarm_tests):
            row = 3 + i
            safe_set(ws333, row, 1, at.get("init_zone", ""))
            safe_set(ws333, row, 2, at.get("init_dev", ""))
            safe_set(ws333, row, 3, at.get("sub_zone", ""))
            safe_set(ws333, row, 4, at.get("sub_dev", ""))
            silent = at.get("silent", "No")
            safe_set(ws333, row, 5, "X" if silent == "Yes" else "")
            safe_set(ws333, row, 6, "X" if silent != "Yes" else "")

    # ── Sheet 33.5: Sound Level Measurements ──
    if '33.5' in wb.sheetnames:
        ws335 = wb['33.5']
        sound_data = vi_data.get("sound_levels", [])
        for i, sd in enumerate(sound_data):
            row = 3 + i
            safe_set(ws335, row, 1, sd.get("zone", ""))
            safe_set(ws335, row, 2, sd.get("location", ""))
            safe_set(ws335, row, 3, sd.get("ambient", ""))
            safe_set(ws335, row, 4, sd.get("alarm", ""))
            safe_set(ws335, row, 5, sd.get("remarks", ""))

    wb.save(output_path)
    return output_path


# ─── PDF MERGING ─────────────────────────────────────────────────────────────

def merge_pdfs(pdf_paths, output_path):
    writer = PdfWriter()
    for path in pdf_paths:
        reader = PdfReader(path)
        for page in reader.pages:
            writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


# ─── HTML TURNOVER PACKAGE ──────────────────────────────────────────────────

def generate_turnover_html(output_path, soo_html_content, doc_tabs):
    """
    Generate a self-contained HTML turnover package with tabbed navigation.
    doc_tabs: list of {"name": "Tab Name", "type": "html"|"pdf", "content": str}
    For html type, content is raw HTML. For pdf type, content is base64-encoded PDF.
    """
    tabs_js = json.dumps([{"name": t["name"], "type": t["type"]} for t in doc_tabs])

    pdf_data_blocks = []
    for i, tab in enumerate(doc_tabs):
        if tab["type"] == "pdf":
            pdf_data_blocks.append(f'const pdfData_{i} = "{tab["content"]}";')
        elif tab["type"] == "html":
            escaped = tab["content"].replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${").replace("</script>", "<\\/script>")
            pdf_data_blocks.append(f'const htmlData_{i} = `{escaped}`;')

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Turnover Package | Defense Fire Protection</title>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&family=Roboto:wght@300;400;500&display=swap" rel="stylesheet">
<style>
:root {{
  --primary: #FF7002;
  --dark: #232728;
  --text: #112337;
  --text-light: #69727d;
  --bg: #f8f6f4;
  --white: #ffffff;
  --border: #e0d8d2;
  --shadow: 0 4px 24px rgba(0,0,0,0.08);
  --radius: 12px;
}}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Roboto',sans-serif; background:var(--bg); color:var(--text); }}
h1,h2,h3 {{ font-family:'Poppins',sans-serif; }}

.tp-header {{
  background:var(--dark); padding:0 32px; height:64px;
  display:flex; align-items:center; justify-content:space-between;
  box-shadow:0 2px 12px rgba(0,0,0,0.2);
}}
.tp-header .brand {{ display:flex; align-items:center; gap:12px; }}
.tp-header .brand img {{ height:36px; }}
.tp-header .brand span {{ color:var(--white); font-family:'Poppins'; font-weight:600; }}
.tp-header .title {{ color:var(--primary); font-family:'Poppins'; font-weight:600; font-size:1.1rem; }}

.tab-bar {{
  background:var(--white); border-bottom:1px solid var(--border);
  padding:0 32px; display:flex; gap:0; overflow-x:auto;
  position:sticky; top:0; z-index:100;
}}
.tab-btn {{
  padding:14px 24px; border:none; background:none; cursor:pointer;
  font-family:'Poppins'; font-size:0.9rem; font-weight:500;
  color:var(--text-light); border-bottom:3px solid transparent;
  transition:all 0.2s; white-space:nowrap;
}}
.tab-btn:hover {{ color:var(--primary); }}
.tab-btn.active {{ color:var(--primary); border-bottom-color:var(--primary); }}

.tab-content {{ display:none; }}
.tab-content.active {{ display:block; }}

.pdf-container {{ width:100%; height:calc(100vh - 130px); }}
.pdf-container iframe, .pdf-container object {{
  width:100%; height:100%; border:none;
}}

.html-container {{ width:100%; min-height:calc(100vh - 130px); }}
.html-container iframe {{
  width:100%; height:calc(100vh - 130px); border:none;
}}

.tp-footer {{
  background:var(--dark); color:rgba(255,255,255,0.5);
  text-align:center; padding:20px; font-size:0.8rem;
}}
.tp-footer strong {{ color:var(--primary); }}

@media print {{
  .tp-header, .tab-bar {{ display:none; }}
  .tab-content {{ display:block !important; }}
}}
</style>
</head>
<body>

<header class="tp-header">
  <div class="brand">
    <img src="https://defensefirepro.com/wp-content/uploads/2025/12/qtq_95.webp" alt="DFP" onerror="this.style.display='none'">
    <span>Defense Fire Protection</span>
  </div>
  <span class="title">Customer Turnover Package</span>
</header>

<div class="tab-bar" id="tab-bar"></div>

<div id="tab-panels"></div>

<footer class="tp-footer">
  <p>Customer Turnover Package by <strong>Defense Fire Protection</strong></p>
  <p style="margin-top:4px;"><a href="https://defensefirepro.com" target="_blank" style="color:var(--primary);">defensefirepro.com</a></p>
</footer>

<script>
const tabs = {tabs_js};
{''.join(pdf_data_blocks)}

const tabBar = document.getElementById('tab-bar');
const tabPanels = document.getElementById('tab-panels');

tabs.forEach((tab, i) => {{
  // Create tab button
  const btn = document.createElement('button');
  btn.className = 'tab-btn' + (i === 0 ? ' active' : '');
  btn.textContent = tab.name;
  btn.onclick = () => switchTab(i);
  tabBar.appendChild(btn);

  // Create tab panel
  const panel = document.createElement('div');
  panel.className = 'tab-content' + (i === 0 ? ' active' : '');
  panel.id = 'tab-' + i;

  if (tab.type === 'pdf') {{
    panel.className += ' pdf-container';
    const b64 = eval('pdfData_' + i);
    const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    const blob = new Blob([bytes], {{type: 'application/pdf'}});
    const url = URL.createObjectURL(blob);
    const iframe = document.createElement('iframe');
    iframe.src = url;
    iframe.style.width = '100%';
    iframe.style.height = 'calc(100vh - 130px)';
    iframe.style.border = 'none';
    panel.appendChild(iframe);
  }} else if (tab.type === 'html') {{
    panel.className += ' html-container';
    const iframe = document.createElement('iframe');
    iframe.srcdoc = eval('htmlData_' + i);
    iframe.sandbox = 'allow-scripts allow-same-origin';
    panel.appendChild(iframe);
  }}

  tabPanels.appendChild(panel);
}});

function switchTab(idx) {{
  document.querySelectorAll('.tab-btn').forEach((b,i) => b.classList.toggle('active', i===idx));
  document.querySelectorAll('.tab-content').forEach((p,i) => p.classList.toggle('active', i===idx));
}}
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    return output_path


# ─── GUI APPLICATION ────────────────────────────────────────────────────────

class TurnoverApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1100x780")
        self.root.minsize(960, 650)

        self.config = load_config()
        self.devices = []
        self.nodes = []           # list of per-node dicts from parse_report_pdf
        self.csv_path = None
        self.panel_model = "Autocall 4007"
        self.logo_image = None
        self.attached_docs = []  # list of {"path": str, "name": str}
        self.nac_circuits = []
        self.city_relays = []
        self.bypass_switches = []
        self.cc_pdf_path = None
        self.cc_data = None

        self._build_ui()
        self._load_logo_preview()

    def _build_ui(self):
        DFP_ORANGE = "#FF7002"
        DFP_DARK   = "#232728"
        DFP_BG     = "#f4f5f7"
        DFP_WHITE  = "#ffffff"

        self.root.configure(bg=DFP_DARK)

        style = ttk.Style()
        style.theme_use("clam")

        # Global widget defaults
        style.configure(".",             background=DFP_BG,    foreground="#1a1a1a",
                        font=("Segoe UI", 9))
        style.configure("TFrame",        background=DFP_BG)
        style.configure("TLabel",        background=DFP_BG,    foreground="#1a1a1a",
                        font=("Segoe UI", 9))
        style.configure("TLabelframe",   background=DFP_BG,    relief="flat",
                        bordercolor="#d0d5dd", borderwidth=1)
        style.configure("TLabelframe.Label", background=DFP_BG, foreground=DFP_DARK,
                        font=("Segoe UI", 9, "bold"))
        style.configure("TEntry",        fieldbackground=DFP_WHITE, relief="flat",
                        bordercolor="#c8cdd5", borderwidth=1, padding=4)
        style.configure("TCombobox",     fieldbackground=DFP_WHITE, relief="flat",
                        bordercolor="#c8cdd5", padding=4)
        style.configure("TCheckbutton",  background=DFP_BG,    foreground="#1a1a1a")
        style.configure("TSeparator",    background="#d0d5dd")
        style.configure("TScrollbar",    background="#d0d5dd", troughcolor=DFP_BG,
                        arrowcolor="#666")

        # Notebook
        style.configure("TNotebook",           background=DFP_DARK, tabmargins=[0, 0, 0, 0])
        style.configure("TNotebook.Tab",       background="#3a3d3f", foreground="#aaaaaa",
                        font=("Segoe UI", 9, "bold"), padding=[16, 8])
        style.map("TNotebook.Tab",
                  background=[("selected", DFP_BG)],
                  foreground=[("selected", DFP_ORANGE)])

        # Buttons
        style.configure("Accent.TButton",  background=DFP_ORANGE, foreground=DFP_WHITE,
                        font=("Segoe UI", 10, "bold"), relief="flat", padding=[8, 6])
        style.map("Accent.TButton",
                  background=[("active", "#e06000"), ("pressed", "#cc5500")])
        style.configure("TButton",  background="#e8eaed", foreground="#1a1a1a",
                        font=("Segoe UI", 9), relief="flat", padding=[6, 4])
        style.map("TButton",
                  background=[("active", "#d5d8dc"), ("pressed", "#c0c4c9")])

        # Treeview
        style.configure("Treeview",        background=DFP_WHITE, fieldbackground=DFP_WHITE,
                        rowheight=22, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background="#e8eaed", foreground="#1a1a1a",
                        font=("Segoe UI", 9, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", "#fff0e5")],
                  foreground=[("selected", DFP_DARK)])

        # ── Header bar ──────────────────────────────────────────────────────────
        header = tk.Frame(self.root, bg=DFP_DARK, height=52)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        tk.Label(header, text="Defense Fire Protection",
                 bg=DFP_DARK, fg=DFP_WHITE,
                 font=("Segoe UI", 12, "bold")).pack(side="left", padx=(20, 0), pady=14)
        tk.Label(header, text="  |  Turnover Package Tool",
                 bg=DFP_DARK, fg=DFP_ORANGE,
                 font=("Segoe UI", 10)).pack(side="left", pady=14)
        tk.Label(header, text=f"v{APP_VERSION}",
                 bg=DFP_DARK, fg="#888888",
                 font=("Segoe UI", 8)).pack(side="right", padx=(0, 20), pady=14)

        # ── Main content ─────────────────────────────────────────────────────────
        main = tk.Frame(self.root, bg=DFP_BG, padx=10, pady=10)
        main.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill="both", expand=True)

        self._build_soo_tab()
        self._build_vi_tab()
        self._build_turnover_tab()
        self._build_projects_tab()

    # ─── TAB 1: SOO ─────────────────────────────────────────────────────────

    def _build_soo_tab(self):
        soo_frame = ttk.Frame(self.notebook)
        self.notebook.add(soo_frame, text="  Sequence of Operations  ")

        pane = ttk.PanedWindow(soo_frame, orient="horizontal")
        pane.pack(fill="both", expand=True)

        # LEFT PANEL - scrollable
        left_outer = ttk.Frame(pane, width=340)
        pane.add(left_outer, weight=1)

        left_canvas = tk.Canvas(left_outer, highlightthickness=0, bg="#f4f5f7")
        left_scrollbar = ttk.Scrollbar(left_outer, orient="vertical", command=left_canvas.yview)
        left = ttk.Frame(left_canvas)
        left.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_win_id = left_canvas.create_window((0, 0), window=left, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        left_canvas.bind("<Configure>", lambda e: left_canvas.itemconfigure(left_win_id, width=e.width))
        left_scrollbar.pack(side="right", fill="y")
        left_canvas.pack(side="left", fill="both", expand=True)

        def _on_left_mousewheel(event):
            left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        left_canvas.bind("<Enter>", lambda e: left_canvas.bind_all("<MouseWheel>", _on_left_mousewheel))
        left_canvas.bind("<Leave>", lambda e: left_canvas.unbind_all("<MouseWheel>"))

        # ── Program File / Network Nodes ────────────────────────────────────────
        file_frame = ttk.LabelFrame(left, text="Program File / Network Nodes", padding=8)
        file_frame.pack(fill="x", padx=4, pady=4)

        node_cols = ("node", "name", "devices")
        self.node_list = ttk.Treeview(file_frame, columns=node_cols, show="headings",
                                      height=3, selectmode="browse")
        self.node_list.heading("node",    text="Node")
        self.node_list.heading("name",    text="Name")
        self.node_list.heading("devices", text="Devices")
        self.node_list.column("node",    width=45,  stretch=False)
        self.node_list.column("name",    width=150)
        self.node_list.column("devices", width=60,  stretch=False)
        self.node_list.pack(fill="x", pady=(0, 4))

        node_btn_row = ttk.Frame(file_frame)
        node_btn_row.pack(fill="x")
        ttk.Button(node_btn_row, text="Add Report PDF...",
                   command=self._browse_add_node,
                   style="Accent.TButton").pack(side="left", padx=(0, 4))
        ttk.Button(node_btn_row, text="CSV (backup)...",
                   command=self._browse_file).pack(side="left", padx=(0, 4))
        ttk.Button(node_btn_row, text="Remove",
                   command=self._remove_node).pack(side="right")

        self.file_var = tk.StringVar(value="No nodes loaded")
        ttk.Label(file_frame, textvariable=self.file_var, foreground="#888",
                  font=("Segoe UI", 8), wraplength=260).pack(anchor="w", pady=(4, 0))

        # Custom Control PDF upload (only shown for CSV / single-panel)
        cc_frame = ttk.LabelFrame(left, text="Custom Control Report (Optional)", padding=8)
        cc_frame.pack(fill="x", padx=4, pady=4)
        cc_row = ttk.Frame(cc_frame)
        cc_row.pack(fill="x")
        self.cc_file_var = tk.StringVar(value="No file selected")
        ttk.Label(cc_row, textvariable=self.cc_file_var, foreground="#333",
                  wraplength=200).pack(side="left", fill="x", expand=True)
        ttk.Button(cc_row, text="Browse...", command=self._browse_cc_pdf).pack(side="right")
        self.cc_detect_var = tk.StringVar(value="")
        ttk.Label(cc_frame, textvariable=self.cc_detect_var, foreground="green",
                  font=("Segoe UI", 8)).pack(anchor="w")

        # Job info
        info_frame = ttk.LabelFrame(left, text="Job Information", padding=8)
        info_frame.pack(fill="x", padx=4, pady=4)
        ttk.Label(info_frame, text="Company Name:").pack(anchor="w")
        self.company_var = tk.StringVar(value=self.config.get("company_name", ""))
        ttk.Entry(info_frame, textvariable=self.company_var, width=38).pack(fill="x", pady=(0, 6))
        ttk.Label(info_frame, text="Job / Network Name:").pack(anchor="w")
        self.job_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.job_var, width=38).pack(fill="x", pady=(0, 6))
        ttk.Label(info_frame, text="Panel Model:").pack(anchor="w")
        self.panel_var = tk.StringVar(value="Autocall 4007")
        ttk.Entry(info_frame, textvariable=self.panel_var, width=38).pack(fill="x", pady=(0, 6))

        ttk.Label(info_frame, text="Panel Series (for Operating Instructions):").pack(anchor="w")
        self.panel_series_var = tk.StringVar(value="4007ES")
        series_combo = ttk.Combobox(info_frame, textvariable=self.panel_series_var,
                                     values=PANEL_SERIES_OPTIONS, width=35, state="readonly")
        series_combo.pack(fill="x", pady=(0, 4))

        self.touchscreen_var = tk.BooleanVar(value=False)
        self.touchscreen_chk = ttk.Checkbutton(info_frame, text="Has Touch Screen Display (4010ES/4100ES only)",
                                                variable=self.touchscreen_var,
                                                command=lambda: self._update_oi_preview())
        self.touchscreen_chk.pack(anchor="w", pady=(0, 6))

        def _on_series_change(*_a):
            series = self.panel_series_var.get()
            if series in TOUCHSCREEN_CAPABLE_SERIES:
                self.touchscreen_chk.config(state="normal")
            else:
                self.touchscreen_var.set(False)
                self.touchscreen_chk.config(state="disabled")
            self._update_oi_preview()
        series_combo.bind("<<ComboboxSelected>>", _on_series_change)
        _on_series_change()

        ttk.Label(info_frame, text="Notes:").pack(anchor="w")
        self.notes_text = tk.Text(info_frame, height=3, width=38, font=("Segoe UI", 9))
        self.notes_text.pack(fill="x", pady=(0, 4))

        ttk.Separator(info_frame, orient="horizontal").pack(fill="x", pady=6)
        ttk.Label(info_frame, text="Panel Access Codes:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        ttk.Label(info_frame, text="Levels 3 & 4 are for technicians only — do not share.",
                  foreground="#888", font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 4))
        codes_frame = ttk.Frame(info_frame)
        codes_frame.pack(fill="x", pady=(0, 4))
        ttk.Label(codes_frame, text="Level 2:", width=8).grid(row=0, column=0, sticky="w")
        self.code_l1_var = tk.StringVar()  # unused, kept for compat
        self.code_l2_var = tk.StringVar(value="222222")
        ttk.Entry(codes_frame, textvariable=self.code_l2_var, width=10).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(codes_frame, text="Bypass / disable points", foreground="#666",
                  font=("Segoe UI", 8)).grid(row=0, column=2, sticky="w")

        # Logo
        logo_frame = ttk.LabelFrame(left, text="Company Logo", padding=8)
        logo_frame.pack(fill="x", padx=4, pady=4)
        self.logo_label = ttk.Label(logo_frame, text="No logo selected", foreground="#999")
        self.logo_label.pack(pady=4)
        logo_btns = ttk.Frame(logo_frame)
        logo_btns.pack()
        ttk.Button(logo_btns, text="Upload Logo", command=self._browse_logo).pack(side="left", padx=2)
        ttk.Button(logo_btns, text="Clear", command=self._clear_logo).pack(side="left", padx=2)

        # Sections
        sec_frame = ttk.LabelFrame(left, text="Include Sections", padding=8)
        sec_frame.pack(fill="x", padx=4, pady=4)
        self.section_vars = {}
        for key, label in [
            ("include_point_list", "Device Point List"),
            ("include_initiating", "Initiating Devices (Alarm)"),
            ("include_supervisory", "Supervisory Devices"),
            ("include_multi_criteria", "Multi-Criteria Groups"),
            ("include_isolators", "Isolator / Monitor Modules"),
            ("include_outputs", "Output Devices (Relays)"),
            ("include_general_alarm", "General Alarm Sequence"),
            ("include_trouble", "Trouble Sequence"),
            ("include_supervisory_seq", "Supervisory Sequence"),
            ("include_reset_silence", "Reset & Silence"),
            ("include_location_summary", "Location Summary"),
        ]:
            var = tk.BooleanVar(value=self.config.get(key, True))
            self.section_vars[key] = var
            ttk.Checkbutton(sec_frame, text=label, variable=var).pack(anchor="w")
        select_btns = ttk.Frame(sec_frame)
        select_btns.pack(fill="x", pady=(4, 0))
        ttk.Button(select_btns, text="Select All", command=lambda: [v.set(True) for v in self.section_vars.values()]).pack(side="left", padx=2)
        ttk.Button(select_btns, text="Deselect All", command=lambda: [v.set(False) for v in self.section_vars.values()]).pack(side="left", padx=2)

        # Export buttons
        btn_frame = ttk.Frame(left)
        btn_frame.pack(fill="x", padx=4, pady=(12, 4))
        ttk.Separator(btn_frame).pack(fill="x", pady=(0, 10))
        self.export_btn = ttk.Button(btn_frame, text="⬇  Export SOO PDF", command=self._export_soo_pdf,
                                     style="Accent.TButton", state="disabled")
        self.export_btn.pack(fill="x", ipady=8)
        self.export_html_btn = ttk.Button(btn_frame, text="Export Customer Interactive HTML",
                                          command=self._export_html, state="disabled")
        self.export_html_btn.pack(fill="x", ipady=5, pady=(6, 0))
        self.status_var = tk.StringVar(value="Load a CSV file to begin")
        ttk.Label(btn_frame, textvariable=self.status_var, foreground="#888",
                  font=("Segoe UI", 8), wraplength=300).pack(pady=(6, 0))

        # RIGHT PANEL
        right = ttk.Frame(pane)
        pane.add(right, weight=2)
        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True, padx=4, pady=4)

        table_frame = ttk.Frame(nb)
        nb.add(table_frame, text="  Devices  ")
        toolbar = ttk.Frame(table_frame)
        toolbar.pack(fill="x", pady=(0, 4))
        self.device_count_var = tk.StringVar(value="0 devices")
        ttk.Label(toolbar, textvariable=self.device_count_var, font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Button(toolbar, text="Add Device", command=self._add_soo_device).pack(side="right", padx=2)
        ttk.Button(toolbar, text="Edit Selected", command=self._edit_device).pack(side="right", padx=2)
        ttk.Button(toolbar, text="Delete Selected", command=self._delete_device).pack(side="right", padx=2)

        tree_container = ttk.Frame(table_frame)
        tree_container.pack(fill="both", expand=True)
        tree_cols = ("node", "address", "type", "signal", "location", "description")
        self.tree = ttk.Treeview(tree_container, columns=tree_cols, show="headings", selectmode="extended")
        for col, text, w in [("node","Node",45),("address","Address",80),("type","Type",70),
                              ("signal","Signal",90),("location","Location",160),
                              ("description","Description",180)]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=w, stretch=(col not in ("node","type","signal")))
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        summary_frame = ttk.Frame(nb)
        nb.add(summary_frame, text="  Summary  ")
        self.summary_text = scrolledtext.ScrolledText(summary_frame, font=("Consolas", 9), wrap="word", state="disabled")
        self.summary_text.pack(fill="both", expand=True)

    # ─── TAB 2: VI REPORT ────────────────────────────────────────────────────

    def _build_vi_tab(self):
        vi_frame = ttk.Frame(self.notebook)
        self.notebook.add(vi_frame, text="  Verification Report  ")

        # Split: left = form, right = preview
        vi_pane = ttk.PanedWindow(vi_frame, orient="horizontal")
        vi_pane.pack(fill="both", expand=True)

        # ── LEFT: scrollable form ──
        left_outer = ttk.Frame(vi_pane)
        vi_pane.add(left_outer, weight=3)

        vi_canvas = tk.Canvas(left_outer, highlightthickness=0)
        vi_sb = ttk.Scrollbar(left_outer, orient="vertical", command=vi_canvas.yview)
        scroll_frame = ttk.Frame(vi_canvas)
        scroll_frame.bind("<Configure>", lambda e: vi_canvas.configure(scrollregion=vi_canvas.bbox("all")))
        vi_win_id = vi_canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        vi_canvas.configure(yscrollcommand=vi_sb.set)
        vi_canvas.bind("<Configure>", lambda e: vi_canvas.itemconfigure(vi_win_id, width=e.width))
        vi_sb.pack(side="right", fill="y")
        vi_canvas.pack(side="left", fill="both", expand=True)
        def _on_vi_mw(event):
            vi_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        vi_canvas.bind("<Enter>", lambda e: vi_canvas.bind_all("<MouseWheel>", _on_vi_mw))
        vi_canvas.bind("<Leave>", lambda e: vi_canvas.unbind_all("<MouseWheel>"))

        self.vi_vars = {}
        self.vi_checklist_vars = {}

        def vi_field(parent, key, label, default="", width=40):
            row = ttk.Frame(parent)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=label, width=24, anchor="w").pack(side="left")
            var = tk.StringVar(value=default)
            var.trace_add("write", lambda *a: self._update_vi_preview())
            self.vi_vars[key] = var
            ttk.Entry(row, textvariable=var, width=width).pack(side="left", fill="x", expand=True)

        def vi_combo(parent, key, label, values, default=""):
            row = ttk.Frame(parent)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=label, width=24, anchor="w").pack(side="left")
            var = tk.StringVar(value=default)
            var.trace_add("write", lambda *a: self._update_vi_preview())
            self.vi_vars[key] = var
            ttk.Combobox(row, textvariable=var, values=values, width=20, state="readonly").pack(side="left")

        def vi_checklist_section(parent, section_def):
            """Build a collapsible checklist section with Y/N/NA toggles."""
            sheet_key, title, yes_col, no_col, na_col, items = section_def
            frame = ttk.LabelFrame(parent, text=title, padding=6)
            frame.pack(fill="x", padx=8, pady=3)

            # Header with item count and expand/collapse
            header = ttk.Frame(frame)
            header.pack(fill="x")
            count_var = tk.StringVar(value=f"{len(items)} items - standard defaults applied")
            ttk.Label(header, textvariable=count_var, foreground="#28a745",
                      font=("Segoe UI", 8)).pack(side="left")
            toggle_btn = ttk.Button(header, text="Show Items", width=12)
            toggle_btn.pack(side="right")

            # Items container (hidden by default)
            items_frame = ttk.Frame(frame)
            items_visible = [False]

            def toggle():
                if items_visible[0]:
                    items_frame.pack_forget()
                    toggle_btn.config(text="Show Items")
                    items_visible[0] = False
                else:
                    items_frame.pack(fill="x", pady=(4, 0))
                    toggle_btn.config(text="Hide Items")
                    items_visible[0] = True
            toggle_btn.config(command=toggle)

            has_na = na_col is not None
            # Column headers
            hdr = ttk.Frame(items_frame)
            hdr.pack(fill="x")
            ttk.Label(hdr, text="Item", width=60, anchor="w",
                      font=("Segoe UI", 8, "bold")).pack(side="left")
            for lbl in (["YES", "NO", "N/A"] if has_na else ["YES", "NO"]):
                ttk.Label(hdr, text=lbl, width=5, anchor="center",
                          font=("Segoe UI", 8, "bold")).pack(side="left")

            for _row, item_id, text, default in items:
                item_frame = ttk.Frame(items_frame)
                item_frame.pack(fill="x", pady=0)

                display_text = f"{item_id}. {text}"
                if len(display_text) > 75:
                    display_text = display_text[:72] + "..."
                ttk.Label(item_frame, text=display_text, width=60, anchor="w",
                          font=("Segoe UI", 8)).pack(side="left")

                var_key = f"{sheet_key}_{item_id}"
                var = tk.StringVar(value=default)
                self.vi_checklist_vars[var_key] = var

                vals = ["Y", "N", "A"] if has_na else ["Y", "N"]
                for v in vals:
                    rb = ttk.Radiobutton(item_frame, variable=var, value=v, width=3)
                    rb.pack(side="left")

        # ═══ FORM SECTIONS ═══

        # ── Job Info ──
        f = ttk.LabelFrame(scroll_frame, text="Building & Job Information", padding=8)
        f.pack(fill="x", padx=8, pady=4)
        vi_field(f, "building_name", "Building Name:")
        vi_field(f, "building_address", "Building Address:")
        vi_field(f, "date", "Date:", default=datetime.now().strftime("%Y-%m-%d"))
        vi_field(f, "building_permit", "Building Permit #:")
        vi_field(f, "electrical_permit", "Electrical Permit #:")
        vi_field(f, "model_number", "Panel Model Number:")
        vi_field(f, "owner_rep", "Owner / Representative:")
        vi_field(f, "s524_edition", "CAN/ULC-S524 Edition:", default="2019")
        vi_field(f, "building_code_edition", "Building Code Edition:", default="2023")
        vi_field(f, "cec_edition", "CEC Edition:", default="2024")
        vi_field(f, "doc_location", "Documentation Location:", default="PANEL INSTRUCTIONS IN FACP")

        # ── Technician ──
        f = ttk.LabelFrame(scroll_frame, text="Technician Information", padding=8)
        f.pack(fill="x", padx=8, pady=4)
        vi_field(f, "tech_name", "Primary Technician:", default=self.config.get("technician_name", ""))
        vi_field(f, "tech_cert", "Certification #:", default=self.config.get("technician_cert", ""))
        vi_field(f, "tech_company", "Company:", default="Defense Fire Protection Inc.")
        vi_field(f, "tech_phone", "Phone:", default="403-807-5822")
        vi_field(f, "assist_tech_name", "Assisting Technician:")
        vi_field(f, "assist_company", "Assisting Company:")
        vi_field(f, "assist_phone", "Assisting Phone:")
        vi_field(f, "assist_cert", "Assisting Cert #:")
        vi_field(f, "designer_name", "Designer / Engineer:")
        vi_field(f, "designer_company", "Designer Company:")

        # ── Panel / Control Unit ──
        f = ttk.LabelFrame(scroll_frame, text="Control Unit / Panel", padding=8)
        f.pack(fill="x", padx=8, pady=4)
        vi_field(f, "panel_location", "Panel Location:", default="FRONT DOOR")
        vi_field(f, "panel_id", "Panel Identification:", default="FACP")
        vi_field(f, "dcl_id", "Data Comm Link ID:", default="M1")
        vi_field(f, "firmware_rev", "Firmware Revision:")
        vi_field(f, "firmware_ver", "Firmware Version:")

        # ── Power Supply ──
        f = ttk.LabelFrame(scroll_frame, text="Power Supply & Battery", padding=8)
        f.pack(fill="x", padx=8, pady=4)
        vi_field(f, "disconnect_location", "Disconnect Location:")
        vi_field(f, "breaker_id", "Breaker Identification:")
        vi_field(f, "battery_type", "Battery Type:", default="12V 12AH")
        vi_field(f, "battery_qty", "Battery Qty:", default="2", width=10)
        vi_combo(f, "alarm_time", "Alarm Operation Time:", ["30", "60", "120"], default="30")
        vi_combo(f, "battery_test_type", "Battery Test Type:", ["24h", "silent", "accelerated"], default="24h")
        vi_field(f, "voltage_on", "Voltage (AC on) V:", width=12)
        vi_field(f, "voltage_off", "Voltage (AC off) V:", width=12)
        vi_field(f, "current_off", "Current (standby) A:", width=12)
        vi_field(f, "voltage_alarm", "Voltage (alarm) V:", width=12)
        vi_field(f, "current_alarm", "Current (alarm) A:", width=12)
        vi_field(f, "charging_current", "Charging Current A:", width=12)
        vi_field(f, "battery_date", "Battery Date Code:")
        vi_field(f, "battery_capacity", "Calc Battery Cap Ah:", width=12)

        # Voltage after test with live 85% pass/fail indicator
        vat_row = ttk.Frame(f)
        vat_row.pack(fill="x", pady=1)
        ttk.Label(vat_row, text="Voltage After Test V:", width=24, anchor="w").pack(side="left")
        vat_var = tk.StringVar()
        vat_var.trace_add("write", lambda *a: self._update_vi_preview())
        self.vi_vars["voltage_after_test"] = vat_var
        ttk.Entry(vat_row, textvariable=vat_var, width=10).pack(side="left", padx=(0, 8))
        self.batt_result_var = tk.StringVar(value="")
        ttk.Label(vat_row, textvariable=self.batt_result_var,
                  font=("Segoe UI", 8, "bold")).pack(side="left")

        def _on_vat_change(*_):
            self._update_batt_85_check()
        vat_var.trace_add("write", lambda *a: _on_vat_change())
        self.vi_vars["battery_type"].trace_add("write", lambda *a: _on_vat_change())

        # ── Device Model Numbers ──
        f = ttk.LabelFrame(scroll_frame, text="Device Model Numbers (Sheet 33.1)", padding=8)
        f.pack(fill="x", padx=8, pady=4)
        vi_field(f, "model_pull", "Manual Station:", default="A4099-9004")
        vi_field(f, "model_smoke", "Smoke Detector:", default="HEAD A4098-9714\nBASE A4098-9792")
        vi_field(f, "model_heat", "Heat Detector:", default="A4098-9733")
        vi_field(f, "model_co", "CO Detector:")
        vi_field(f, "model_strobe", "Visual Signal:", default="A4906-9101")
        vi_field(f, "model_hornstrobe", "Horn/Strobe Combo:", default="A4906-9127")
        vi_field(f, "model_eol", "End of Line:", default="A2081-9039C")

        # ── Remarks ──
        f = ttk.LabelFrame(scroll_frame, text="Remarks / Comments", padding=8)
        f.pack(fill="x", padx=8, pady=4)
        self.vi_remarks = tk.Text(f, height=4, width=50, font=("Segoe UI", 9))
        self.vi_remarks.pack(fill="x")
        vi_field(f, "comments", "C28 Comments:")

        # ── Manual Devices (Non-Addressable) ──
        # ── Additional Devices (Non-Addressable) ──
        f_manual = ttk.LabelFrame(scroll_frame, text="Additional Devices (Non-Addressable)", padding=8)
        f_manual.pack(fill="x", padx=8, pady=4)
        ttk.Label(f_manual, text="Horn/strobes, speakers, bells, etc. not in CSV export. These go on Sheet 33.2 after addressable devices.",
                  font=("Segoe UI", 8), foreground="#666").pack(anchor="w")

        self.vi_manual_devices = []

        md_cols = ("floor", "location", "type", "circuit", "installed", "operation", "supervision")
        self.vi_manual_tree = ttk.Treeview(f_manual, columns=md_cols, show="headings", height=6, selectmode="browse")
        for c, text, w in [("floor", "Floor", 100), ("location", "Location", 140), ("type", "Type", 60),
                            ("circuit", "Circuit", 65), ("installed", "Installed", 60),
                            ("operation", "Operation", 65), ("supervision", "Supervision", 70)]:
            self.vi_manual_tree.heading(c, text=text)
            self.vi_manual_tree.column(c, width=w)
        self.vi_manual_tree.pack(fill="x", pady=(4, 4))

        md_btn_frame = ttk.Frame(f_manual)
        md_btn_frame.pack(fill="x")
        ttk.Button(md_btn_frame, text="Add Device", command=self._add_manual_device).pack(side="left", padx=(0, 4))
        ttk.Button(md_btn_frame, text="Remove Selected", command=self._remove_manual_device).pack(side="left")

        # ── 32.13 Ancillary Device Circuit Test ──
        f_anc = ttk.LabelFrame(scroll_frame, text="Ancillary Device Circuit Test (Sheet 32.13)", padding=8)
        f_anc.pack(fill="x", padx=8, pady=4)
        ttk.Label(f_anc, text="Relay outputs and controlled equipment. Auto-filled from CSV relays, add more below.",
                  font=("Segoe UI", 8), foreground="#666").pack(anchor="w")

        self.vi_ancillary_devices = []
        anc_cols = ("device", "powered_by", "confirmed", "method")
        self.vi_ancillary_tree = ttk.Treeview(f_anc, columns=anc_cols, show="headings", height=4, selectmode="browse")
        for c, text, w in [("device", "Circuit / Device", 180), ("powered_by", "Powered By", 120),
                            ("confirmed", "Confirmed", 70), ("method", "Method", 100)]:
            self.vi_ancillary_tree.heading(c, text=text)
            self.vi_ancillary_tree.column(c, width=w)
        self.vi_ancillary_tree.pack(fill="x", pady=(4, 4))

        anc_btn = ttk.Frame(f_anc)
        anc_btn.pack(fill="x")
        ttk.Button(anc_btn, text="Add Ancillary", command=self._add_ancillary).pack(side="left", padx=(0, 4))
        ttk.Button(anc_btn, text="Remove Selected", command=self._remove_ancillary).pack(side="left")

        # ── 32.4 Non-Life Safety Message Circuit Supervision ──
        f = ttk.LabelFrame(scroll_frame, text="Non-Life Safety Message Circuit Test (Sheet 32.4)", padding=8)
        f.pack(fill="x", padx=8, pady=4)
        vi_field(f, "speakers_installed", "Speakers Installed:", default="NO SPEAKERS")

        # ── Fire Signal Receiving Centre ──
        f = ttk.LabelFrame(scroll_frame, text="Fire Signal Receiving Centre", padding=8)
        f.pack(fill="x", padx=8, pady=4)
        self.vi_fsrc_detected_var = tk.StringVar(value="No city/monitoring relays detected yet - load a CSV.")
        ttk.Label(f, textvariable=self.vi_fsrc_detected_var, font=("Segoe UI", 8),
                  foreground="#28a745", wraplength=380).pack(anchor="w", pady=(0, 4))
        vi_combo(f, "fsrc_connected", "Connected to FSRC:", ["Yes", "No"], default="No")
        vi_field(f, "fsrc_name", "FSRC Name:")
        vi_field(f, "fsrc_phone", "FSRC Phone:")

        # ── 33.3 Subsequent Alarm Activation Test ──
        f_333 = ttk.LabelFrame(scroll_frame, text="Subsequent Alarm Activation Test (Sheet 33.3)", padding=8)
        f_333.pack(fill="x", padx=8, pady=4)
        ttk.Label(f_333, text="Test that signals re-silence after subsequent alarm in different zone.",
                  font=("Segoe UI", 8), foreground="#666").pack(anchor="w")

        self.vi_333_entries = []
        for i in range(3):
            row = ttk.Frame(f_333)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=f"Test {i+1}:", width=6, font=("Segoe UI", 8, "bold")).pack(side="left")
            vars_row = {}
            for lbl, key, w in [("Init Zone", "init_zone", 12), ("Init Device", "init_dev", 14),
                                 ("Subseq Zone", "sub_zone", 12), ("Subseq Device", "sub_dev", 14)]:
                ttk.Label(row, text=lbl, font=("Segoe UI", 7)).pack(side="left", padx=(3, 0))
                var = tk.StringVar()
                ttk.Entry(row, textvariable=var, width=w).pack(side="left", padx=(1, 0))
                vars_row[key] = var
            var_silent = tk.StringVar(value="No")
            ttk.Label(row, text="Silent?", font=("Segoe UI", 7)).pack(side="left", padx=(3, 0))
            ttk.Combobox(row, textvariable=var_silent, values=["Yes", "No"], width=4, state="readonly").pack(side="left", padx=(1, 0))
            vars_row["silent"] = var_silent
            self.vi_333_entries.append(vars_row)

        # ── 33.4 Circuit Fault Tolerance Test ──
        f_334 = ttk.LabelFrame(scroll_frame, text="Circuit Fault Tolerance Test (Sheet 33.4)", padding=8)
        f_334.pack(fill="x", padx=8, pady=4)
        ttk.Label(f_334, text="Auto-populated from device locations. Edit or add rows below.",
                  font=("Segoe UI", 8), foreground="#666").pack(anchor="w")

        self.vi_334_entries = []
        hdr = ttk.Frame(f_334)
        hdr.pack(fill="x")
        for lbl, w in [("Fault Location", 14), ("S", 2), ("O", 2), ("G", 2),
                        ("Isolation Result", 14), ("Non-Fault Device", 14), ("P/F", 3)]:
            ttk.Label(hdr, text=lbl, width=w, font=("Segoe UI", 7, "bold"), anchor="center").pack(side="left", padx=1)
        for i in range(6):
            row = ttk.Frame(f_334)
            row.pack(fill="x", pady=1)
            vars_row = {}
            for key, w in [("fault_loc", 14), ("short", 2), ("open", 2), ("ground", 2),
                            ("isolation", 14), ("nonfault_dev", 14), ("result", 3)]:
                var = tk.StringVar()
                if key in ("short", "open", "ground"):
                    var.set("X")
                    ttk.Entry(row, textvariable=var, width=w, justify="center").pack(side="left", padx=1)
                elif key == "result":
                    var.set("P")
                    ttk.Combobox(row, textvariable=var, values=["P", "F"], width=w, state="readonly").pack(side="left", padx=1)
                else:
                    ttk.Entry(row, textvariable=var, width=w).pack(side="left", padx=1)
                vars_row[key] = var
            self.vi_334_entries.append(vars_row)

        # ── Sound Level Measurements (33.5) ──
        f_snd = ttk.LabelFrame(scroll_frame, text="Sound Level Measurements (Sheet 33.5)", padding=8)
        f_snd.pack(fill="x", padx=8, pady=4)
        self.vi_sound_entries = []
        for i in range(4):
            row = ttk.Frame(f_snd)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=f"Row {i+1}:", width=6, anchor="w").pack(side="left")
            vars_row = {}
            for lbl, key, w in [("Zone", "zone", 12), ("Location", "location", 16),
                                 ("Ambient dBA", "ambient", 8), ("Alarm dBA", "alarm", 8),
                                 ("Remarks", "remarks", 14)]:
                ttk.Label(row, text=lbl, font=("Segoe UI", 8)).pack(side="left", padx=(4, 0))
                var = tk.StringVar()
                ttk.Entry(row, textvariable=var, width=w).pack(side="left", padx=(2, 0))
                vars_row[key] = var
            self.vi_sound_entries.append(vars_row)

        # ── Battery Calculation (DATA sheet) ──
        # ── Checklist Sections ──
        ttk.Label(scroll_frame, text="S537 Checklist Sections",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=12, pady=(8, 2))
        ttk.Label(scroll_frame, text="Standard defaults applied for Autocall single-stage. Expand any section to modify.",
                  font=("Segoe UI", 8), foreground="#666").pack(anchor="w", padx=12, pady=(0, 4))

        for sec in VI_SECTIONS:
            vi_checklist_section(scroll_frame, sec)
        vi_checklist_section(scroll_frame, VI_SECTIONS_32_14)

        # ── Export ──
        btn_frame = ttk.Frame(scroll_frame)
        btn_frame.pack(fill="x", padx=8, pady=12)
        self.export_vi_btn = ttk.Button(btn_frame, text="Generate VI Report (.xlsm)",
                                        command=self._export_vi, style="Accent.TButton", state="disabled")
        self.export_vi_btn.pack(fill="x", ipady=6)
        self.export_vi_pdf_btn = ttk.Button(btn_frame, text="Export as PDF",
                                             command=self._export_vi_pdf, state="disabled")
        self.export_vi_pdf_btn.pack(fill="x", ipady=4, pady=(4, 0))
        self.vi_status_var = tk.StringVar(value="Load a CSV on the SOO tab first")
        ttk.Label(btn_frame, textvariable=self.vi_status_var, foreground="#666").pack(pady=(4, 0))

        # ── RIGHT: live preview ──
        right_frame = ttk.Frame(vi_pane)
        vi_pane.add(right_frame, weight=2)

        ttk.Label(right_frame, text="VI Report Preview",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(8, 4))

        self.vi_preview = scrolledtext.ScrolledText(right_frame, font=("Consolas", 9),
                                                     wrap="word", state="disabled",
                                                     background="#1e1e2e", foreground="#cdd6f4")
        self.vi_preview.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self._update_vi_preview()

    def _update_vi_preview(self):
        """Update the live VI preview panel."""
        if not hasattr(self, 'vi_preview'):
            return
        self.vi_preview.config(state="normal")
        self.vi_preview.delete("1.0", "end")

        def val(key, fallback="___"):
            if key in self.vi_vars:
                v = self.vi_vars[key].get()
                return v if v else fallback
            return fallback

        dev_count = len(self.devices)
        vi_dev_count = len([d for d in self.devices if d.get("device_type") not in ("MCOP", "MAOFF")])
        relay_count = len([d for d in self.devices if d.get("cat") == "RELAY"])

        lines = []
        lines.append("=" * 52)
        lines.append("  CAN/ULC-S537 VERIFICATION REPORT")
        lines.append("  Defense Fire Protection Inc.")
        lines.append("=" * 52)
        lines.append("")
        lines.append(f"  Building:    {val('building_name')}")
        lines.append(f"  Address:     {val('building_address')}")
        lines.append(f"  Date:        {val('date')}")
        lines.append(f"  Panel:       AUTOCALL {val('model_number')}")
        lines.append(f"  Bldg Permit: {val('building_permit')}")
        lines.append(f"  Elec Permit: {val('electrical_permit')}")
        lines.append(f"  Owner/Rep:   {val('owner_rep')}")
        lines.append("")
        lines.append("─" * 52)
        lines.append("  TECHNICIAN")
        lines.append(f"  Primary:     {val('tech_name')}")
        lines.append(f"  Cert #:      {val('tech_cert')}")
        lines.append(f"  Assisting:   {val('assist_tech_name')}")
        lines.append(f"  Designer:    {val('designer_name')}")
        lines.append("")
        lines.append("─" * 52)
        lines.append("  CONTROL UNIT")
        lines.append(f"  Location:    {val('panel_location')}")
        lines.append(f"  ID:          {val('panel_id')}")
        lines.append(f"  Firmware:    Rev {val('firmware_rev')} / Ver {val('firmware_ver')}")
        lines.append(f"  DCL:         {val('dcl_id')}")
        lines.append("")
        lines.append("─" * 52)
        lines.append("  BATTERY")
        lines.append(f"  Type:        {val('battery_type')} x{val('battery_qty')}")
        lines.append(f"  V (AC on):   {val('voltage_on')} Vdc")
        lines.append(f"  V (AC off):  {val('voltage_off')} Vdc")
        lines.append(f"  I (standby): {val('current_off')} A")
        lines.append(f"  V (alarm):   {val('voltage_alarm')} Vdc")
        lines.append(f"  I (alarm):   {val('current_alarm')} A")
        lines.append(f"  Charging:    {val('charging_current')} A")
        lines.append(f"  After test:  {val('voltage_after_test')} Vdc")
        lines.append(f"  Date code:   {val('battery_date')}")
        lines.append(f"  Capacity:    {val('battery_capacity')} Ah")
        lines.append("")
        lines.append("─" * 52)
        lines.append("  DEVICES (auto-filled from CSV)")
        if dev_count > 0:
            lines.append(f"  Total parsed:   {dev_count}")
            lines.append(f"  VI device rows: {vi_dev_count}")
            lines.append(f"  Ancillary:      {relay_count}")
            by_cat = defaultdict(int)
            for d in self.devices:
                by_cat[d.get("cat", "?")] += 1
            for cat in ["SMOKE","HEAT","MULTI","CO","PULL","ISO","MON","RELAY","NAC","WFLOW","TAMPER"]:
                if cat in by_cat:
                    lines.append(f"    {CATEGORY_LABELS.get(cat, cat):<26} {by_cat[cat]:>3}")
        else:
            lines.append("  (no CSV loaded)")
        if hasattr(self, 'vi_manual_devices') and self.vi_manual_devices:
            lines.append(f"  Manual (non-addr): {len(self.vi_manual_devices)}")
        lines.append("")
        lines.append("─" * 52)
        lines.append("  CHECKLIST SECTIONS")
        total_items = 0
        for sec in VI_SECTIONS:
            total_items += len(sec[5])
        total_items += len(VI_SECTIONS_32_14[5])
        overrides = sum(1 for k, v in getattr(self, 'vi_checklist_vars', {}).items()
                       if v.get() != self._get_checklist_default(k))
        lines.append(f"  Total items:    {total_items}")
        lines.append(f"  Using defaults: {total_items - overrides}")
        lines.append(f"  Modified:       {overrides}")
        lines.append("")
        lines.append("=" * 52)

        self.vi_preview.insert("1.0", "\n".join(lines))
        self.vi_preview.config(state="disabled")

    def _get_checklist_default(self, key):
        """Look up the default value for a checklist item key."""
        for sec in VI_SECTIONS:
            sheet_key = sec[0]
            for _row, item_id, _text, default in sec[5]:
                if f"{sheet_key}_{item_id}" == key:
                    return default
        for _row, item_id, _text, default in VI_SECTIONS_32_14[5]:
            if f"32.14_{item_id}" == key:
                return default
        return ""

    # ─── TAB 3: TURNOVER PACKAGE ─────────────────────────────────────────────

    def _build_turnover_tab(self):
        tp_frame = ttk.Frame(self.notebook)
        self.notebook.add(tp_frame, text="  Turnover Package  ")

        canvas = tk.Canvas(tp_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(tp_frame, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        def _on_tp_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_tp_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        info = ttk.LabelFrame(scroll_frame, text="Build Interactive HTML Turnover Package", padding=12)
        info.pack(fill="x", padx=12, pady=8)
        ttk.Label(info, text=(
            "Generates a self-contained HTML file with tabbed navigation.\n"
            "Each tab shows a document: SOO (interactive), or PDF documents\n"
            "embedded in a viewer. Give each document a custom tab name.\n"
            "The customer opens it in any browser - no software needed."
        ), wraplength=600, foreground="#555").pack(anchor="w")

        # SOO tab
        soo_frame = ttk.LabelFrame(scroll_frame, text="Sequence of Operations", padding=12)
        soo_frame.pack(fill="x", padx=12, pady=4)
        self.tp_include_soo = tk.BooleanVar(value=True)
        row = ttk.Frame(soo_frame)
        row.pack(fill="x")
        ttk.Checkbutton(row, text="Include SOO", variable=self.tp_include_soo).pack(side="left")
        ttk.Label(row, text="Tab Name:", foreground="#555").pack(side="left", padx=(20, 4))
        self.tp_soo_name = tk.StringVar(value="Sequence of Operations")
        ttk.Entry(row, textvariable=self.tp_soo_name, width=30).pack(side="left")

        # Operating Instructions
        oi_frame = ttk.LabelFrame(scroll_frame, text="Panel Operating Instructions", padding=12)
        oi_frame.pack(fill="x", padx=12, pady=4)
        self.tp_include_oi = tk.BooleanVar(value=True)
        ttk.Checkbutton(oi_frame, text="Include Operating Instructions for this panel",
                        variable=self.tp_include_oi,
                        command=self._update_oi_preview).pack(anchor="w")
        self.tp_oi_preview_var = tk.StringVar(value="")
        ttk.Label(oi_frame, textvariable=self.tp_oi_preview_var, foreground="#555",
                  font=("Segoe UI", 8), wraplength=600).pack(anchor="w", pady=(4, 0))

        # PDF Documents
        attach_frame = ttk.LabelFrame(scroll_frame, text="PDF Documents (VI Report, As-Builts, Engineer Docs, etc.)", padding=12)
        attach_frame.pack(fill="both", expand=True, padx=12, pady=4)

        attach_toolbar = ttk.Frame(attach_frame)
        attach_toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(attach_toolbar, text="Add PDF...", command=self._add_attachment).pack(side="left", padx=2)
        ttk.Button(attach_toolbar, text="Remove Selected", command=self._remove_attachment).pack(side="left", padx=2)
        ttk.Button(attach_toolbar, text="Move Up", command=lambda: self._move_attachment(-1)).pack(side="left", padx=2)
        ttk.Button(attach_toolbar, text="Move Down", command=lambda: self._move_attachment(1)).pack(side="left", padx=2)
        ttk.Button(attach_toolbar, text="Rename Tab", command=self._rename_attachment).pack(side="left", padx=2)

        # Treeview for attachments with custom names
        attach_cols = ("name", "file")
        self.attach_tree = ttk.Treeview(attach_frame, columns=attach_cols, show="headings", height=8, selectmode="browse")
        self.attach_tree.heading("name", text="Tab Name")
        self.attach_tree.heading("file", text="File")
        self.attach_tree.column("name", width=200)
        self.attach_tree.column("file", width=350)
        self.attach_tree.pack(fill="both", expand=True)

        # Export
        export_frame = ttk.Frame(scroll_frame)
        export_frame.pack(fill="x", padx=12, pady=12)
        self.tp_export_btn = ttk.Button(export_frame, text="Build Turnover Package HTML",
                                        command=self._export_turnover, style="Accent.TButton", state="disabled")
        self.tp_export_btn.pack(fill="x", ipady=8)

        # Also offer merged PDF
        self.tp_export_pdf_btn = ttk.Button(export_frame, text="Build Turnover Package PDF (Merged)",
                                            command=self._export_turnover_pdf, state="disabled")
        self.tp_export_pdf_btn.pack(fill="x", ipady=4, pady=(4, 0))

        self.tp_status_var = tk.StringVar(value="Load a CSV file on the SOO tab first")
        ttk.Label(export_frame, textvariable=self.tp_status_var, foreground="#666").pack(pady=(4, 0))

        self._update_oi_preview()

    # ─── FILE HANDLING ───────────────────────────────────────────────────────

    def _browse_file(self):
        initial = self.config.get("last_directory", "")
        if not initial or not os.path.isdir(initial):
            initial = os.path.expanduser("~")
        path = filedialog.askopenfilename(
            title="Select Autocall Program CSV", initialdir=initial,
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if path:
            self._load_file(path)

    def _browse_cc_pdf(self):
        initial = os.path.dirname(self.csv_path) if self.csv_path else self.config.get("last_directory", "")
        path = filedialog.askopenfilename(
            title="Select Custom Control Report PDF", initialdir=initial,
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if path:
            self._load_cc_pdf(path)

    def _load_cc_pdf(self, path):
        try:
            self.cc_data = parse_custom_controls(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse custom control PDF:\n{e}")
            return
        self.cc_pdf_path = path
        fname = os.path.basename(path)
        self.cc_file_var.set(fname)
        stype = self.cc_data.get("system_type", "single_stage")
        eq_count = len(self.cc_data.get("equations", []))
        feats = self.cc_data.get("features", [])
        type_labels = {"single_stage": "Single Stage", "two_stage": "2-Stage",
                       "auto_silence": "Auto-Silence"}
        feat_labels = {"ground_fault": "Ground Fault Logic", "bypass_logic": "Bypass Logic",
                       "auto_silence": "Auto-Silence", "two_stage": "2-Stage"}
        feat_str = ", ".join(feat_labels[f] for f in feats if f in feat_labels) or "Standard"
        self.cc_detect_var.set(f"Detected: {type_labels.get(stype, stype)} | {eq_count} equations | {feat_str}")

    def _browse_report_pdf(self):
        """Legacy single-panel browse — replaces all nodes."""
        initial = os.path.dirname(self.csv_path) if self.csv_path else self.config.get("last_directory", "")
        if not initial or not os.path.isdir(initial):
            initial = os.path.expanduser("~")
        path = filedialog.askopenfilename(
            title="Select 4100U Panel Programmer Report PDF", initialdir=initial,
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if path:
            self._load_report_pdf(path, append=False)

    def _browse_add_node(self):
        """Add a node PDF to the current job (append mode)."""
        initial = os.path.dirname(self.csv_path) if self.csv_path else self.config.get("last_directory", "")
        if not initial or not os.path.isdir(initial):
            initial = os.path.expanduser("~")
        path = filedialog.askopenfilename(
            title="Select Node Report PDF", initialdir=initial,
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if path:
            self._load_report_pdf(path, append=bool(self.nodes))

    def _remove_node(self):
        sel = self.node_list.selection()
        if not sel:
            return
        vals = self.node_list.item(sel[0], "values")
        node_num_str = str(vals[0])
        self.nodes = [n for n in self.nodes
                      if str(n.get("node_num") or n.get("job_name", "?")) != node_num_str]
        self.node_list.delete(sel[0])
        self._rebuild_devices_from_nodes()
        self._populate_tree()
        self._update_summary()

    def _rebuild_devices_from_nodes(self):
        """Flatten all node data into self.devices / nac_circuits / etc."""
        self.devices = []
        self.nac_circuits = []
        self.city_relays = []
        self.bypass_switches = []
        for nd in self.nodes:
            self.devices.extend(nd["devices"])
            self.nac_circuits.extend(nd["nac_circuits"])
            self.city_relays.extend(nd["city_relays"])
            self.bypass_switches.extend(nd["bypass_switches"])

    def _refresh_node_list(self):
        if not hasattr(self, "node_list"):
            return
        self.node_list.delete(*self.node_list.get_children())
        for nd in self.nodes:
            num  = nd.get("node_num") or "—"
            name = nd.get("description") or nd.get("job_name") or "Unknown"
            cnt  = len(nd["devices"])
            self.node_list.insert("", "end", values=(num, name, f"{cnt} dev"))
        total = len(self.devices)
        nc    = len(self.nodes)
        if nc == 0:
            self.file_var.set("No nodes loaded")
        elif nc == 1:
            self.file_var.set(f"1 node — {total} devices")
        else:
            self.file_var.set(f"{nc} nodes — {total} total devices")
        # Show/hide Node column in device tree
        if hasattr(self, "tree"):
            w = 45 if nc > 1 else 0
            self.tree.column("node", width=w, minwidth=w)

    def _load_report_pdf(self, path, append=False):
        try:
            result = parse_report_pdf(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse Report PDF:\n{e}")
            return

        # Tag each device with its node number
        node_num = result.get("node_num")
        for dev in result["devices"]:
            dev["node"] = node_num

        self.csv_path = path
        self.config["last_directory"] = os.path.dirname(path)
        save_config(self.config)

        if not append:
            # Replace mode: clear existing nodes, auto-fill job info from this PDF
            self.nodes = []
            if result["description"]:
                self.job_var.set(result["description"])
            elif result["job_name"]:
                self.job_var.set(result["job_name"])
            if result["panel_type"]:
                self.panel_var.set(result["panel_type"])
                self.panel_model = result["panel_type"]
                detected_series = detect_panel_series(result["panel_type"])
                self.panel_series_var.set(detected_series)
                if detected_series not in TOUCHSCREEN_CAPABLE_SERIES:
                    self.touchscreen_var.set(False)
                    self.touchscreen_chk.config(state="disabled")
                else:
                    self.touchscreen_chk.config(state="normal")
        else:
            # Append mode: only set panel model / series if not already set
            if not self.panel_var.get() and result["panel_type"]:
                self.panel_var.set(result["panel_type"])
                self.panel_model = result["panel_type"]

        # Replace duplicate node (same node_num) or append
        if node_num is not None:
            self.nodes = [n for n in self.nodes if n.get("node_num") != node_num]
        self.nodes.append(result)

        # Rebuild flat device / circuit lists from all nodes
        self._rebuild_devices_from_nodes()

        # CC data: use first node's CC for single-node; for multi-node summarise across all
        all_cc = [n["cc_data"] for n in self.nodes if n.get("cc_data")]
        if all_cc:
            # Merge: most severe system_type wins; union features
            type_priority = {"two_stage": 2, "auto_silence": 1, "single_stage": 0}
            best_cc = max(all_cc, key=lambda c: type_priority.get(c.get("system_type", ""), 0))
            merged_features = list({f for c in all_cc for f in c.get("features", [])})
            self.cc_data = {**best_cc, "features": merged_features}
            self.cc_pdf_path = path
            stype = self.cc_data.get("system_type", "single_stage")
            eq_total = sum(len(c.get("equations", [])) for c in all_cc)
            type_labels = {"single_stage": "Single Stage", "two_stage": "2-Stage",
                           "auto_silence": "Auto-Silence"}
            feat_labels = {"ground_fault": "Ground Fault", "bypass_logic": "Bypass",
                           "auto_silence": "Auto-Silence", "two_stage": "2-Stage"}
            feat_str = ", ".join(feat_labels[f] for f in merged_features if f in feat_labels) or "Standard"
            nc = len(self.nodes)
            self.cc_file_var.set(f"From Report PDF{'s' if nc>1 else ''} ({eq_total} equations, {nc} node{'s' if nc>1 else ''})")
            self.cc_detect_var.set(f"Detected: {type_labels.get(stype, stype)} | {feat_str}")

        self._refresh_node_list()
        self._populate_tree()
        self._update_summary()
        self._apply_fsrc_autodetect()
        self._update_oi_preview()

        for btn in [self.export_btn, self.export_html_btn, self.export_vi_btn,
                    self.export_vi_pdf_btn, self.tp_export_btn, self.tp_export_pdf_btn]:
            btn.config(state="normal")

        nc = len(self.nodes)
        dev_count = len(self.devices)
        node_label = f"Node {node_num}" if node_num else os.path.basename(path)
        action = "Added" if append else "Loaded"
        self.status_var.set(
            f"{action} {node_label} — {len(result['devices'])} devices  |  "
            f"{nc} node(s) total, {dev_count} devices")
        self.vi_status_var.set(f"Ready — {dev_count} devices across {nc} node(s). Fill in fields and generate.")
        self.tp_status_var.set("Ready — add documents and build package")

    def _load_file(self, path):
        try:
            self.devices, self.panel_model, extras = parse_csv(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse file:\n{e}")
            return

        self.nac_circuits = extras["nac_circuits"]
        self.city_relays = extras["city_relays"]
        self.bypass_switches = extras["bypass_switches"]

        self.csv_path = path
        self.nodes = []  # CSV mode — clear any previously loaded PDF nodes
        self.config["last_directory"] = os.path.dirname(path)
        save_config(self.config)

        job = os.path.splitext(os.path.basename(path))[0]
        self.job_var.set(job)
        self.panel_var.set(self.panel_model)
        self._refresh_node_list()
        self.file_var.set(f"CSV: {os.path.basename(path)}")

        detected_series = detect_panel_series(self.panel_model)
        self.panel_series_var.set(detected_series)
        if detected_series not in TOUCHSCREEN_CAPABLE_SERIES:
            self.touchscreen_var.set(False)
            self.touchscreen_chk.config(state="disabled")
        else:
            self.touchscreen_chk.config(state="normal")
        self._update_oi_preview()

        self._populate_tree()
        self._update_summary()
        self._apply_fsrc_autodetect()

        for btn in [self.export_btn, self.export_html_btn, self.export_vi_btn, self.export_vi_pdf_btn, self.tp_export_btn, self.tp_export_pdf_btn]:
            btn.config(state="normal")

        self.status_var.set(f"Loaded {len(self.devices)} devices from {os.path.basename(path)}")
        self.vi_status_var.set(f"Ready - {len(self.devices)} devices loaded. Fill in fields and generate.")
        self.tp_status_var.set("Ready - add documents and build package")

    def _populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        multi = len(self.nodes) > 1
        for dev in self.devices:
            node_val = dev.get("node") or ("" if not multi else "?")
            self.tree.insert("", "end", values=(
                node_val, dev["address"], dev["device_type"],
                dev["signal_type"], dev["location"], dev["desc"]))
        self.device_count_var.set(f"{len(self.devices)} devices")

    def _update_summary(self):
        self.summary_text.config(state="normal")
        self.summary_text.delete("1.0", "end")
        by_cat = defaultdict(list)
        by_loc = defaultdict(list)
        for d in self.devices:
            by_cat[d["cat"]].append(d)
            by_loc[d["location"]].append(d)

        text = f"JOB: {self.job_var.get()}\nPANEL: {self.panel_var.get()}\n"
        text += f"TOTAL DEVICES: {len(self.devices)}\nTOTAL LOCATIONS: {len(by_loc)}\n"
        text += "\n--- DEVICE SUMMARY ---\n"
        for cat in ["SMOKE","HEAT","MULTI","CO","PULL","ISO","MON","RELAY","NAC","WFLOW","TAMPER","OTHER"]:
            if cat in by_cat:
                text += f"  {CATEGORY_LABELS.get(cat, cat):<30} {len(by_cat[cat]):>4}\n"
        mc_groups = get_multi_criteria_groups(self.devices)
        if mc_groups:
            text += f"\n--- MULTI-CRITERIA GROUPS ({len(mc_groups)}) ---\n"
            for base, group in mc_groups.items():
                text += f"  {base}: {group[0]['location']} ({len(group)} sensors)\n"
        text += f"\n--- LOCATIONS ({len(by_loc)}) ---\n"
        for loc in sorted(by_loc.keys()):
            text += f"  {loc}: {len(by_loc[loc])} device(s)\n"

        self.summary_text.insert("1.0", text)
        self.summary_text.config(state="disabled")

    def _edit_device(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Edit", "Select a device to edit.")
            return
        item = sel[0]
        vals = self.tree.item(item, "values")
        idx = self.tree.index(item)

        win = tk.Toplevel(self.root)
        win.title("Edit Device")
        win.geometry("400x280")
        win.transient(self.root)
        win.grab_set()

        fields = [("Address", vals[0]), ("Type", vals[1]), ("Signal", vals[2]),
                  ("Location", vals[3]), ("Description", vals[4])]
        entries = {}
        for i, (label, val) in enumerate(fields):
            ttk.Label(win, text=label + ":").grid(row=i, column=0, padx=10, pady=6, sticky="e")
            ent = ttk.Entry(win, width=40)
            ent.insert(0, val)
            ent.grid(row=i, column=1, padx=10, pady=6)
            entries[label] = ent

        def save():
            nv = tuple(entries[f].get() for f in ["Address","Type","Signal","Location","Description"])
            self.tree.item(item, values=nv)
            dev = self.devices[idx]
            dev["address"], dev["device_type"], dev["signal_type"], dev["location"], dev["desc"] = nv
            self._update_summary()
            win.destroy()

        ttk.Button(win, text="Save", command=save).grid(row=len(fields), column=1, pady=12, sticky="e")

    def _delete_device(self):
        sel = self.tree.selection()
        if not sel:
            return
        count = len(sel)
        msg = f"Delete {count} device(s)?" if count > 1 else f"Delete device {self.tree.item(sel[0], 'values')[0]}?"
        if messagebox.askyesno("Delete", msg):
            # Collect indices in reverse order so deletions don't shift remaining indices
            indices = sorted([self.tree.index(item) for item in sel], reverse=True)
            for item in sel:
                self.tree.delete(item)
            for idx in indices:
                if idx < len(self.devices):
                    del self.devices[idx]
            self.device_count_var.set(f"{len(self.devices)} devices")
            self._update_summary()

    def _add_soo_device(self):
        """Dialog to manually add a conventional (non-addressable) device to the SOO device list."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Add Device")
        dlg.geometry("420x280")
        dlg.grab_set()

        f = ttk.Frame(dlg, padding=16)
        f.pack(fill="both", expand=True)

        fields = {}
        def row(label, key, default="", width=30):
            ttk.Label(f, text=label, width=20, anchor="w").grid(
                row=len(fields), column=0, sticky="w", pady=2)
            var = tk.StringVar(value=default)
            fields[key] = var
            ttk.Entry(f, textvariable=var, width=width).grid(
                row=len(fields)-1, column=1, sticky="ew", padx=(4,0))

        def combo_row(label, key, values, default=""):
            ttk.Label(f, text=label, width=20, anchor="w").grid(
                row=len(fields), column=0, sticky="w", pady=2)
            var = tk.StringVar(value=default)
            fields[key] = var
            ttk.Combobox(f, textvariable=var, values=values, width=28,
                         state="readonly").grid(row=len(fields)-1, column=1, sticky="ew", padx=(4,0))

        ADD_DEVICE_OPTIONS = [
            ("Horn Strobe",      "AV"),
            ("Horn",             "AO"),
            ("Strobe",           "VO"),
            ("Smoke Detector",   "PHOTO"),
            ("Heat Detector",    "HEAT"),
            ("Pull Station",     "ADRPUL"),
            ("EOL Resistor",     "EOL"),
        ]
        option_labels = [o[0] for o in ADD_DEVICE_OPTIONS]
        label_to_dtype = {o[0]: o[1] for o in ADD_DEVICE_OPTIONS}

        row("Address / ID:", "address", "CON-")
        combo_row("Device Type:", "device_type", option_labels, "Horn Strobe")
        f.columnconfigure(1, weight=1)

        # Location row — kept as a named widget so we can focus/select it on "Add Another"
        loc_entry_ref = {}
        loc_row_idx = len(fields)
        ttk.Label(f, text="Location:", width=20, anchor="w").grid(
            row=loc_row_idx, column=0, sticky="w", pady=2)
        loc_var = tk.StringVar()
        fields["location"] = loc_var
        loc_entry = ttk.Entry(f, textvariable=loc_var, width=30)
        loc_entry.grid(row=loc_row_idx, column=1, sticky="ew", padx=(4, 0))
        loc_entry_ref["widget"] = loc_entry

        def _save(keep_open=False):
            label = fields["device_type"].get()
            dtype = label_to_dtype.get(label, "PHOTO")
            if not fields["address"].get():
                messagebox.showwarning("Add Device", "Address is required.", parent=dlg)
                return
            info = DEVICE_TYPES.get(dtype, {"desc": label, "cat": "UNKNOWN", "fn": "ALARM", "vi_code": ""})
            dev = {
                "address": fields["address"].get(),
                "device_type": dtype,
                "signal_type": info.get("cat", "ALARM"),
                "location": fields["location"].get(),
                "desc": info["desc"], "cat": info["cat"],
                "fn": info["fn"], "vi_code": info.get("vi_code", ""),
            }
            self.devices.append(dev)
            self.tree.insert("", "end", values=(
                dev["address"], dev["device_type"],
                dev["signal_type"], dev["location"], dev["desc"]))
            self.device_count_var.set(f"{len(self.devices)} devices")
            self._update_summary()
            if keep_open:
                # Select all text in the location field so user can type over it
                w = loc_entry_ref["widget"]
                w.focus_set()
                w.select_range(0, "end")
            else:
                dlg.destroy()

        btn_row = ttk.Frame(dlg, padding=(16, 4))
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Save & Add Another",
                   command=lambda: _save(keep_open=True)).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Save & Close",
                   command=lambda: _save(keep_open=False),
                   style="Accent.TButton").pack(side="right", padx=4)
        ttk.Button(btn_row, text="Cancel", command=dlg.destroy).pack(side="right")

    def _browse_logo(self):
        path = filedialog.askopenfilename(title="Select Logo Image",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.gif"), ("All Files", "*.*")])
        if path:
            self.config["logo_path"] = path
            save_config(self.config)
            self._load_logo_preview()

    def _load_logo_preview(self):
        logo = self.config.get("logo_path", "")
        if logo and os.path.exists(logo):
            try:
                img = Image.open(logo)
                img.thumbnail((140, 50))
                self.logo_image = ImageTk.PhotoImage(img)
                self.logo_label.config(image=self.logo_image, text="")
                return
            except Exception:
                pass
        self.logo_label.config(image="", text="No logo selected")
        self.logo_image = None

    def _clear_logo(self):
        self.config["logo_path"] = ""
        save_config(self.config)
        self.logo_label.config(image="", text="No logo selected")
        self.logo_image = None

    # ─── VI HELPERS ────────────────────────────────────────────────────────────

    # ─── OPERATING INSTRUCTIONS ──────────────────────────────────────────────

    def _update_oi_preview(self):
        if not hasattr(self, 'tp_oi_preview_var'):
            return
        if not self.tp_include_oi.get():
            self.tp_oi_preview_var.set("Not included.")
            return
        series = self.panel_series_var.get()
        touchscreen = self.touchscreen_var.get()
        docs = get_operating_instruction_paths(series, touchscreen)
        if not docs:
            self.tp_oi_preview_var.set(f"No instruction file found for {series} in the OPERATING INSTRUCTIONS folder.")
            return
        names = ", ".join(label for label, _ in docs)
        self.tp_oi_preview_var.set(f"Will include: {names}")

    def _apply_fsrc_autodetect(self):
        self._apply_vi_autodetect()

    def _apply_vi_autodetect(self):
        """Auto-populate VI fields and checklist items from loaded CSV data."""
        if not hasattr(self, 'vi_vars'):
            return

        cv = self.vi_checklist_vars if hasattr(self, 'vi_checklist_vars') else {}

        def set_var(key, value):
            if key in self.vi_vars:
                self.vi_vars[key].set(value)

        def set_check(key, value):
            if key in cv:
                cv[key].set(value)

        def set_all_section(prefix, value):
            for key in cv:
                if key.startswith(prefix):
                    cv[key].set(value)

        # ── Sync basic fields from SOO tab ───────────────────────────────────
        if self.job_var.get():
            set_var("building_name", self.job_var.get())
        if self.panel_var.get():
            set_var("model_number", self.panel_var.get())

        # ── Detect what device categories are actually installed ─────────────
        dtypes = {d["device_type"] for d in self.devices}
        fns    = {d["fn"] for d in self.devices}
        has_smoke     = any(d["cat"] in ("SMOKE", "MULTI") for d in self.devices)
        has_heat      = any(d["cat"] == "HEAT" for d in self.devices)
        has_co        = any(d["cat"] == "CO" for d in self.devices)
        has_pull      = any(d["cat"] == "PULL" for d in self.devices)
        has_supv      = any(d["fn"] == "SUPV" for d in self.devices)
        has_waterflow = any(d["cat"] == "WFLOW" for d in self.devices)
        has_relay     = any(d["cat"] in ("RELAY",) and d["fn"] == "OUTPUT" for d in self.devices)
        has_nac       = bool(self.nac_circuits) or any(d["cat"] == "NAC" for d in self.devices)

        # ── 32.14 FSRC / Monitoring ──────────────────────────────────────────
        connected = bool(self.city_relays)
        set_var("fsrc_connected", "Yes" if connected else "No")

        if hasattr(self, 'vi_fsrc_detected_var'):
            if connected:
                kinds = ", ".join(sorted({r["kind"] for r in self.city_relays}))
                self.vi_fsrc_detected_var.set(
                    f"Detected {len(self.city_relays)} monitoring relay(s): {kinds}. "
                    "Enter FSRC name/phone manually.")
            else:
                self.vi_fsrc_detected_var.set("No monitoring relays found — system appears stand-alone.")

        if not connected:
            # No monitoring: entire 32.14 section is N/A
            set_all_section("32.14_", "A")
        else:
            # Connected: mark tested relays Y, rest N/A
            kinds_present = {r["kind"] for r in self.city_relays}
            relay_map = {
                "32.14_G": "Fire Alarm",
                "32.14_I": "Supervisory",
                "32.14_J": "Trouble",
            }
            for var_key, kind in relay_map.items():
                set_check(var_key, "Y" if kind in kinds_present else "A")

        # ── 2-stage vs single-stage (from CC data or system type) ───────────
        sys_type = (self.cc_data or {}).get("system_type", "single_stage")
        is_two_stage = sys_type == "two_stage"

        # Sheet 28 A/B: single-stage vs two-stage
        set_check("28_A", "Y" if not is_two_stage else "N")
        set_check("28_B", "Y" if is_two_stage else "N")

        # Sheet 32.2: alert/two-stage items
        two_stage_items = ["32.2_G", "32.2_I", "32.2_J", "32.2_K", "32.2_P"]
        for k in two_stage_items:
            set_check(k, "Y" if is_two_stage else "N")

        # ── Waterflow response time: N/A if no waterflow devices ─────────────
        if not has_waterflow:
            set_check("32.5_I", "A")

        self._update_vi_preview()

    def _update_batt_85_check(self):
        """Auto-calculate 85% battery voltage check and update checklist item 32.8_S."""
        if not hasattr(self, 'vi_vars') or not hasattr(self, 'batt_result_var'):
            return
        try:
            vat = float(self.vi_vars.get("voltage_after_test", tk.StringVar()).get())
            btype = self.vi_vars.get("battery_type", tk.StringVar()).get()
            # Parse rated voltage from battery_type string e.g. "12V 12AH" or "24V 7AH"
            m = re.search(r"(\d+(?:\.\d+)?)\s*[Vv]", btype)
            if not m:
                self.batt_result_var.set("")
                return
            rated = float(m.group(1))
            threshold = rated * 0.85
            passed = vat >= threshold
            pct = (vat / rated * 100) if rated else 0
            result = f"{'✓ PASS' if passed else '✗ FAIL'}  {vat:.1f}V / {rated:.0f}V  ({pct:.1f}%)  min {threshold:.1f}V"
            self.batt_result_var.set(result)
            # Update checklist item
            if hasattr(self, 'vi_checklist_vars') and "32.8_S" in self.vi_checklist_vars:
                self.vi_checklist_vars["32.8_S"].set("Y" if passed else "N")
        except (ValueError, ZeroDivisionError):
            self.batt_result_var.set("")

    # ─── ATTACHMENTS ─────────────────────────────────────────────────────────

    def _add_attachment(self):
        paths = filedialog.askopenfilenames(title="Select PDF Document(s)",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        for p in paths:
            name = os.path.splitext(os.path.basename(p))[0]
            self.attached_docs.append({"path": p, "name": name})
            self.attach_tree.insert("", "end", values=(name, os.path.basename(p)))

    def _remove_attachment(self):
        sel = self.attach_tree.selection()
        if not sel:
            return
        idx = self.attach_tree.index(sel[0])
        self.attach_tree.delete(sel[0])
        del self.attached_docs[idx]

    def _move_attachment(self, direction):
        sel = self.attach_tree.selection()
        if not sel:
            return
        idx = self.attach_tree.index(sel[0])
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.attached_docs):
            return
        self.attached_docs[idx], self.attached_docs[new_idx] = self.attached_docs[new_idx], self.attached_docs[idx]
        self._refresh_attach_tree()
        children = self.attach_tree.get_children()
        self.attach_tree.selection_set(children[new_idx])

    def _rename_attachment(self):
        sel = self.attach_tree.selection()
        if not sel:
            messagebox.showinfo("Rename", "Select a document to rename its tab.")
            return
        idx = self.attach_tree.index(sel[0])
        current_name = self.attached_docs[idx]["name"]

        win = tk.Toplevel(self.root)
        win.title("Rename Tab")
        win.geometry("350x100")
        win.transient(self.root)
        win.grab_set()
        ttk.Label(win, text="Tab Name:").pack(padx=10, pady=(10, 4), anchor="w")
        name_var = tk.StringVar(value=current_name)
        ent = ttk.Entry(win, textvariable=name_var, width=40)
        ent.pack(padx=10, fill="x")
        ent.select_range(0, "end")
        ent.focus()
        def save():
            self.attached_docs[idx]["name"] = name_var.get()
            self._refresh_attach_tree()
            win.destroy()
        ttk.Button(win, text="Save", command=save).pack(pady=8)
        ent.bind("<Return>", lambda e: save())

    def _refresh_attach_tree(self):
        self.attach_tree.delete(*self.attach_tree.get_children())
        for doc in self.attached_docs:
            self.attach_tree.insert("", "end", values=(doc["name"], os.path.basename(doc["path"])))

    # ─── EXPORT: SOO PDF ─────────────────────────────────────────────────────

    def _export_soo_pdf(self):
        if not self.devices:
            messagebox.showwarning("Export", "No devices loaded.")
            return
        self.config["company_name"] = self.company_var.get()
        for key, var in self.section_vars.items():
            self.config[key] = var.get()
        save_config(self.config)

        path = filedialog.asksaveasfilename(title="Save SOO PDF",
            initialdir=os.path.dirname(self.csv_path) if self.csv_path else "",
            initialfile=f"{self.job_var.get()}_SOO.pdf",
            defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not path:
            return
        self.status_var.set("Generating SOO PDF...")
        self.root.update()
        try:
            generate_soo_pdf(self.devices, self.config, path,
                           self.job_var.get(), self.panel_var.get(),
                           self.notes_text.get("1.0", "end"),
                           nac_circuits=self.nac_circuits, bypass_switches=self.bypass_switches)
            self.status_var.set(f"SOO exported: {os.path.basename(path)}")
            if messagebox.askyesno("Success", f"SOO PDF saved to:\n{path}\n\nOpen it now?"):
                os.startfile(path)
        except Exception as e:
            self.status_var.set("Export failed")
            messagebox.showerror("Export Error", str(e))

    # ─── EXPORT: VI ──────────────────────────────────────────────────────────

    def _add_manual_device(self):
        DEVICE_CODES = ["H/V", "H", "V", "S", "B", "SP", "HSP", "NAC", "AD", "EOL", "DS", "K", "AUX"]

        dlg = tk.Toplevel(self.root)
        dlg.title("Add Non-Addressable Device")
        dlg.geometry("420x380")
        dlg.transient(self.root)
        dlg.grab_set()

        fields = {}

        # Floor
        row = ttk.Frame(dlg)
        row.pack(fill="x", padx=12, pady=3)
        ttk.Label(row, text="Floor:", width=14, anchor="w").pack(side="left")
        fields["floor"] = tk.StringVar()
        ttk.Entry(row, textvariable=fields["floor"], width=28).pack(side="left", fill="x", expand=True)

        # Location
        row = ttk.Frame(dlg)
        row.pack(fill="x", padx=12, pady=3)
        ttk.Label(row, text="Location:", width=14, anchor="w").pack(side="left")
        fields["location"] = tk.StringVar()
        ttk.Entry(row, textvariable=fields["location"], width=28).pack(side="left", fill="x", expand=True)

        # Type - autocomplete combobox (type to filter)
        row = ttk.Frame(dlg)
        row.pack(fill="x", padx=12, pady=3)
        ttk.Label(row, text="Device Type:", width=14, anchor="w").pack(side="left")
        fields["type"] = tk.StringVar()
        type_combo = ttk.Combobox(row, textvariable=fields["type"], values=DEVICE_CODES, width=10)
        type_combo.pack(side="left")
        # Autocomplete: filter list as you type
        def _type_keyrelease(event):
            typed = fields["type"].get().upper()
            if not typed:
                type_combo['values'] = DEVICE_CODES
                return
            filtered = [c for c in DEVICE_CODES if c.startswith(typed)]
            type_combo['values'] = filtered if filtered else DEVICE_CODES
        type_combo.bind('<KeyRelease>', _type_keyrelease)

        # Circuit - populated from detected conventional NAC circuits when available
        row = ttk.Frame(dlg)
        row.pack(fill="x", padx=12, pady=3)
        ttk.Label(row, text="Circuit:", width=14, anchor="w").pack(side="left")
        fields["circuit"] = tk.StringVar()
        conv_circuits = [c["id"] for c in getattr(self, 'nac_circuits', []) if not c["addressable"]]
        circuit_combo = ttk.Combobox(row, textvariable=fields["circuit"],
                                      values=conv_circuits if conv_circuits else [], width=18)
        circuit_combo.pack(side="left")
        if conv_circuits:
            circuit_label_var = tk.StringVar(value="")
            circuit_label_lbl = ttk.Label(dlg, textvariable=circuit_label_var,
                                           font=("Segoe UI", 8), foreground="#666")
            circuit_label_lbl.pack(anchor="w", padx=12)
            def _circuit_changed(*_a):
                cid = fields["circuit"].get()
                match = next((c for c in self.nac_circuits if c["id"] == cid), None)
                circuit_label_var.set(match["label"] if match else "")
            fields["circuit"].trace_add("write", _circuit_changed)

        # Checkboxes for test results
        ttk.Separator(dlg, orient="horizontal").pack(fill="x", padx=12, pady=6)
        ttk.Label(dlg, text="Test Results:", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12)

        chk_frame = ttk.Frame(dlg)
        chk_frame.pack(fill="x", padx=12, pady=4)
        fields["installed"] = tk.BooleanVar(value=True)
        fields["operation"] = tk.BooleanVar(value=True)
        fields["supervision"] = tk.BooleanVar(value=True)
        ttk.Checkbutton(chk_frame, text="Correctly Installed", variable=fields["installed"]).pack(anchor="w")
        ttk.Checkbutton(chk_frame, text="Alarm / Operation Confirmed", variable=fields["operation"]).pack(anchor="w")
        ttk.Checkbutton(chk_frame, text="Supervision of Wiring Confirmed", variable=fields["supervision"]).pack(anchor="w")

        # Buttons
        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(fill="x", padx=12, pady=8)

        def _do_add(close_after=True):
            dev = {
                "floor": fields["floor"].get().strip(),
                "location": fields["location"].get().strip(),
                "type": fields["type"].get().strip().upper(),
                "circuit": fields["circuit"].get().strip(),
                "installed": "P" if fields["installed"].get() else "",
                "operation": "P" if fields["operation"].get() else "",
                "supervision": "P" if fields["supervision"].get() else "",
            }
            if not dev["type"]:
                messagebox.showwarning("Missing", "Device type is required.", parent=dlg)
                return
            self.vi_manual_devices.append(dev)
            self.vi_manual_tree.insert("", "end", values=(
                dev["floor"], dev["location"], dev["type"], dev["circuit"],
                dev["installed"], dev["operation"], dev["supervision"]))
            self._update_vi_preview()
            if close_after:
                dlg.destroy()
            else:
                # Clear for next entry, keep floor and circuit
                fields["location"].set("")
                fields["type"].set("")
                type_combo.focus()

        ttk.Button(btn_frame, text="Save & Add Another", command=lambda: _do_add(False)).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="Save & Close", command=lambda: _do_add(True), style="Accent.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side="left")

        type_combo.focus()

    def _add_ancillary(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Add Ancillary Device")
        dlg.geometry("400x200")
        dlg.transient(self.root)
        dlg.grab_set()

        fields = {}
        for label, key in [("Circuit / Device:", "device"), ("Powered By:", "powered_by"),
                            ("Confirmation Method:", "method")]:
            row = ttk.Frame(dlg)
            row.pack(fill="x", padx=12, pady=4)
            ttk.Label(row, text=label, width=18, anchor="w").pack(side="left")
            var = tk.StringVar()
            if key == "method":
                var.set("VISUAL")
            fields[key] = var
            ttk.Entry(row, textvariable=var, width=28).pack(side="left", fill="x", expand=True)

        row = ttk.Frame(dlg)
        row.pack(fill="x", padx=12, pady=4)
        ttk.Label(row, text="Confirmed:", width=18, anchor="w").pack(side="left")
        fields["confirmed"] = tk.StringVar(value="Yes")
        ttk.Combobox(row, textvariable=fields["confirmed"], values=["Yes", "No"], width=6, state="readonly").pack(side="left")

        def do_add():
            dev = {k: v.get().strip() for k, v in fields.items()}
            if not dev["device"]:
                messagebox.showwarning("Missing", "Device name required.", parent=dlg)
                return
            self.vi_ancillary_devices.append(dev)
            self.vi_ancillary_tree.insert("", "end", values=(
                dev["device"], dev["powered_by"], dev["confirmed"], dev["method"]))
            dlg.destroy()

        ttk.Button(dlg, text="Save", command=do_add, style="Accent.TButton").pack(pady=8)

    def _remove_ancillary(self):
        sel = self.vi_ancillary_tree.selection()
        if not sel:
            return
        idx = self.vi_ancillary_tree.index(sel[0])
        self.vi_ancillary_tree.delete(sel[0])
        if idx < len(self.vi_ancillary_devices):
            self.vi_ancillary_devices.pop(idx)

    def _remove_manual_device(self):
        sel = self.vi_manual_tree.selection()
        if not sel:
            return
        for item in sel:
            idx = self.vi_manual_tree.index(item)
            self.vi_manual_tree.delete(item)
            if idx < len(self.vi_manual_devices):
                self.vi_manual_devices.pop(idx)
        self._update_vi_preview()

    def _collect_vi_data(self):
        """Collect all VI form data from the UI into a dict."""
        vi_data = {k: v.get() for k, v in self.vi_vars.items()}
        vi_data["remarks"] = self.vi_remarks.get("1.0", "end").strip()
        vi_data["manual_devices"] = self.vi_manual_devices
        vi_data["ancillary_devices"] = getattr(self, 'vi_ancillary_devices', [])

        # Sound level data
        sound_levels = []
        for row_vars in self.vi_sound_entries:
            row_data = {k: v.get().strip() for k, v in row_vars.items()}
            if any(row_data.values()):
                sound_levels.append(row_data)
        vi_data["sound_levels"] = sound_levels

        # 33.3 Subsequent alarm tests
        alarm_tests = []
        for row_vars in getattr(self, 'vi_333_entries', []):
            row_data = {k: v.get().strip() for k, v in row_vars.items()}
            if row_data.get("init_zone") or row_data.get("init_dev"):
                alarm_tests.append(row_data)
        vi_data["alarm_tests_333"] = alarm_tests

        # 33.4 Fault tolerance tests
        fault_tests = []
        for row_vars in getattr(self, 'vi_334_entries', []):
            row_data = {k: v.get().strip() for k, v in row_vars.items()}
            if row_data.get("fault_loc"):
                fault_tests.append(row_data)
        vi_data["fault_tolerance_tests"] = fault_tests

        return vi_data

    def _export_vi(self):
        if not self.devices:
            messagebox.showwarning("Export", "No devices loaded. Load a report PDF or CSV on the SOO tab first.")
            return

        vi_data = self._collect_vi_data()

        # Save technician info for next time
        self.config["technician_name"] = vi_data.get("tech_name", "")
        self.config["technician_cert"] = vi_data.get("tech_cert", "")
        save_config(self.config)

        path = filedialog.asksaveasfilename(title="Save VI Report",
            initialdir=os.path.dirname(self.csv_path) if self.csv_path else "",
            initialfile=f"{self.job_var.get()} VI.xlsm",
            defaultextension=".xlsm", filetypes=[("Excel Files", "*.xlsm"), ("All Files", "*.*")])
        if not path:
            return

        self.vi_status_var.set("Generating VI Report...")
        self.root.update()
        try:
            checklist_overrides = {k: v.get() for k, v in self.vi_checklist_vars.items()}
            generate_vi_report(path, self.devices, vi_data, self.panel_var.get(), checklist_overrides)
            self.vi_status_var.set(f"VI Report generated: {os.path.basename(path)}")
            if messagebox.askyesno("Success",
                    f"CAN/ULC-S537 Verification Report saved to:\n{path}\n\n"
                    f"{len(self.devices)} devices auto-filled across all sheets.\n"
                    "Open in Excel to review?"):
                os.startfile(path)
        except Exception as e:
            self.vi_status_var.set("Export failed")
            messagebox.showerror("Export Error", str(e))

    def _export_vi_pdf(self):
        """Export VI report as PDF via Excel COM automation."""
        if not self.devices:
            messagebox.showwarning("Export", "No devices loaded.")
            return

        vi_data = self._collect_vi_data()

        pdf_path = filedialog.asksaveasfilename(title="Save VI Report as PDF",
            initialdir=os.path.dirname(self.csv_path) if self.csv_path else "",
            initialfile=f"{self.job_var.get()} VI.pdf",
            defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if not pdf_path:
            return

        self.vi_status_var.set("Generating VI Report for PDF...")
        self.root.update()

        import tempfile
        tmp_xlsm = os.path.join(tempfile.gettempdir(), f"vi_temp_{os.getpid()}.xlsm")

        try:
            checklist_overrides = {k: v.get() for k, v in self.vi_checklist_vars.items()}
            generate_vi_report(tmp_xlsm, self.devices, vi_data, self.panel_var.get(), checklist_overrides)

            try:
                import win32com.client
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(os.path.abspath(tmp_xlsm))
                wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
                wb.Close(False)
                excel.Quit()
                self.vi_status_var.set(f"PDF exported: {os.path.basename(pdf_path)}")
                if messagebox.askyesno("Success",
                        f"VI Report PDF saved to:\n{pdf_path}\n\nOpen it now?"):
                    os.startfile(pdf_path)
            except ImportError:
                # win32com not available - save xlsm and inform user
                fallback_xlsm = pdf_path.replace(".pdf", ".xlsm")
                shutil.copy2(tmp_xlsm, fallback_xlsm)
                self.vi_status_var.set("PDF requires Excel - .xlsm saved instead")
                messagebox.showinfo("PDF Export",
                    f"PDF export requires Microsoft Excel and the pywin32 package.\n\n"
                    f"The .xlsm has been saved to:\n{fallback_xlsm}\n\n"
                    "Open it in Excel and use File > Save As > PDF.")
            except Exception as e:
                self.vi_status_var.set("PDF export failed")
                messagebox.showerror("PDF Export Error",
                    f"Excel PDF conversion failed:\n{str(e)}\n\n"
                    f"The .xlsm has been saved at:\n{tmp_xlsm}")
        except Exception as e:
            self.vi_status_var.set("Export failed")
            messagebox.showerror("Export Error", str(e))
        finally:
            if os.path.exists(tmp_xlsm):
                try:
                    os.remove(tmp_xlsm)
                except Exception:
                    pass

    # ─── TAB 4: PROJECTS ─────────────────────────────────────────────────────

    def _build_projects_tab(self):
        proj_frame = ttk.Frame(self.notebook)
        self.notebook.add(proj_frame, text="  Projects  ")

        top = ttk.Frame(proj_frame, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="Saved Jobs",
                  font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Button(top, text="Save Current Job",
                   command=self._archive_project, style="Accent.TButton").pack(side="right", padx=4)
        ttk.Button(top, text="Refresh", command=self._refresh_projects).pack(side="right")

        ttk.Label(proj_frame,
                  text="Click \"Save Current Job\" to save all devices, job info, and documents for this job. "
                       "Double-click a saved job below to reopen it and keep editing — no CSV/PDF re-upload needed.",
                  foreground="#666", font=("Segoe UI", 8), padding=(8, 0, 8, 4),
                  wraplength=700).pack(anchor="w")

        cols = ("job", "date", "panel", "files")
        self.proj_tree = ttk.Treeview(proj_frame, columns=cols, show="headings", selectmode="browse")
        self.proj_tree.heading("job",   text="Job Name")
        self.proj_tree.heading("date",  text="Date Saved")
        self.proj_tree.heading("panel", text="Panel")
        self.proj_tree.heading("files", text="Files")
        self.proj_tree.column("job",   width=180)
        self.proj_tree.column("date",  width=120)
        self.proj_tree.column("panel", width=130)
        self.proj_tree.column("files", width=300)
        psb = ttk.Scrollbar(proj_frame, orient="vertical", command=self.proj_tree.yview)
        self.proj_tree.configure(yscrollcommand=psb.set)
        self.proj_tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        psb.pack(side="right", fill="y", pady=8, padx=(0, 8))

        self.proj_tree.bind("<Double-1>", lambda e: self._open_project())

        btn_row = ttk.Frame(proj_frame)
        btn_row.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btn_row, text="Open Selected", command=self._open_project).pack(side="left", padx=2)
        ttk.Button(btn_row, text="Open Folder", command=self._open_project_folder).pack(side="left", padx=2)
        ttk.Button(btn_row, text="Delete Archive", command=self._delete_project).pack(side="right", padx=2)

        self._refresh_projects()

    def _projects_dir(self):
        d = os.path.join(CONFIG_DIR, "Projects")
        os.makedirs(d, exist_ok=True)
        return d

    def _refresh_projects(self):
        if not hasattr(self, 'proj_tree'):
            return
        self.proj_tree.delete(*self.proj_tree.get_children())
        pdir = self._projects_dir()
        entries = []
        for name in os.listdir(pdir):
            meta_path = os.path.join(pdir, name, "meta.json")
            if os.path.isfile(meta_path):
                try:
                    with open(meta_path, "r", encoding="utf-8") as f:
                        meta = json.load(f)
                    entries.append((meta.get("date_saved", ""), name, meta))
                except Exception:
                    pass
        entries.sort(key=lambda x: x[0], reverse=True)
        for _, folder, meta in entries:
            proj_path = os.path.join(pdir, folder)
            files = [f for f in os.listdir(proj_path) if f not in ("meta.json", "project_data.json")]
            self.proj_tree.insert("", "end", iid=folder, values=(
                meta.get("job", folder),
                meta.get("date_saved", "")[:10],
                meta.get("panel", ""),
                ", ".join(files),
            ))

    def _archive_project(self):
        if not self.devices and not self.attached_docs:
            messagebox.showwarning("Archive", "Nothing to archive — load a job first.")
            return
        job = self.job_var.get() or "Unnamed"
        date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"{job}_{date_str}"
        proj_path = os.path.join(self._projects_dir(), folder_name)
        os.makedirs(proj_path, exist_ok=True)

        saved = []
        # Save SOO HTML (rendered snapshot, for quick viewing/printing)
        soo_path = os.path.join(proj_path, f"{job}_SOO.html")
        try:
            soo_html = self._generate_soo_html_content()
            with open(soo_path, "w", encoding="utf-8") as f:
                f.write(soo_html)
            saved.append(f"{job}_SOO.html")
        except Exception:
            pass

        # Copy any already-generated attached docs alongside, and record them
        # so they can be re-attached when this project is reopened.
        doc_records = []
        for doc in self.attached_docs:
            if os.path.isfile(doc["path"]):
                fname = os.path.basename(doc["path"])
                dest = os.path.join(proj_path, fname)
                shutil.copy2(doc["path"], dest)
                saved.append(fname)
                doc_records.append({"path": fname, "name": doc.get("name", fname)})

        # Save the full working data set so this job can be reopened and edited later.
        project_data = {
            "company": self.company_var.get(),
            "job": job,
            "panel": self.panel_var.get(),
            "panel_series": self.panel_series_var.get(),
            "touchscreen": self.touchscreen_var.get(),
            "notes": self.notes_text.get("1.0", "end").strip(),
            "code_l2": self.code_l2_var.get(),
            "devices": self.devices,
            "nodes": self.nodes,
            "nac_circuits": self.nac_circuits,
            "city_relays": self.city_relays,
            "bypass_switches": self.bypass_switches,
            "cc_data": self.cc_data,
            "attached_docs": doc_records,
        }
        with open(os.path.join(proj_path, "project_data.json"), "w", encoding="utf-8") as f:
            json.dump(project_data, f, indent=2)

        meta = {
            "job": job,
            "date_saved": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "panel": self.panel_var.get(),
            "devices": len(self.devices),
            "files": saved,
        }
        with open(os.path.join(proj_path, "meta.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2)

        self._refresh_projects()
        messagebox.showinfo("Saved",
            f"'{job}' saved with full working data ({len(self.devices)} devices).\n"
            f"Folder: {proj_path}\n\n"
            f"Double-click it in this list any time to reopen and keep editing.")

    def _open_project(self):
        sel = self.proj_tree.selection()
        if not sel:
            return
        folder = sel[0]
        proj_path = os.path.join(self._projects_dir(), folder)
        data_path = os.path.join(proj_path, "project_data.json")
        if os.path.isfile(data_path):
            if self.devices or self.attached_docs:
                if not messagebox.askyesno("Open Project",
                        "This will replace your current unsaved work in the app with "
                        "the saved project.\n\nContinue?"):
                    return
            self._restore_project_data(data_path, proj_path)
            return
        # Older archive with no saved working data — just open the rendered HTML.
        for fname in os.listdir(proj_path):
            if fname.endswith((".html", ".pdf")):
                os.startfile(os.path.join(proj_path, fname))
                return
        os.startfile(proj_path)

    def _restore_project_data(self, data_path, proj_path):
        try:
            with open(data_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Open Project", f"Could not read saved project data:\n{e}")
            return

        self.company_var.set(data.get("company", ""))
        self.job_var.set(data.get("job", ""))
        self.panel_var.set(data.get("panel", ""))
        self.panel_series_var.set(data.get("panel_series", "4007ES"))
        self.touchscreen_var.set(data.get("touchscreen", False))
        self.notes_text.delete("1.0", "end")
        self.notes_text.insert("1.0", data.get("notes", ""))
        self.code_l2_var.set(data.get("code_l2", "222222"))

        self.devices = data.get("devices", [])
        self.nodes = data.get("nodes", [])
        self.nac_circuits = data.get("nac_circuits", [])
        self.city_relays = data.get("city_relays", [])
        self.bypass_switches = data.get("bypass_switches", [])
        self.cc_data = data.get("cc_data")
        self.csv_path = None

        self.attached_docs = []
        for rec in data.get("attached_docs", []):
            src = os.path.join(proj_path, rec["path"])
            if os.path.isfile(src):
                self.attached_docs.append({"path": src, "name": rec.get("name", rec["path"])})

        self.file_var.set(f"Loaded from saved project: {data.get('job', '')}")
        self._refresh_node_list()
        self._populate_tree()
        self._update_summary()
        self._apply_fsrc_autodetect()
        self._update_oi_preview()
        if hasattr(self, "attach_tree"):
            self._refresh_attach_tree()

        for btn in [self.export_btn, self.export_html_btn, self.export_vi_btn,
                    self.export_vi_pdf_btn, self.tp_export_btn, self.tp_export_pdf_btn]:
            btn.config(state="normal")

        self.notebook.select(0)
        messagebox.showinfo("Project Loaded",
            f"'{data.get('job', '')}' loaded — {len(self.devices)} devices. You can now edit and re-save.")

    def _open_project_folder(self):
        sel = self.proj_tree.selection()
        if not sel:
            return
        os.startfile(os.path.join(self._projects_dir(), sel[0]))

    def _delete_project(self):
        sel = self.proj_tree.selection()
        if not sel:
            return
        folder = sel[0]
        job = self.proj_tree.item(folder, "values")[0]
        if messagebox.askyesno("Delete Archive",
                f"Permanently delete archive for '{job}'?\nThis cannot be undone."):
            shutil.rmtree(os.path.join(self._projects_dir(), folder), ignore_errors=True)
            self._refresh_projects()

    # ─── EXPORT: TURNOVER HTML ───────────────────────────────────────────────

    def _export_turnover(self):
        if not self.devices:
            messagebox.showwarning("Export", "No devices loaded.")
            return

        include_soo = self.tp_include_soo.get()
        if not include_soo and not self.attached_docs and not self.tp_include_oi.get():
            messagebox.showwarning("Export", "Include at least one document.")
            return

        path = filedialog.asksaveasfilename(title="Save Turnover Package HTML",
            initialdir=os.path.dirname(self.csv_path) if self.csv_path else "",
            initialfile=f"{self.job_var.get()}_Turnover.html",
            defaultextension=".html", filetypes=[("HTML Files", "*.html")])
        if not path:
            return

        self.tp_status_var.set("Building turnover package...")
        self.root.update()

        try:
            doc_tabs = []

            if include_soo:
                self.tp_status_var.set("Generating SOO content...")
                self.root.update()
                # Generate the interactive SOO HTML content
                soo_html = self._generate_soo_html_content()
                doc_tabs.append({
                    "name": self.tp_soo_name.get() or "Sequence of Operations",
                    "type": "html",
                    "content": soo_html,
                })

            if self.tp_include_oi.get():
                oi_docs = get_operating_instruction_paths(self.panel_series_var.get(), self.touchscreen_var.get())
                for label, oi_path in oi_docs:
                    self.tp_status_var.set(f"Encoding {label}...")
                    self.root.update()
                    with open(oi_path, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode("ascii")
                    doc_tabs.append({
                        "name": label,
                        "type": "pdf",
                        "content": b64,
                    })

            skipped = []
            for doc in self.attached_docs:
                if os.path.exists(doc["path"]):
                    self.tp_status_var.set(f"Encoding {doc['name']}...")
                    self.root.update()
                    with open(doc["path"], "rb") as f:
                        b64 = base64.b64encode(f.read()).decode("ascii")
                    doc_tabs.append({
                        "name": doc["name"],
                        "type": "pdf",
                        "content": b64,
                    })
                else:
                    skipped.append(doc["name"])

            self.tp_status_var.set("Assembling HTML...")
            self.root.update()
            generate_turnover_html(path, None, doc_tabs)

            self.tp_status_var.set(f"Turnover package exported ({len(doc_tabs)} tabs)")
            if skipped:
                messagebox.showwarning("Missing Attachment(s)",
                    "These attached documents could not be found on disk and were "
                    "left out of the export:\n\n" + "\n".join(f"- {n}" for n in skipped) +
                    "\n\nRe-add them on the Turnover Package tab and export again.")
            if messagebox.askyesno("Success",
                    f"Turnover package saved to:\n{path}\n\n"
                    f"{len(doc_tabs)} tab(s). Self-contained HTML - customer opens in any browser.\n\nOpen it now?"):
                os.startfile(path)
        except Exception as e:
            self.tp_status_var.set("Export failed")
            messagebox.showerror("Export Error", str(e))

    def _generate_soo_html_content(self):
        """Generate the SOO interactive HTML with device data baked in."""
        csv_lines = []
        for dev in self.devices:
            line = " {}, {}, {}, {},".format(
                dev["address"], dev["device_type"], dev["signal_type"], dev["location"])
            csv_lines.append(line)
        csv_data = "\\n".join(csv_lines)
        job_name = self.job_var.get()

        nac_circuits_json = json.dumps(getattr(self, 'nac_circuits', []))
        city_relays_json = json.dumps(getattr(self, 'city_relays', []))
        bypass_switches_json = json.dumps(getattr(self, 'bypass_switches', []))
        cc_data_json = json.dumps(getattr(self, 'cc_data', None))
        access_codes_json = json.dumps({
            "level1": getattr(self, 'code_l1_var', tk.StringVar()).get(),
            "level2": getattr(self, 'code_l2_var', tk.StringVar()).get(),
        })
        panel_series_json = json.dumps(getattr(self, 'panel_series_var', tk.StringVar(value="4007ES")).get())

        template_path = os.path.join(os.path.dirname(__file__), "soo_interactive.html")
        with open(template_path, "r", encoding="utf-8") as f:
            html = f.read()

        boot_script = """
<script>
(function() {
  const csvData = `%s`;
  window._fileName = "%s.csv";
  window._nacCircuits = %s;
  window._cityRelays = %s;
  window._bypassSwitches = %s;
  window._ccData = %s;
  window._accessCodes = %s;
  window._panelSeries = %s;
  window.addEventListener('DOMContentLoaded', function() {
    document.getElementById('upload-screen').style.display = 'none';
    document.getElementById('app').style.display = 'block';
    var devices = parseCSV(csvData);
    renderApp(devices);
  });
})();
</script>
</body>""" % (csv_data, job_name, nac_circuits_json, city_relays_json, bypass_switches_json, cc_data_json, access_codes_json, panel_series_json)
        html = html.replace("</body>", boot_script)
        return html

    # ─── EXPORT: TURNOVER PDF (merged) ───────────────────────────────────────

    def _export_turnover_pdf(self):
        if not self.devices and not self.attached_docs:
            messagebox.showwarning("Export", "No content to merge.")
            return

        path = filedialog.asksaveasfilename(title="Save Turnover Package PDF",
            initialdir=os.path.dirname(self.csv_path) if self.csv_path else "",
            initialfile=f"{self.job_var.get()}_Turnover.pdf",
            defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not path:
            return

        self.tp_status_var.set("Building merged PDF...")
        self.root.update()

        try:
            import tempfile
            temp_dir = tempfile.mkdtemp()
            pdfs = []

            if self.tp_include_soo.get() and self.devices:
                self.tp_status_var.set("Generating SOO PDF...")
                self.root.update()
                self.config["company_name"] = self.company_var.get()
                for key, var in self.section_vars.items():
                    self.config[key] = var.get()
                soo_path = os.path.join(temp_dir, "soo.pdf")
                generate_soo_pdf(self.devices, self.config, soo_path,
                               self.job_var.get(), self.panel_var.get(),
                               self.notes_text.get("1.0", "end"),
                               nac_circuits=self.nac_circuits, bypass_switches=self.bypass_switches)
                pdfs.append(soo_path)

            if self.tp_include_oi.get():
                oi_docs = get_operating_instruction_paths(self.panel_series_var.get(), self.touchscreen_var.get())
                for _label, oi_path in oi_docs:
                    pdfs.append(oi_path)

            skipped = []
            for doc in self.attached_docs:
                if os.path.exists(doc["path"]):
                    pdfs.append(doc["path"])
                else:
                    skipped.append(doc["name"])

            if not pdfs:
                messagebox.showwarning("Export", "No PDFs to merge.")
                return

            self.tp_status_var.set("Merging documents...")
            self.root.update()

            if len(pdfs) == 1:
                shutil.copy2(pdfs[0], path)
            else:
                merge_pdfs(pdfs, path)

            # Cleanup
            for f in os.listdir(temp_dir):
                os.remove(os.path.join(temp_dir, f))
            os.rmdir(temp_dir)

            self.tp_status_var.set(f"Merged PDF exported ({len(pdfs)} documents)")
            if skipped:
                messagebox.showwarning("Missing Attachment(s)",
                    "These attached documents could not be found on disk and were "
                    "left out of the merge:\n\n" + "\n".join(f"- {n}" for n in skipped) +
                    "\n\nRe-add them on the Turnover Package tab and export again.")
            if messagebox.askyesno("Success", f"Merged PDF saved to:\n{path}\n\nOpen it now?"):
                os.startfile(path)
        except Exception as e:
            self.tp_status_var.set("Export failed")
            messagebox.showerror("Export Error", str(e))

    # ─── EXPORT: CUSTOMER HTML ───────────────────────────────────────────────

    def _export_html(self):
        if not self.devices:
            messagebox.showwarning("Export", "No devices loaded.")
            return

        path = filedialog.asksaveasfilename(title="Save Customer Interactive HTML",
            initialdir=os.path.dirname(self.csv_path) if self.csv_path else "",
            initialfile=f"{self.job_var.get()}_SOO.html",
            defaultextension=".html", filetypes=[("HTML Files", "*.html")])
        if not path:
            return

        self.status_var.set("Generating interactive document...")
        self.root.update()

        try:
            html = self._generate_soo_html_content()
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            self.status_var.set(f"Customer file exported: {os.path.basename(path)}")
            if messagebox.askyesno("Success",
                    f"Customer interactive file saved to:\n{path}\n\n"
                    "Self-contained file for USB stick. Opens in any browser.\n\nOpen it now?"):
                os.startfile(path)
        except Exception as e:
            self.status_var.set("Export failed")
            messagebox.showerror("Export Error", str(e))


def _show_splash(root):
    """Shows a brief branded splash screen while the main window builds."""
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg="#232728")

    w, h = 420, 320
    sw = splash.winfo_screenwidth()
    sh = splash.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    splash.geometry(f"{w}x{h}+{x}+{y}")

    logo_img = None
    logo_path = os.path.join(os.path.dirname(__file__), "splash_logo.png")
    if os.path.exists(logo_path):
        try:
            img = Image.open(logo_path)
            img.thumbnail((150, 150), Image.LANCZOS)
            logo_img = ImageTk.PhotoImage(img)
        except Exception:
            logo_img = None

    if logo_img is not None:
        lbl = tk.Label(splash, image=logo_img, bg="#232728")
        lbl.image = logo_img
        lbl.pack(pady=(36, 10))

    tk.Label(splash, text="Defense Fire Protection", bg="#232728", fg="#ffffff",
             font=("Segoe UI", 15, "bold")).pack()
    tk.Label(splash, text="Turnover Package Tool", bg="#232728", fg="#FF7002",
             font=("Segoe UI", 10)).pack(pady=(2, 0))
    tk.Label(splash, text="Loading...", bg="#232728", fg="#888888",
             font=("Segoe UI", 8)).pack(side="bottom", pady=14)

    splash.update()
    return splash


def _check_for_update_async(root):
    import threading

    def worker():
        info = check_for_update()
        if info:
            root.after(0, lambda: _show_update_dialog(info))

    threading.Thread(target=worker, daemon=True).start()


def _show_update_dialog(info):
    if messagebox.askyesno("Update Available",
            f"A new version is available: {info['version']} (you have {APP_VERSION}).\n\n"
            "Open the download page now?"):
        import webbrowser
        webbrowser.open(info["url"])


def main():
    root = tk.Tk()
    root.iconname(APP_TITLE)
    root.withdraw()

    splash = _show_splash(root)

    app = TurnoverApp(root)

    def _reveal():
        splash.destroy()
        root.deiconify()
        root.lift()
        root.focus_force()
        root.after(1500, lambda: _check_for_update_async(root))

    root.after(1200, _reveal)
    root.mainloop()


if __name__ == "__main__":
    main()
