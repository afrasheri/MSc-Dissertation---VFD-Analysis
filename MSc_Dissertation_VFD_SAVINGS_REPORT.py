import io
import os
import posixpath
import re
from pathlib import Path
from xml.sax.saxutils import unescape as xml_unescape
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pandas as pd
import numpy as np


# =====================================================
# CONFIGURATION
# =====================================================

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE_PREFIX = "VFD Savings Report"


VFD_SAMPLE_MINUTES = 3
FAN_VFD_SCALE_THRESHOLD = 110
FAN_VFD_SCALE_FACTOR = 10
VESSEL_INFO_EXCEL = "POWER_AND_VOLTAGE_PER_SW_&_FANS(1).xlsx"
SAVED_FUEL_MT_PER_KWH = 0.220 / 1000
HFO_USD_PER_MT = 490
ULS_MGO_USD_PER_MT = 600
DEFAULT_USD_PER_MT = 550
AVERAGE_SAMPLE_MINUTES = 5
GENERATOR_RUNNING_LOAD_THRESHOLD = 30
DATA_CUTOFF_DATE = pd.Timestamp("2026-04-01")
AE_FOC_REPORTS_VIEW_THRESHOLD = 60
MAIN_KPI_REPORTS_VIEW_FALLBACK_PREFIX = "__MAIN_KPI_REPORTS_VIEW_FALLBACK_CELL__"
MAIN_KPI_METRIC_COLUMNS = [
    "Sea Water Temperature",
    "ME Load",
    "Total AE kWh",
    "Av. AE kWh",
    "AE FOC"
]
MOTOR_CATEGORY_MAX_RUNNING = {
    "SW": 3,
    "FW": 3,
    "FANS": 4
}

# Keep the shared inputs next to this script; the large per-vessel workbooks
# live in its DATA PER VESSEL subdirectory. Environment variables still allow
# these locations to be overridden (for example, with the network share)
# without editing the script.
DEFAULT_DATA_DIR = BASE_DIR
DATA_DIR = Path(
    os.environ.get("VFD_DATA_DIR", str(DEFAULT_DATA_DIR))
).expanduser().resolve()
RAW_DATA_DIR = Path(
    os.environ.get("VFD_RAW_DATA_DIR", str(DATA_DIR / "DATA PER VESSEL"))
).expanduser().resolve()

VOYAGES_FILE_CANDIDATES = [
    "Voyages.xlsx",
    "Voyages(1).xlsx",
    "VOYAGES_ALL.xlsx",
]

REPORTS_FILE_CANDIDATES = [
    "REPORT_VIEWER.xlsx",
    "REPORT_VIEWER(1).xlsx",
    "REPORTS_VIEW_ALL.xlsx",
]

TELEMETRY_CALC_FILE_CANDIDATES = [
    "Consumptions.xlsx",
    "Consumptions(1).xlsx",
    "MIDDAY_REPORTS_ALL.xlsx",
]

VESSEL_INFO_FILE_CANDIDATES = [
    "POWER_AND_VOLTAGE_PER_SW_&_FANS.xlsx",
    "POWER_AND_VOLTAGE_PER_SW_&_FANS(1).xlsx",
    "VESSEL_VFD_POWER_AND_CURRENTS_FW_SW_FANS.xlsx",
    "VESSEL_INFO.xlsx",
]

RAW_WORKBOOK_PATTERNS = [
    "*RAW_DATA*.xlsx",
    "*raw_data*.xlsx",
]

COMMON_VESSEL_COLUMN_CANDIDATES = [
    "VESSEL_NAME",
    "Vessel Name",
    "VESSEL",
    "Vessel",
]


# =====================================================
# VFD CANDIDATE MAPS
# =====================================================

VFD_CANDIDATES = {
    "SW1": {
        "current": ["P8_01_Current", "SW1_OUT_Current"],
        "auto": ["P8_01_AutoMan", "P8_01_Automan", "SW1_Auto"],
        "status": ["P8_01_Status", "SW1_OUT_STATUS_HMI"],
        "load": ["P8_01_Load", "SW1_OUT_FreqFeedBack"],
        "bypass": ["P8_01_Bypass", "SW1_Bypass"]
    },
    "SW2": {
        "current": ["P8_02_Current", "SW2_OUT_Current"],
        "auto": ["P8_02_AutoMan", "P8_02_Automan", "SW2_Auto"],
        "status": ["P8_02_Status", "SW2_OUT_STATUS_HMI"],
        "load": ["P8_02_Load", "SW2_OUT_FreqFeedBack"],
        "bypass": ["P8_02_Bypass", "SW2_Bypass"]
    },
    "SW3": {
        "current": ["P8_03_Current", "SW3_OUT_Current"],
        "auto": ["P8_03_AutoMan", "P8_03_Automan", "SW3_Auto"],
        "status": ["P8_03_Status", "SW3_OUT_STATUS_HMI"],
        "load": ["P8_03_Load", "SW3_OUT_FreqFeedBack"],
        "bypass": ["P8_03_Bypass", "SW3_Bypass"]
    },

    "FW1": {
        "current": ["FW1_OUT_Current"],
        "auto": ["FW1_Auto"],
        "status": ["FW1_OUT_STATUS_HMI"],
        "load": ["FW1_OUT_FreqFeedBack"],
        "bypass": ["FW1_Bypass"]
    },
    "FW2": {
        "current": ["FW2_OUT_Current"],
        "auto": ["FW2_Auto"],
        "status": ["FW2_OUT_STATUS_HMI"],
        "load": ["FW2_OUT_FreqFeedBack"],
        "bypass": ["FW2_Bypass"]
    },
    "FW3": {
        "current": ["FW3_OUT_Current"],
        "auto": ["FW3_Auto"],
        "status": ["FW3_OUT_STATUS_HMI"],
        "load": ["FW3_OUT_FreqFeedBack"],
        "bypass": ["FW3_Bypass"]
    },

    "FAN1": {
        "current": ["Fan_01_Current", "FAN1_OUT_Current"],
        "auto": ["Fan_01_Automan", "Fan_01_AutoMan", "FAN1_Auto"],
        "status": ["Fan_01_Status", "FAN1_OUT_STATUS_HMI"],
        "load": ["Fan_01_Load", "FAN1_OUT_FreqFeedBack"],
        "bypass": ["Fan_01_Bypass", "FAN1_Bypass"]
    },
    "FAN2": {
        "current": ["Fan_02_Current", "FAN2_OUT_Current"],
        "auto": ["Fan_02_Automan", "Fan_02_AutoMan", "FAN2_Auto"],
        "status": ["Fan_02_Status", "FAN2_OUT_STATUS_HMI"],
        "load": ["Fan_02_Load", "FAN2_OUT_FreqFeedBack"],
        "bypass": ["Fan_02_Bypass", "FAN2_Bypass"]
    },
    "FAN3": {
        "current": ["Fan_03_Current", "FAN3_OUT_Current"],
        "auto": ["Fan_03_Automan", "Fan_03_AutoMan", "FAN3_Auto"],
        "status": ["Fan_03_Status", "FAN3_OUT_STATUS_HMI"],
        "load": ["Fan_03_Load", "FAN3_OUT_FreqFeedBack"],
        "bypass": ["Fan_03_Bypass", "FAN3_Bypass"]
    },
    "FAN4": {
        "current": ["Fan_04_Current", "FAN4_OUT_Current"],
        "auto": ["Fan_04_Automan", "Fan_04_AutoMan", "FAN4_Auto"],
        "status": ["Fan_04_Status", "FAN4_OUT_STATUS_HMI"],
        "load": ["Fan_04_Load", "FAN4_OUT_FreqFeedBack"],
        "bypass": ["Fan_04_Bypass", "FAN4_Bypass"]
    }
}


# =====================================================
# AVERAGES COLUMN CANDIDATES
# =====================================================

AVERAGE_CANDIDATES = {
    "FM_ME_IN_FLOW": [
        "FM_ME_IN_FLOW",
        "M/E FUEL OIL FLOW"
    ],
    "FM_DG_IN_FLOW": [
        "FM_DG_IN_FLOW",
        "D/G FUEL OIL FLOW",
        "G/E F.O. FLOW INLET",
        "G/E FUEL OIL FLOW"
    ],
    "SHAFT_PWR": [
        "SHAFT_PWR",
        "SHAFT POWER",
        "SHAFT TORQUE",
        "SHAFT_POWER",
        "ME POWER",
        "ME_POWER",
        "ME_PWR",
        "M/E LOAD",
        "M/E LOAD ESTIMATED"
    ],
    "PWR_G1": [
        "PWR_G1",
        "NO.1 D/G ACTIVE LOAD",
        "DG_1_LOAD",
        "No1 DG POWER",
        "NO.1 G/E ACTIVE LOAD"
    ],
    "PWR_G2": [
        "PWR_G2",
        "NO.2 D/G ACTIVE LOAD",
        "DG_2_LOAD",
        "No2 DG POWER",
        "NO.2 G/E ACTIVE LOAD"
    ],
    "PWR_G3": [
        "PWR_G3",
        "NO.3 D/G ACTIVE LOAD",
        "DG_3_LOAD",
        "No3 DG POWER",
        "NO.3 G/E ACTIVE LOAD"
    ],
    "SEA_WATER_TEMPERATURE": [
        "SEA_WATER_TEMPERATURE",
        "SEA_WATER_TEMP",
        "SEA WATER TEMPERATURE",
        "SEA WATER TEMP",
        "SEA WATER TEMPERATURE C",
        "SEA WATER TEMP C",
        "SEAWATER_TEMPERATURE",
        "SEAWATER_TEMP",
        "SW_TEMPERATURE",
        "SW_TEMP",
        "SW TEMP",
        "SW TEMPERATURE",
        "S.W. TEMP",
        "SEA_WATER_INLET_TEMPERATURE",
        "SEA_WATER_INLET_TEMP",
        "SW_INLET_TEMP",
        "SEA_TEMP"
    ]
}

PID_SET_POINT_CANDIDATES = [
    "PID_set_point",
    "PID_Set_Point",
    "PID_SET_POINT",
    "PID Set Point",
    "PID Setpoint",
    "PID_setpoint",
    "PID_SETPOINT",
    "PID_SP",
    "PID SP",
    "SW_PID_set_point",
    "SW_PID_SET_POINT",
    "TT4_set_point",
    "TT4_Set_point",
    "TT4_SP"
]

PID_SET_POINT_VARIABLE_EXCEL_COLUMNS = [
    "SW_Set_Point Variable",
    "SW Set Point Variable",
    "PID_Set_Point Variable",
    "PID Set Point Variable",
    "PID Setpoint Variable",
    "Set Point Variable",
    "Setpoint Variable"
]

PID_SET_POINT_MISSING_DISPLAY = "-"

Fan_PID_Temp_HI_CANDIDATES = [ 
    "Fan_PID_Temp_HI",
    "TT6_H_SP",
    "TT6_SPH"
] 

Fan_PID_Temp_LOW_CANDIDATES = [
    "Fan_PID_Temp_LOW",
    "TT6_L_SP",
    "TT6_SPL"
]

# =====================================================
# REPORT CONTEXT
# =====================================================

class ReportContext:
    def __init__(self, vessel_name, start_date, end_date, log_callback=None):
        self.vessel_name = str(vessel_name).strip()

        self.global_start_date = pd.Timestamp(start_date)
        self.global_end_date = (
            pd.Timestamp(end_date)
            + pd.Timedelta(days=1)
            - pd.Timedelta(seconds=1)
        )

        self.log_callback = log_callback or print

    def log(self, message):
        self.log_callback(str(message))


# =====================================================
# BASIC HELPERS
# =====================================================

def find_first_existing_column(df, candidates):
    for col in candidates:
        if col in df.columns:
            return col

    normalized_columns = {
        str(existing_col).strip().lower(): existing_col
        for existing_col in df.columns
    }

    for col in candidates:
        matched_col = normalized_columns.get(str(col).strip().lower())
        if matched_col is not None:
            return matched_col

    return None


def build_total_consumption_candidates(equipment):
    equipment_key = str(equipment).strip()

    return [
        f"TotalConsumpt_{equipment_key}",
        f"TotalConsumption_{equipment_key}",
        f"{equipment_key}_TotalConsumpt",
        f"{equipment_key}_TotalConsumption"
    ]


def coerce_binary_like_series(series):
    def normalize_value(value):
        if isinstance(value, (bytes, bytearray)):
            return int.from_bytes(value, byteorder="big", signed=False)

        if isinstance(value, memoryview):
            return int.from_bytes(value.tobytes(), byteorder="big", signed=False)

        if isinstance(value, str):
            text = value.strip().lower()
            if text in {"true", "t", "yes", "y", "on"}:
                return 1
            if text in {"false", "f", "no", "n", "off", ""}:
                return 0

        return value

    return pd.to_numeric(series.map(normalize_value), errors="coerce")


def sanitize_filename_part(text):
    safe_text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(text).strip())
    safe_text = re.sub(r"\s+", " ", safe_text).strip(" .")

    return safe_text or "UNKNOWN VESSEL"


def build_output_paths(vessel_name):
    file_base = f"{OUTPUT_FILE_PREFIX} - {sanitize_filename_part(vessel_name)}"

    return OUTPUT_DIR / f"{file_base}.xlsx"



def normalize_vessel_key(value):
    """Return a case-insensitive alphanumeric key used for vessel matching."""
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def find_vessel_column(df):
    if df is None or df.empty:
        return None

    normalized = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    for candidate in COMMON_VESSEL_COLUMN_CANDIDATES:
        match = normalized.get(candidate.strip().lower())
        if match is not None:
            return match

    return None


def filter_dataframe_by_vessel(df, vessel_name):
    """Filter a shared dataframe by vessel, preserving dedicated sheets without a vessel column."""
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.copy()

    vessel_col = find_vessel_column(df)
    if vessel_col is None:
        return df.copy()

    target_key = normalize_vessel_key(vessel_name)
    keys = df[vessel_col].map(normalize_vessel_key)
    return df.loc[keys == target_key].copy()


def coerce_excel_datetime_series(series):
    """Convert Excel serial dates, strings and existing datetime values to pandas timestamps."""
    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    numeric = pd.to_numeric(series, errors="coerce")
    serial_mask = numeric.between(20000, 80000, inclusive="both")

    if serial_mask.any():
        result.loc[serial_mask] = (
            pd.Timestamp("1899-12-30")
            + pd.to_timedelta(numeric.loc[serial_mask], unit="D")
        )

    other_mask = ~serial_mask
    if other_mask.any():
        result.loc[other_mask] = pd.to_datetime(
            series.loc[other_mask],
            errors="coerce"
        )

    return result.dt.round("s")


def coerce_excel_datetime_value(value):
    converted = coerce_excel_datetime_series(pd.Series([value]))
    return converted.iloc[0]


def build_excel_date_filter(column, start_date, end_date, kind="excel_datetime"):
    start_timestamp = pd.Timestamp(start_date)
    end_timestamp = pd.Timestamp(end_date)

    if kind == "date_only":
        start_value = (
            start_timestamp.year * 10000
            + start_timestamp.month * 100
            + start_timestamp.day
        )
        end_value = (
            end_timestamp.year * 10000
            + end_timestamp.month * 100
            + end_timestamp.day
        )
    else:
        origin = pd.Timestamp("1899-12-30")
        start_value = (start_timestamp - origin).total_seconds() / 86400
        end_value = (end_timestamp - origin).total_seconds() / 86400

    return {
        "column": column,
        "kind": kind,
        "start_value": start_value,
        "end_value": end_value,
        "assume_sorted": True
    }


def resolve_input_excel(candidate_names, keyword_groups=None, required=True):
    """Resolve a common Excel input file by exact candidate name or filename keywords."""
    for name in candidate_names:
        path = DATA_DIR / name
        if path.exists() and path.is_file():
            return path

    keyword_groups = keyword_groups or []
    excel_files = sorted(DATA_DIR.glob("*.xlsx"))

    for keywords in keyword_groups:
        lowered_keywords = [str(k).lower() for k in keywords]
        for path in excel_files:
            lowered_name = path.name.lower()
            if all(keyword in lowered_name for keyword in lowered_keywords):
                return path

    if required:
        raise FileNotFoundError(
            "Could not locate an input Excel file. Checked: "
            + ", ".join(candidate_names)
            + f" in {DATA_DIR}"
        )

    return None


def list_raw_workbooks():
    workbooks = []

    for pattern in RAW_WORKBOOK_PATTERNS:
        workbooks.extend(RAW_DATA_DIR.glob(pattern))

    unique = []
    seen = set()

    for path in sorted(workbooks):
        if path.name.startswith("~$"):
            continue
        resolved = path.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(path)

    return unique


def list_available_vessels():
    """Return vessel names inferred from the available raw Excel workbooks."""
    vessel_numbers = set()

    for workbook in list_raw_workbooks():
        match = re.search(r"VESSEL[_ ]*(\d+)", workbook.stem, flags=re.IGNORECASE)
        if match:
            vessel_numbers.add(int(match.group(1)))

    return [f"Vessel {number}" for number in sorted(vessel_numbers)]


def find_vessel_raw_workbook(vessel_name):
    """Find the dedicated raw workbook for a vessel using exact vessel number or filename matching."""
    workbooks = list_raw_workbooks()
    if not workbooks:
        raise FileNotFoundError(
            f"No vessel raw workbooks matching {RAW_WORKBOOK_PATTERNS} were found in {RAW_DATA_DIR}."
        )

    vessel_text = str(vessel_name).strip()
    vessel_number_match = re.fullmatch(r"vessel\s*0*(\d+)", vessel_text, flags=re.IGNORECASE)

    if vessel_number_match:
        requested_number = int(vessel_number_match.group(1))
        numbered_matches = []

        for path in workbooks:
            match = re.search(r"(?:^|[^a-z0-9])vessel[_\s-]*0*(\d+)(?:[^0-9]|$)", path.stem, flags=re.IGNORECASE)
            if match and int(match.group(1)) == requested_number:
                numbered_matches.append(path)

        if len(numbered_matches) == 1:
            return numbered_matches[0]
        if len(numbered_matches) > 1:
            return max(numbered_matches, key=lambda p: p.stat().st_mtime)

    requested_key = normalize_vessel_key(vessel_text)
    filename_matches = [
        path for path in workbooks
        if requested_key and requested_key in normalize_vessel_key(path.stem)
    ]

    if len(filename_matches) == 1:
        return filename_matches[0]
    if len(filename_matches) > 1:
        return max(filename_matches, key=lambda p: p.stat().st_mtime)

    available = ", ".join(path.name for path in workbooks)
    raise FileNotFoundError(
        f"Could not match vessel '{vessel_name}' to a raw workbook in {RAW_DATA_DIR}. "
        f"Available raw workbooks: {available}"
    )


_XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XLSX_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_XLSX_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XLSX_ROW_RE = re.compile(rb"<row\b[^>]*>(.*?)</row>", re.DOTALL)
_XLSX_CELL_RE = re.compile(rb"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.DOTALL)
_XLSX_CELL_REF_RE = re.compile(rb'\br="([A-Z]+)\d+"')
_XLSX_CELL_TYPE_RE = re.compile(rb'\bt="([^"]+)"')
_XLSX_VALUE_RE = re.compile(rb"<v>(.*?)</v>", re.DOTALL)
_XLSX_TEXT_RE = re.compile(rb"<t(?:\s[^>]*)?>(.*?)</t>", re.DOTALL)
_XLSX_COLUMN_INDEX_CACHE = {}


def _xlsx_column_index(column_letters):
    cached = _XLSX_COLUMN_INDEX_CACHE.get(column_letters)
    if cached is not None:
        return cached

    index = 0
    for char_code in column_letters:
        index = index * 26 + char_code - 64

    zero_based_index = index - 1
    _XLSX_COLUMN_INDEX_CACHE[column_letters] = zero_based_index
    return zero_based_index


def _xlsx_workbook_sheet_map(archive):
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships_root = ET.fromstring(
        archive.read("xl/_rels/workbook.xml.rels")
    )

    relationship_map = {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in relationships_root.findall(
            f"{{{_XLSX_PACKAGE_REL_NS}}}Relationship"
        )
    }

    sheet_map = {}
    sheets_element = workbook_root.find(f"{{{_XLSX_MAIN_NS}}}sheets")

    for sheet in sheets_element:
        relationship_id = sheet.attrib[
            f"{{{_XLSX_DOC_REL_NS}}}id"
        ]
        target = relationship_map[relationship_id]
        worksheet_path = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(f"xl/{target}")
        )
        sheet_map[sheet.attrib["name"]] = worksheet_path

    return sheet_map


def _xlsx_shared_strings(archive):
    shared_strings_path = "xl/sharedStrings.xml"
    if shared_strings_path not in archive.namelist():
        return []

    root = ET.fromstring(archive.read(shared_strings_path))
    shared_strings = []

    for shared_item in root.findall(f"{{{_XLSX_MAIN_NS}}}si"):
        shared_strings.append(
            "".join(
                text_node.text or ""
                for text_node in shared_item.iter(
                    f"{{{_XLSX_MAIN_NS}}}t"
                )
            )
        )

    return shared_strings


def _xlsx_cell_value(attributes, body, shared_strings):
    type_match = _XLSX_CELL_TYPE_RE.search(attributes)
    cell_type = type_match.group(1) if type_match else b""

    if body is None:
        return None

    if cell_type == b"s":
        value_match = _XLSX_VALUE_RE.search(body)
        if value_match is None:
            return None
        try:
            return shared_strings[int(value_match.group(1))]
        except (ValueError, IndexError):
            return value_match.group(1).decode("utf-8", errors="replace")

    if cell_type == b"inlineStr":
        text_parts = _XLSX_TEXT_RE.findall(body)
        if not text_parts:
            return None
        return xml_unescape(
            b"".join(text_parts).decode("utf-8", errors="replace")
        )

    value_match = _XLSX_VALUE_RE.search(body)
    if value_match is None:
        return None

    raw_value = value_match.group(1)

    if cell_type == b"b":
        return raw_value == b"1"

    if cell_type in {b"str", b"e"}:
        return xml_unescape(
            raw_value.decode("utf-8", errors="replace")
        )

    try:
        numeric_text = raw_value.decode("ascii")
        if "." not in numeric_text and "e" not in numeric_text.lower():
            return int(numeric_text)
        return float(numeric_text)
    except (UnicodeDecodeError, ValueError):
        return xml_unescape(
            raw_value.decode("utf-8", errors="replace")
        )


def get_excel_sheet_names(path):
    with ZipFile(path) as archive:
        return list(_xlsx_workbook_sheet_map(archive).keys())


def _date_only_key(value):
    if value is None or pd.isna(value):
        return None

    # Numeric values (including numeric strings) in the valid Excel-date range
    # must be converted with the Excel epoch before pandas sees them. Otherwise,
    # pandas can interpret an integer such as 45804 as nanoseconds from 1970.
    parsed = coerce_excel_datetime_value(value)
    if pd.isna(parsed):
        return None
    return parsed.year * 10000 + parsed.month * 100 + parsed.day


def _excel_datetime_numeric(value):
    if value is None or pd.isna(value):
        return None

    parsed = coerce_excel_datetime_value(value)
    if pd.isna(parsed):
        return None

    return (
        parsed - pd.Timestamp("1899-12-30")
    ).total_seconds() / 86400


def _row_matches_date_filter(filter_value, date_filter):
    if date_filter is None:
        return True, False

    kind = date_filter.get("kind", "excel_datetime")
    start_value = date_filter["start_value"]
    end_value = date_filter["end_value"]

    if kind == "date_only":
        comparable_value = _date_only_key(filter_value)
    else:
        comparable_value = _excel_datetime_numeric(filter_value)

    if comparable_value is None:
        return False, False

    return (
        start_value <= comparable_value <= end_value,
        comparable_value > end_value
    )


def read_excel_sheet(path, sheet_name, ctx=None, date_filter=None):
    """Read an XLSX sheet directly from its XML for substantially faster large-file loading."""
    path = Path(path)

    if ctx is not None:
        ctx.log(f"Reading {path.name} [{sheet_name}]...")

    with ZipFile(path) as archive:
        sheet_map = _xlsx_workbook_sheet_map(archive)
        worksheet_path = sheet_map.get(sheet_name)

        if worksheet_path is None:
            return pd.DataFrame()

        shared_strings = _xlsx_shared_strings(archive)
        worksheet_xml = archive.read(worksheet_path)

    headers = None
    column_count = 0
    records = []
    filter_column_index = None
    assume_sorted = bool(date_filter and date_filter.get("assume_sorted", True))

    for row_match in _XLSX_ROW_RE.finditer(worksheet_xml):
        row_xml = row_match.group(1)

        if headers is None:
            header_values = {}

            for cell_match in _XLSX_CELL_RE.finditer(row_xml):
                attributes = cell_match.group(1)
                reference_match = _XLSX_CELL_REF_RE.search(attributes)
                if reference_match is None:
                    continue

                column_index = _xlsx_column_index(reference_match.group(1))
                header_values[column_index] = _xlsx_cell_value(
                    attributes,
                    cell_match.group(2),
                    shared_strings
                )

            if not header_values:
                continue

            column_count = max(header_values) + 1
            headers = [
                header_values.get(index)
                for index in range(column_count)
            ]

            if date_filter is not None:
                requested_column = str(date_filter["column"]).strip().lower()
                for index, header in enumerate(headers):
                    if str(header).strip().lower() == requested_column:
                        filter_column_index = index
                        break

                if filter_column_index is None:
                    raise KeyError(
                        f"Date-filter column '{date_filter['column']}' was not found "
                        f"in {path.name} [{sheet_name}]."
                    )

            continue

        if date_filter is not None:
            filter_value = None

            for cell_match in _XLSX_CELL_RE.finditer(row_xml):
                attributes = cell_match.group(1)
                reference_match = _XLSX_CELL_REF_RE.search(attributes)
                if reference_match is None:
                    continue

                column_index = _xlsx_column_index(reference_match.group(1))
                if column_index == filter_column_index:
                    filter_value = _xlsx_cell_value(
                        attributes,
                        cell_match.group(2),
                        shared_strings
                    )
                    break
                if column_index > filter_column_index:
                    break

            row_is_in_range, row_is_after_range = _row_matches_date_filter(
                filter_value,
                date_filter
            )

            if not row_is_in_range:
                if assume_sorted and row_is_after_range:
                    break
                continue

        row_values = [None] * column_count

        for cell_match in _XLSX_CELL_RE.finditer(row_xml):
            attributes = cell_match.group(1)
            reference_match = _XLSX_CELL_REF_RE.search(attributes)
            if reference_match is None:
                continue

            column_index = _xlsx_column_index(reference_match.group(1))
            if column_index >= column_count:
                continue

            row_values[column_index] = _xlsx_cell_value(
                attributes,
                cell_match.group(2),
                shared_strings
            )

        records.append(row_values)

    if headers is None:
        return pd.DataFrame()

    return pd.DataFrame.from_records(records, columns=headers)


def first_existing_sheet(path, candidates):
    sheet_names = get_excel_sheet_names(path)
    normalized = {str(name).strip().lower(): name for name in sheet_names}

    for candidate in candidates:
        match = normalized.get(str(candidate).strip().lower())
        if match is not None:
            return match

    return None


def load_shared_or_embedded_sheet(
    vessel_name,
    common_path,
    common_sheet_candidates,
    raw_workbook,
    embedded_sheet_candidates,
    ctx=None
):
    """Load a vessel-specific subset from a common workbook, then fall back to an embedded sheet."""
    if common_path is not None:
        common_sheet = first_existing_sheet(common_path, common_sheet_candidates)
        if common_sheet is not None:
            common_df = read_excel_sheet(common_path, common_sheet, ctx=ctx)
            vessel_df = filter_dataframe_by_vessel(common_df, vessel_name)
            if not vessel_df.empty:
                return vessel_df, f"{common_path.name}::{common_sheet}"

    embedded_sheet = first_existing_sheet(raw_workbook, embedded_sheet_candidates)
    if embedded_sheet is not None:
        embedded_df = read_excel_sheet(raw_workbook, embedded_sheet, ctx=ctx)
        vessel_df = filter_dataframe_by_vessel(embedded_df, vessel_name)
        if vessel_df.empty and find_vessel_column(embedded_df) is None:
            vessel_df = embedded_df.copy()
        if not vessel_df.empty:
            return vessel_df, f"{raw_workbook.name}::{embedded_sheet}"

    return pd.DataFrame(), None


# =====================================================
# AVERAGE DATA HELPERS
# =====================================================

def resolve_average_columns(df, candidate_map):
    resolved = {}

    for logical_name, candidates in candidate_map.items():
        resolved[logical_name] = find_first_existing_column(df, candidates)

    return resolved


def rename_average_columns_to_standard(df, resolved_map):
    if df.empty:
        return df

    df = df.copy()

    rename_map = {}

    for logical_name, actual_name in resolved_map.items():
        if actual_name is not None and actual_name in df.columns and actual_name != logical_name:
            rename_map[actual_name] = logical_name

    if rename_map:
        df = df.rename(columns=rename_map)

    return df


# =====================================================
# VESSEL INFO EXCEL HELPERS
# =====================================================

def load_vessel_info_excel(file_path, ctx=None):
    try:
        candidate_paths = [
            Path(file_path),
            BASE_DIR / file_path,
            BASE_DIR.parent / file_path
        ]

        resolved_path = next(
            (path for path in candidate_paths if path.exists()),
            Path(file_path)
        )

        sheet_names = get_excel_sheet_names(resolved_path)
        if not sheet_names:
            raise ValueError(f"No worksheets found in {resolved_path}")

        df = read_excel_sheet(resolved_path, sheet_names[0], ctx=ctx)
        df.columns = [str(c).strip() for c in df.columns]
        return df

    except Exception as e:
        if ctx is not None:
            ctx.log(f"❌ Error loading vessel info excel: {e}")
        else:
            print(f"❌ Error loading vessel info excel: {e}")

        return pd.DataFrame()


def get_vessel_info_row(vessel_info_df, vessel_name, ctx=None):
    if vessel_info_df.empty:
        return None

    possible_vessel_cols = [
        "Vessel Name",
        "VESSEL_NAME",
        "Vessel",
        "VESSEL"
    ]

    vessel_col = None

    for c in possible_vessel_cols:
        if c in vessel_info_df.columns:
            vessel_col = c
            break

    if vessel_col is None:
        msg = "❌ Could not find vessel name column in vessel info excel."

        if ctx is not None:
            ctx.log(msg)
        else:
            print(msg)

        return None

    tmp = vessel_info_df.copy()
    tmp[vessel_col] = tmp[vessel_col].astype(str).str.strip().str.upper()

    match = tmp[tmp[vessel_col] == vessel_name.strip().upper()]

    if match.empty:
        msg = f"❌ Vessel {vessel_name} not found in vessel info excel."

        if ctx is not None:
            ctx.log(msg)
        else:
            print(msg)

        return None

    return match.iloc[0]


def get_vessel_voltage(vessel_info_row, ctx=None):
    if vessel_info_row is None:
        return np.nan

    possible_voltage_cols = [
        "Voltage",
        "VOLTAGE",
        "voltage",
        "V",
        "Volt",
        "VOLT",
        "Supply Voltage"
    ]

    for c in possible_voltage_cols:
        if c in vessel_info_row.index:
            return pd.to_numeric(vessel_info_row[c], errors="coerce")

    msg = "❌ Voltage column not found in vessel info excel."

    if ctx is not None:
        ctx.log(msg)
    else:
        print(msg)

    return np.nan


def get_vessel_info_value(vessel_info_row, possible_cols):
    if vessel_info_row is None:
        return ""

    normalized_index = {
        str(col).strip().upper(): col
        for col in vessel_info_row.index
    }

    for col in possible_cols:
        actual_col = normalized_index.get(str(col).strip().upper())

        if actual_col is not None:
            value = vessel_info_row[actual_col]

            if pd.isna(value):
                return ""

            return value

    return ""


def get_vessel_pid_set_point_variable(vessel_info_row):
    value = get_vessel_info_value(
        vessel_info_row,
        PID_SET_POINT_VARIABLE_EXCEL_COLUMNS
    )

    if value == "" or pd.isna(value):
        return ""

    value = str(value).strip()

    if not value or value.strip().lower() in ["nan", "none", "-"]:
        return ""

    return value


def find_configured_pid_set_point_column(dt1, vessel_info_row):
    configured_column = get_vessel_pid_set_point_variable(vessel_info_row)

    if not configured_column:
        return None, ""

    return find_first_existing_column(dt1, [configured_column]), configured_column


def format_vessel_info_date(value):
    if value == "" or pd.isna(value):
        return ""

    parsed_date = coerce_excel_datetime_value(value)

    if pd.notna(parsed_date):
        return parsed_date.strftime("%Y-%m-%d")

    return str(value)


def format_vessel_info_money(value):
    if value == "" or pd.isna(value):
        return ""

    numeric_value = pd.to_numeric(value, errors="coerce")

    if pd.notna(numeric_value):
        return f"{numeric_value:,.2f}"

    return str(value)


def is_blank_input(value):
    if value is None:
        return True

    if pd.isna(value):
        return True

    return str(value).strip() == ""


def parse_money_value(value):
    if value == "" or pd.isna(value):
        return np.nan

    if isinstance(value, str):
        value = value.replace(",", "").replace("$", "").strip()

    return pd.to_numeric(value, errors="coerce")


def get_vessel_installation_date_value(vessel_info_row):
    value = get_vessel_info_value(
        vessel_info_row,
        [
            "INSTALLATION_DATE",
            "Installation Date",
            "Install Date",
            "VFD Installation Date"
        ]
    )

    return coerce_excel_datetime_value(value)


def get_vessel_installation_date(vessel_info_row):
    parsed_date = get_vessel_installation_date_value(vessel_info_row)

    if pd.notna(parsed_date):
        return parsed_date.strftime("%Y-%m-%d")

    value = get_vessel_info_value(
        vessel_info_row,
        [
            "INSTALLATION_DATE",
            "Installation Date",
            "Install Date",
            "VFD Installation Date"
        ]
    )

    return format_vessel_info_date(value)


def get_vessel_retrofit_cost_value(vessel_info_row):
    return parse_money_value(
        get_vessel_info_value(
            vessel_info_row,
            [
                "VFD Retrofit Cost",
                "VFD_RETROFIT_COST",
                "TOTAL_COST",
                "Total Cost",
                "VFD Total Cost",
                "Installation Cost"
            ]
        )
    )


def get_vessel_retrofit_cost(vessel_info_row):
    return format_vessel_info_money(get_vessel_retrofit_cost_value(vessel_info_row))


def get_vessel_total_cost(vessel_info_row):
    value = get_vessel_info_value(
        vessel_info_row,
        [
            "TOTAL_COST",
            "Total Cost",
            "VFD Total Cost",
            "Installation Cost"
        ]
    )

    return format_vessel_info_money(value)


def resolve_report_date_inputs(start_date, end_date, vessel_info_row):
    if is_blank_input(start_date):
        resolved_start = get_vessel_installation_date_value(vessel_info_row)

        if pd.isna(resolved_start):
            raise ValueError(
                "Start Date was not provided and no valid Installation Date "
                "was found for this vessel."
            )
    else:
        resolved_start = pd.to_datetime(start_date, errors="coerce")

        if pd.isna(resolved_start):
            raise ValueError("Invalid Start Date. Please use YYYY-MM-DD.")

    if is_blank_input(end_date):
        resolved_end = DATA_CUTOFF_DATE
    else:
        resolved_end = pd.to_datetime(end_date, errors="coerce")

        if pd.isna(resolved_end):
            raise ValueError("Invalid End Date. Please use YYYY-MM-DD.")

    resolved_start = resolved_start.normalize()
    resolved_end = resolved_end.normalize()

    # All MC Assessment exports contain data only through 2026-04-01.
    # Never count later dates as missing observations.
    resolved_end = min(resolved_end, DATA_CUTOFF_DATE)

    if resolved_end < resolved_start:
        raise ValueError("End Date cannot be earlier than Start Date.")

    return resolved_start, resolved_end


def get_equipment_power_from_excel(vessel_info_row, equipment):
    if vessel_info_row is None:
        return np.nan

    possible_cols = [
        f"{equipment}_POWER",
        f"{equipment}_Power",
        f"{equipment}_power"
    ]

    for col in possible_cols:
        if col in vessel_info_row.index:
            return pd.to_numeric(vessel_info_row[col], errors="coerce")

    return np.nan


def should_include_fw_pumps_table(vessel_info_row):
    for equipment in ["FW1", "FW2", "FW3"]:
        equipment_power = get_equipment_power_from_excel(
            vessel_info_row,
            equipment
        )

        if pd.notna(equipment_power) and equipment_power > 0:
            return True

    return False


# =====================================================
# REPORTS / SEGMENTS HELPERS
# =====================================================

def preprocess_reports(reports_df):
    if reports_df.empty:
        return reports_df

    reports_df = reports_df.copy()
    reports_df = reports_df.sort_values("REPORT_DT_UTC").reset_index(drop=True)

    reports_df["PREV_REPORT_DT_UTC"] = reports_df["REPORT_DT_UTC"].shift(1)

    reports_df["CALCULATED_HOURS"] = (
        reports_df["REPORT_DT_UTC"] - reports_df["PREV_REPORT_DT_UTC"]
    ).dt.total_seconds() / 3600

    reports_df["HOURS_FIXED"] = reports_df["CALCULATED_HOURS"]

    return reports_df


def get_last_available_timestamp(reports_df, avg_df, dt1):
    candidates = []

    if not reports_df.empty and "REPORT_DT_UTC" in reports_df.columns:
        max_rep = pd.to_datetime(
            reports_df["REPORT_DT_UTC"],
            errors="coerce"
        ).max()

        if pd.notna(max_rep):
            candidates.append(max_rep)

    if not avg_df.empty and "DateTimeStamp" in avg_df.columns:
        max_avg = pd.to_datetime(
            avg_df["DateTimeStamp"],
            errors="coerce"
        ).max()

        if pd.notna(max_avg):
            candidates.append(max_avg)

    if not dt1.empty and "DateTimeStamp" in dt1.columns:
        max_vfd = pd.to_datetime(
            dt1["DateTimeStamp"],
            errors="coerce"
        ).max()

        if pd.notna(max_vfd):
            candidates.append(max_vfd)

    if candidates:
        return max(candidates)

    return pd.Timestamp.now()


def build_segments(voyages_df, reports_df=None, avg_df=None, dt1=None, ctx=None):
    segments = []

    if voyages_df.empty:
        return pd.DataFrame()

    voyages_df = voyages_df.copy()
    voyages_df = voyages_df.sort_values("Arrival Date").reset_index(drop=True)

    if ctx is not None:
        fallback_end_date = min(
            get_last_available_timestamp(reports_df, avg_df, dt1),
            ctx.global_end_date
        )
    else:
        fallback_end_date = get_last_available_timestamp(reports_df, avg_df, dt1)

    for i in range(len(voyages_df)):
        row = voyages_df.iloc[i]

        vessel = row["Vessel Name"]
        arrival = row["Arrival Date"]
        departure = row["Departure Date"]
        port_name = row["Port Name"]

        safe_port_name = (
            "UNKNOWN PORT"
            if pd.isna(port_name) or str(port_name).strip() == ""
            else str(port_name).strip()
        )

        # ---------------- PORT STAY ----------------
        if pd.notna(arrival) and pd.notna(departure) and departure >= arrival:
            segments.append({
                "segment_id": f"SEG_{i + 1}_PORT",
                "leg_no": i + 1,
                "segment_type": "Port Stay",
                "start_date": arrival,
                "end_date": departure,
                "from_port": safe_port_name,
                "to_port": safe_port_name,
                "reference_port": safe_port_name,
                "vessel_name": vessel,
                "title": "PORT STAY"
            })

        # ---------------- SEA PASSAGE ----------------
        if i < len(voyages_df) - 1:
            next_row = voyages_df.iloc[i + 1]
            next_arrival = next_row["Arrival Date"]
            next_port = next_row["Port Name"]

            safe_next_port = (
                "UNKNOWN PORT"
                if pd.isna(next_port) or str(next_port).strip() == ""
                else str(next_port).strip()
            )

            if pd.notna(departure) and pd.notna(next_arrival) and next_arrival > departure:
                segments.append({
                    "segment_id": f"SEG_{i + 1}_SEA",
                    "leg_no": i + 1,
                    "segment_type": "Sea Passage",
                    "start_date": departure,
                    "end_date": next_arrival,
                    "from_port": safe_port_name,
                    "to_port": safe_next_port,
                    "reference_port": None,
                    "vessel_name": vessel,
                    "title": "SEA PASSAGE"
                })

        # ---------------- LAST ONGOING SEA PASSAGE ----------------
        else:
            if pd.notna(departure) and pd.notna(fallback_end_date) and fallback_end_date > departure:
                segments.append({
                    "segment_id": f"SEG_{i + 1}_SEA_ONGOING",
                    "leg_no": i + 1,
                    "segment_type": "Sea Passage",
                    "start_date": departure,
                    "end_date": fallback_end_date,
                    "from_port": safe_port_name,
                    "to_port": "ONGOING VOYAGE",
                    "reference_port": None,
                    "vessel_name": vessel,
                    "title": "SEA PASSAGE"
                })

            elif pd.notna(arrival) and pd.isna(departure) and pd.notna(fallback_end_date) and fallback_end_date > arrival:
                segments.append({
                    "segment_id": f"SEG_{i + 1}_PORT_ONGOING",
                    "leg_no": i + 1,
                    "segment_type": "Port Stay",
                    "start_date": arrival,
                    "end_date": fallback_end_date,
                    "from_port": safe_port_name,
                    "to_port": safe_port_name,
                    "reference_port": safe_port_name,
                    "vessel_name": vessel,
                    "title": "PORT STAY"
                })

    segments_df = pd.DataFrame(segments)

    if not segments_df.empty:
        if ctx is not None:
            segments_df = segments_df[
                (segments_df["end_date"] >= ctx.global_start_date) &
                (segments_df["start_date"] <= ctx.global_end_date)
            ].copy()

            segments_df["start_date"] = segments_df["start_date"].apply(
                lambda x: max(x, ctx.global_start_date) if pd.notna(x) else x
            )

            segments_df["end_date"] = segments_df["end_date"].apply(
                lambda x: min(x, ctx.global_end_date) if pd.notna(x) else x
            )

        segments_df["duration_hours"] = (
            segments_df["end_date"] - segments_df["start_date"]
        ).dt.total_seconds() / 3600

        segments_df["duration_days"] = (
            segments_df["end_date"] - segments_df["start_date"]
        ).dt.total_seconds() / 86400

        segments_df = segments_df.sort_values("start_date").reset_index(drop=True)

    return segments_df


def get_reports_for_segment(segment_row, reports_df, ctx=None):
    if reports_df.empty:
        return pd.DataFrame()

    if ctx is not None:
        start_dt = max(segment_row["start_date"], ctx.global_start_date)
        end_dt = min(segment_row["end_date"], ctx.global_end_date)
    else:
        start_dt = segment_row["start_date"]
        end_dt = segment_row["end_date"]

    if segment_row["segment_type"] == "Port Stay":
        allowed_types = ["Port", "Shift",'Drift']
        start_mask = reports_df["REPORT_DT_UTC"] > start_dt

    elif segment_row["segment_type"] == "Sea Passage":
        allowed_types = ["Departure", "Arrival", "Noon"]
        start_mask = reports_df["REPORT_DT_UTC"] >= start_dt

    else:
        allowed_types = []
        start_mask = reports_df["REPORT_DT_UTC"] >= start_dt

    seg_reports = reports_df[
        start_mask &
        (reports_df["REPORT_DT_UTC"] <= end_dt)
    ].copy()

    if seg_reports.empty:
        return seg_reports

    seg_reports = seg_reports[
        seg_reports["REPORT_TYPE"].isin(allowed_types)
    ].copy()

    seg_reports = seg_reports.sort_values("REPORT_DT_UTC").reset_index(drop=True)

    return seg_reports


def apply_segment_intervals(seg_reports, segment_row):
    seg_reports = seg_reports.copy()
    seg_reports = seg_reports.sort_values("REPORT_DT_UTC").reset_index(drop=True)

    segment_start = segment_row["start_date"]
    segment_end = segment_row["end_date"]

    seg_reports["INTERVAL_START"] = seg_reports["REPORT_DT_UTC"].shift(1)
    seg_reports["INTERVAL_END"] = seg_reports["REPORT_DT_UTC"]

    if seg_reports.empty:
        return seg_reports

    seg_reports.loc[0, "INTERVAL_START"] = segment_start

    first_report_dt = seg_reports.loc[0, "REPORT_DT_UTC"]
    first_report_type = str(seg_reports.loc[0, "REPORT_TYPE"]).strip()

    is_first_sea_departure = (
        segment_row["segment_type"] == "Sea Passage" and
        first_report_type == "Departure" and
        pd.notna(segment_start) and
        pd.notna(first_report_dt) and
        first_report_dt == segment_start
    )

    if is_first_sea_departure:
        previous_report_dt = seg_reports.loc[0].get(
            "PREV_REPORT_DT_UTC",
            pd.NaT
        )

        if pd.notna(previous_report_dt):
            seg_reports.loc[0, "INTERVAL_START"] = previous_report_dt

        else:
            first_hours = pd.to_numeric(
                seg_reports.loc[0].get("HOURS_FIXED", np.nan),
                errors="coerce"
            )

            if pd.notna(first_hours) and first_hours > 0:
                seg_reports.loc[0, "INTERVAL_START"] = (
                    first_report_dt - pd.to_timedelta(first_hours, unit="h")
                )

    if pd.notna(segment_start):
        start_index = 1 if is_first_sea_departure else 0

        if start_index < len(seg_reports):
            seg_reports.loc[start_index:, "INTERVAL_START"] = (
                seg_reports.loc[start_index:, "INTERVAL_START"].apply(
                    lambda x: max(x, segment_start) if pd.notna(x) else segment_start
                )
            )

    if pd.notna(segment_end):
        seg_reports["INTERVAL_END"] = seg_reports["INTERVAL_END"].apply(
            lambda x: min(x, segment_end) if pd.notna(x) else x
        )

    seg_reports["INTERVAL_HOURS"] = (
        seg_reports["INTERVAL_END"] - seg_reports["INTERVAL_START"]
    ).dt.total_seconds() / 3600

    seg_reports.loc[
        seg_reports["INTERVAL_HOURS"] < 0,
        "INTERVAL_HOURS"
    ] = np.nan

    return seg_reports


# =====================================================
# TELEMETRY FALLBACK HELPERS
# =====================================================

def get_fallback_main_kpis_from_telemetry(report_start, report_end, telemetry_calc_df, hours_value=None):
    fallback_empty = {
        "Sea Water Temperature": np.nan,
        "ME Load": np.nan,
        "Total AE kWh": np.nan,
        "Av. AE kWh": np.nan,
        "AE FOC": np.nan
    }

    if telemetry_calc_df is None or telemetry_calc_df.empty:
        return fallback_empty.copy()

    if pd.isna(report_end):
        return fallback_empty.copy()

    work_df = telemetry_calc_df.copy()

    if "REPORT_DT_UTC" not in work_df.columns:
        return fallback_empty.copy()

    exact_match = work_df[
        work_df["REPORT_DT_UTC"] == report_end
    ].copy()

    if not exact_match.empty:
        interval_df = exact_match
    else:
        if pd.notna(report_start):
            interval_df = work_df[
                (work_df["REPORT_DT_UTC"] > report_start) &
                (work_df["REPORT_DT_UTC"] <= report_end)
            ].copy()
        else:
            interval_df = work_df[
                work_df["REPORT_DT_UTC"] <= report_end
            ].copy()

    if interval_df.empty:
        return fallback_empty.copy()

    me_load_col = find_first_existing_column(
        interval_df,
        [
            "VPM_ME_LOAD_AVG",
            "VPM_ME_POWER_AVG",
            "VPM_SHAFT_POWER_AVG",
            "VPM_SHAFT_PWR_AVG",
            "SHAFT_PWR",
            "ME_LOAD",
            "ME_POWER"
        ]
    )

    avg_kw_col = find_first_existing_column(
        interval_df,
        [
            "VPM_GE_POWER_AVG",
            "VPM_GE_LOAD_AVG",
            "VPM_DG_POWER_AVG",
            "VPM_DG_LOAD_AVG",
            "GE_POWER_AVG",
            "DG_POWER_AVG"
        ]
    )

    total_kwh_col = find_first_existing_column(
        interval_df,
        [
            "VPM_GE_TOTAL_KWH",
            "VPM_TOTAL_KWH",
            "TOTAL_KWH",
            "GE_TOTAL_KWH",
            "DG_TOTAL_KWH"
        ]
    )

    me_load_value = (
        pd.to_numeric(interval_df[me_load_col], errors="coerce").dropna().mean()
        if me_load_col is not None
        else np.nan
    )

    avg_kw_value = (
        pd.to_numeric(interval_df[avg_kw_col], errors="coerce").dropna().mean()
        if avg_kw_col is not None
        else np.nan
    )

    sea_water_temp_col = find_first_existing_column(
        interval_df,
        AVERAGE_CANDIDATES["SEA_WATER_TEMPERATURE"]
    )

    sea_water_temp_value = (
        pd.to_numeric(
            interval_df[sea_water_temp_col],
            errors="coerce"
        ).dropna().mean()
        if sea_water_temp_col is not None
        else np.nan
    )

    if total_kwh_col is not None:
        total_kwh_value = (
            pd.to_numeric(interval_df[total_kwh_col], errors="coerce")
            .dropna()
            .sum()
        )
    elif pd.notna(avg_kw_value) and pd.notna(hours_value) and hours_value > 0:
        total_kwh_value = avg_kw_value * hours_value
    else:
        total_kwh_value = np.nan

    foc_value = (
        interval_df["VPM_GE_FOC"].dropna().sum()
        if "VPM_GE_FOC" in interval_df.columns
        else np.nan
    )

    return {
        "Sea Water Temperature": sea_water_temp_value,
        "ME Load": me_load_value,
        "Total AE kWh": total_kwh_value,
        "Av. AE kWh": avg_kw_value,
        "AE FOC": foc_value
    }


def get_fallback_foc_from_telemetry(report_start, report_end, telemetry_calc_df):
    fallback_values = get_fallback_main_kpis_from_telemetry(
        report_start=report_start,
        report_end=report_end,
        telemetry_calc_df=telemetry_calc_df
    )

    return {"AE FOC": fallback_values["AE FOC"]}


def is_missing_main_kpi_value(value):
    numeric_value = pd.to_numeric(value, errors="coerce")
    return pd.isna(numeric_value) or numeric_value == 0


def get_main_kpi_reports_view_fallback_marker(col_name):
    return f"{MAIN_KPI_REPORTS_VIEW_FALLBACK_PREFIX}{col_name}"


def is_reports_view_fallback_cell(row, col_name):
    marker_col = get_main_kpi_reports_view_fallback_marker(col_name)
    return bool(row.get(marker_col, False))


def get_reports_view_main_kpis_from_row(report_row, hours_value=None):
    fallback_empty = {
        "Sea Water Temperature": np.nan,
        "ME Load": np.nan,
        "Total AE kWh": np.nan,
        "Av. AE kWh": np.nan,
        "AE FOC": np.nan
    }

    if report_row is None:
        return fallback_empty.copy()

    sea_water_temp_col = None

    for col in AVERAGE_CANDIDATES["SEA_WATER_TEMPERATURE"]:
        if col in report_row.index:
            sea_water_temp_col = col
            break

    sea_water_temp_value = (
        pd.to_numeric(report_row.get(sea_water_temp_col, np.nan), errors="coerce")
        if sea_water_temp_col is not None
        else np.nan
    )

    me_load_value = pd.to_numeric(
        report_row.get("POWER_TM", np.nan),
        errors="coerce"
    )

    foc_cols = ["HFO_GENS", "ULS_MGO_GENS", "LS_HFO_GENS"]
    foc_values = [
        pd.to_numeric(report_row.get(col, np.nan), errors="coerce")
        for col in foc_cols
    ]

    if any(pd.notna(value) for value in foc_values):
        foc_value = sum(value for value in foc_values if pd.notna(value))
    else:
        foc_value = np.nan

    total_kwh_value = 0
    has_kwh_component = False

    for dg_no in [1, 2, 3]:
        kw_value = pd.to_numeric(
            report_row.get(f"DG_{dg_no}_KW_AVERAGE", np.nan),
            errors="coerce"
        )
        minutes_value = pd.to_numeric(
            report_row.get(f"DG_{dg_no}_OPERATION_TOTAL_MINS", np.nan),
            errors="coerce"
        )

        if pd.notna(kw_value) and pd.notna(minutes_value):
            total_kwh_value += kw_value * minutes_value / 60
            has_kwh_component = True

    if not has_kwh_component:
        total_kwh_value = np.nan

    hours_numeric = pd.to_numeric(hours_value, errors="coerce")

    if (
        pd.notna(total_kwh_value) and
        pd.notna(hours_numeric) and
        hours_numeric > 0
    ):
        avg_kw_value = total_kwh_value / hours_numeric
    else:
        avg_kw_value = np.nan

    return {
        "Sea Water Temperature": sea_water_temp_value,
        "ME Load": me_load_value,
        "Total AE kWh": total_kwh_value,
        "Av. AE kWh": avg_kw_value,
        "AE FOC": foc_value
    }


def calculate_number_of_aes_from_operation_hours(total_operation_hours, hours_value=None):
    hours_numeric = pd.to_numeric(hours_value, errors="coerce")
    operation_hours_numeric = pd.to_numeric(total_operation_hours, errors="coerce")

    if (
        pd.isna(hours_numeric) or
        hours_numeric < 0 or
        pd.isna(operation_hours_numeric) or
        operation_hours_numeric <= 0
    ):
        return 0

    if operation_hours_numeric <= hours_numeric + 1:
        return 1

    if operation_hours_numeric <= (2 * hours_numeric) + 1:
        return 2

    return 3


def calculate_reported_dg_operation_hours(report_row):
    if report_row is None:
        return np.nan

    total_operation_minutes = 0
    has_operation_minutes = False

    for dg_no in [1, 2, 3]:
        minutes_value = pd.to_numeric(
            report_row.get(f"DG_{dg_no}_OPERATION_TOTAL_MINS", np.nan),
            errors="coerce"
        )

        if pd.notna(minutes_value):
            total_operation_minutes += minutes_value
            has_operation_minutes = True

    if not has_operation_minutes:
        return np.nan

    return total_operation_minutes / 60


def calculate_number_of_aes_from_row(report_row, hours_value=None):
    total_operation_hours = calculate_reported_dg_operation_hours(report_row)
    return calculate_number_of_aes_from_operation_hours(
        total_operation_hours,
        hours_value=hours_value
    )


def calculate_number_of_aes_from_average_loads(
    avg_df,
    report_start,
    report_end,
    hours_value=None
):
    if (
        avg_df is None or
        avg_df.empty or
        "DateTimeStamp" not in avg_df.columns or
        pd.isna(report_start) or
        pd.isna(report_end) or
        report_end < report_start
    ):
        return 0

    interval_df = avg_df[
        (avg_df["DateTimeStamp"] > report_start) &
        (avg_df["DateTimeStamp"] <= report_end)
    ].copy()

    if interval_df.empty:
        return 0

    total_operation_hours = 0
    has_generator_load_data = False

    for col in ["PWR_G1", "PWR_G2", "PWR_G3"]:
        if col not in interval_df.columns:
            continue

        load_values = pd.to_numeric(interval_df[col], errors="coerce")

        if not load_values.notna().any():
            continue

        has_generator_load_data = True
        running_samples = (load_values > GENERATOR_RUNNING_LOAD_THRESHOLD).sum()
        total_operation_hours += running_samples * AVERAGE_SAMPLE_MINUTES / 60

    if not has_generator_load_data:
        return 0

    return calculate_number_of_aes_from_operation_hours(
        total_operation_hours,
        hours_value=hours_value
    )


def calculate_number_of_aes_for_interval(
    report_row,
    hours_value=None,
    avg_df=None,
    report_start=None,
    report_end=None,
    return_source=False
):
    reported_number_of_aes = calculate_number_of_aes_from_row(
        report_row,
        hours_value=hours_value
    )

    if reported_number_of_aes > 0:
        if return_source:
            return reported_number_of_aes, "REPORTS_VIEW"

        return reported_number_of_aes

    average_number_of_aes = calculate_number_of_aes_from_average_loads(
        avg_df=avg_df,
        report_start=report_start,
        report_end=report_end,
        hours_value=hours_value
    )

    if return_source:
        return (
            average_number_of_aes,
            "AVERAGES" if average_number_of_aes > 0 else ""
        )

    return average_number_of_aes


# =====================================================
# MAIN KPI CALCULATIONS
# =====================================================

def compute_report_interval_metrics(
    report_start,
    report_end,
    avg_df,
    hours_value,
    telemetry_calc_df=None,
    reports_view_row=None
):
    if pd.isna(report_start) or pd.isna(report_end) or report_end < report_start:
        return {
            "Sea Water Temperature": np.nan,
            "ME Load": np.nan,
            "Total AE kWh": np.nan,
            "Av. AE kWh": np.nan,
            "AE FOC": np.nan,
            "__MAIN_KPI_REPORTS_VIEW_FALLBACK_USED": False
        }

    if not avg_df.empty and "DateTimeStamp" in avg_df.columns:
        seg_avg = avg_df[
            (avg_df["DateTimeStamp"] > report_start) &
            (avg_df["DateTimeStamp"] <= report_end)
        ].copy()
    else:
        seg_avg = pd.DataFrame()

    if not seg_avg.empty:
        for col in ["SHAFT_PWR", "PWR_G1", "PWR_G2", "PWR_G3", "FM_DG_IN_FLOW"]:
            if col in seg_avg.columns:
                seg_avg[col] = pd.to_numeric(seg_avg[col], errors="coerce")

        sea_water_temp = (
            pd.to_numeric(
                seg_avg["SEA_WATER_TEMPERATURE"],
                errors="coerce"
            ).dropna().mean()
            if "SEA_WATER_TEMPERATURE" in seg_avg.columns
            else np.nan
        )

        me_load = (
            seg_avg["SHAFT_PWR"].mean()
            if "SHAFT_PWR" in seg_avg.columns
            else np.nan
        )

        if all(col in seg_avg.columns for col in ["PWR_G1", "PWR_G2", "PWR_G3"]):
            seg_avg["TOTAL_KWH_ROW"] = (
                (
                    seg_avg["PWR_G1"].fillna(0) +
                    seg_avg["PWR_G2"].fillna(0) +
                    seg_avg["PWR_G3"].fillna(0)
                ) * 5 / 60
            )

            total_kwh = seg_avg["TOTAL_KWH_ROW"].sum()

        else:
            total_kwh = np.nan

        if pd.notna(total_kwh) and pd.notna(hours_value) and hours_value > 0:
            avg_kw = total_kwh / hours_value
        else:
            avg_kw = np.nan

        foc = np.nan
        if "FM_DG_IN_FLOW" in seg_avg.columns:
            fm_values = pd.to_numeric(
                seg_avg["FM_DG_IN_FLOW"],
                errors="coerce"
            )

            if fm_values.notna().any():
                seg_avg["FOC_ROW"] = (
                    fm_values.fillna(0)
                    * 5 / 60
                    * 0.9297
                    / 1000
                )

                foc = seg_avg["FOC_ROW"].sum()

        if (
            pd.isna(sea_water_temp) or sea_water_temp == 0 or
            pd.isna(me_load) or me_load == 0 or
            pd.isna(total_kwh) or total_kwh == 0 or
            pd.isna(avg_kw) or avg_kw == 0 or
            pd.isna(foc) or foc == 0
        ):
            fallback_values = get_fallback_main_kpis_from_telemetry(
                report_start=report_start,
                report_end=report_end,
                telemetry_calc_df=telemetry_calc_df,
                hours_value=hours_value
            )

            if pd.isna(sea_water_temp) or sea_water_temp == 0:
                sea_water_temp = fallback_values["Sea Water Temperature"]

            if pd.isna(me_load) or me_load == 0:
                me_load = fallback_values["ME Load"]

            if pd.isna(total_kwh) or total_kwh == 0:
                total_kwh = fallback_values["Total AE kWh"]

            if pd.isna(avg_kw) or avg_kw == 0:
                avg_kw = fallback_values["Av. AE kWh"]

            if pd.isna(foc) or foc == 0:
                foc = fallback_values["AE FOC"]

    else:
        sea_water_temp = np.nan
        me_load = np.nan
        total_kwh = np.nan
        avg_kw = np.nan

        fallback_values = get_fallback_main_kpis_from_telemetry(
            report_start=report_start,
            report_end=report_end,
            telemetry_calc_df=telemetry_calc_df,
            hours_value=hours_value
        )

        me_load = fallback_values["ME Load"]
        total_kwh = fallback_values["Total AE kWh"]
        avg_kw = fallback_values["Av. AE kWh"]
        foc = fallback_values["AE FOC"]
        sea_water_temp = fallback_values["Sea Water Temperature"]

    reports_view_fallback_used = False
    reports_view_fallback_columns = set()
    reports_view_fallback_values = get_reports_view_main_kpis_from_row(
        reports_view_row,
        hours_value=hours_value
    )

    if (
        is_missing_main_kpi_value(sea_water_temp) and
        not is_missing_main_kpi_value(
            reports_view_fallback_values["Sea Water Temperature"]
        )
    ):
        sea_water_temp = reports_view_fallback_values["Sea Water Temperature"]
        reports_view_fallback_used = True
        reports_view_fallback_columns.add("Sea Water Temperature")

    if (
        is_missing_main_kpi_value(me_load) and
        not is_missing_main_kpi_value(reports_view_fallback_values["ME Load"])
    ):
        me_load = reports_view_fallback_values["ME Load"]
        reports_view_fallback_used = True
        reports_view_fallback_columns.add("ME Load")

    if (
        is_missing_main_kpi_value(total_kwh) and
        not is_missing_main_kpi_value(reports_view_fallback_values["Total AE kWh"])
    ):
        total_kwh = reports_view_fallback_values["Total AE kWh"]
        reports_view_fallback_used = True
        reports_view_fallback_columns.add("Total AE kWh")

    if (
        is_missing_main_kpi_value(avg_kw) and
        not is_missing_main_kpi_value(reports_view_fallback_values["Av. AE kWh"])
    ):
        avg_kw = reports_view_fallback_values["Av. AE kWh"]
        reports_view_fallback_used = True
        reports_view_fallback_columns.add("Av. AE kWh")

    if (
        is_missing_main_kpi_value(foc) and
        not is_missing_main_kpi_value(reports_view_fallback_values["AE FOC"])
    ):
        foc = reports_view_fallback_values["AE FOC"]
        reports_view_fallback_used = True
        reports_view_fallback_columns.add("AE FOC")

    foc_numeric = pd.to_numeric(foc, errors="coerce")
    reports_view_foc = reports_view_fallback_values["AE FOC"]
    reports_view_foc_available = not is_missing_main_kpi_value(reports_view_foc)
    foc_from_reports_view = "AE FOC" in reports_view_fallback_columns

    if (
        pd.notna(foc_numeric) and
        reports_view_foc_available and
        (
            foc_numeric > AE_FOC_REPORTS_VIEW_THRESHOLD or
            (foc_numeric < 0 and not foc_from_reports_view)
        )
    ):
        foc = reports_view_foc
        reports_view_fallback_used = True
        reports_view_fallback_columns.add("AE FOC")

    metrics = {
        "Sea Water Temperature": (
            round(sea_water_temp, 2)
            if pd.notna(sea_water_temp)
            else 0
        ),
        "ME Load": round(me_load, 2) if pd.notna(me_load) else 0,
        "Total AE kWh": round(total_kwh, 2) if pd.notna(total_kwh) else 0,
        "Av. AE kWh": round(avg_kw, 2) if pd.notna(avg_kw) else 0,
        "AE FOC": round(foc, 2) if pd.notna(foc) else 0,
        "__MAIN_KPI_REPORTS_VIEW_FALLBACK_USED": reports_view_fallback_used
    }

    for col_name in MAIN_KPI_METRIC_COLUMNS:
        metrics[get_main_kpi_reports_view_fallback_marker(col_name)] = (
            col_name in reports_view_fallback_columns
        )

    return metrics


def add_total_row_main_kpi(df):
    if df.empty:
        return df

    df = df.copy()

    numeric_cols = [
        "Hours",
        "Sea Water Temperature",
        "ME Load",
        "Total AE kWh",
        "Av. AE kWh",
        "Number of AEs",
        "AE FOC"
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    total_hours = df["Hours"].sum(skipna=True) if "Hours" in df.columns else 0
    total_kwh = df["Total AE kWh"].sum(skipna=True) if "Total AE kWh" in df.columns else 0
    total_foc = df["AE FOC"].sum(skipna=True) if "AE FOC" in df.columns else 0

    if "Sea Water Temperature" in df.columns and "Hours" in df.columns:
        sea_temp_mask = (
            df["Sea Water Temperature"].notna() &
            df["Hours"].notna() &
            (df["Hours"] > 0) &
            (df["Sea Water Temperature"] > 0)
        )

        if sea_temp_mask.any():
            sea_water_temp_weighted = (
                df.loc[sea_temp_mask, "Sea Water Temperature"] *
                df.loc[sea_temp_mask, "Hours"]
            ).sum() / df.loc[sea_temp_mask, "Hours"].sum()
        else:
            sea_water_temp_weighted = 0
    else:
        sea_water_temp_weighted = 0

    if "ME Load" in df.columns and "Hours" in df.columns:
        me_mask = (
            df["ME Load"].notna() &
            df["Hours"].notna() &
            (df["Hours"] > 0) &
            (df["ME Load"] > 0)
        )

        if me_mask.any():
            me_load_weighted = (
                df.loc[me_mask, "ME Load"] *
                df.loc[me_mask, "Hours"]
            ).sum() / df.loc[me_mask, "Hours"].sum()
        else:
            me_load_weighted = 0
    else:
        me_load_weighted = 0

    if pd.notna(total_hours) and total_hours > 0:
        avg_kw_total = total_kwh / total_hours
    else:
        avg_kw_total = 0

    if "Number of AEs" in df.columns and "Hours" in df.columns:
        ae_mask = (
            df["Number of AEs"].notna() &
            df["Hours"].notna() &
            (df["Hours"] > 0) &
            (df["Number of AEs"] > 0)
        )

        if ae_mask.any():
            number_of_aes_weighted = (
                df.loc[ae_mask, "Number of AEs"] *
                df.loc[ae_mask, "Hours"]
            ).sum() / df.loc[ae_mask, "Hours"].sum()
        else:
            number_of_aes_weighted = 0
    else:
        number_of_aes_weighted = 0

    total_row = {
        "Report": "TOTAL",
        "Date Time": "",
        "Sea Water Temperature": round(sea_water_temp_weighted, 2),
        "Hours": round(total_hours, 2),
        "ME Load": round(me_load_weighted, 2),
        "Total AE kWh": round(total_kwh, 2),
        "Av. AE kWh": round(avg_kw_total, 2),
        "Number of AEs": round(number_of_aes_weighted, 2),
        "AE FOC": round(total_foc, 4)
    }

    for marker_col in [col for col in df.columns if str(col).startswith("__")]:
        total_row[marker_col] = False

    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    return df


def calculate_weighted_sea_water_temperature(report_df):
    if report_df is None or report_df.empty:
        return 0

    if (
        "Sea Water Temperature" not in report_df.columns or
        "Hours" not in report_df.columns
    ):
        return 0

    work_df = report_df.copy()

    if "Report" in work_df.columns:
        report_labels = work_df["Report"].astype(str).str.strip().str.upper()
        work_df = work_df[
            ~report_labels.isin(["TOTAL", "NO REPORTS"])
        ].copy()

    sea_temp = pd.to_numeric(
        work_df["Sea Water Temperature"],
        errors="coerce"
    )
    hours = pd.to_numeric(work_df["Hours"], errors="coerce")

    valid_mask = (
        sea_temp.notna() &
        hours.notna() &
        (sea_temp > 0) &
        (hours > 0)
    )

    if not valid_mask.any():
        return 0

    return (
        (sea_temp[valid_mask] * hours[valid_mask]).sum() /
        hours[valid_mask].sum()
    )


def build_report_level_table(
    segment_row,
    reports_df,
    avg_df,
    telemetry_calc_df=None,
    ctx=None
):
    cache_key = (
        str(segment_row.get("segment_id", "")),
        str(segment_row.get("start_date", "")),
        str(segment_row.get("end_date", ""))
    )
    report_cache = reports_df.attrs.setdefault("_main_report_level_cache", {})
    if cache_key in report_cache:
        return report_cache[cache_key].copy(deep=False)

    seg_reports = get_reports_for_segment(segment_row, reports_df, ctx)

    if seg_reports.empty:
        empty_result = pd.DataFrame([{
            "Report": "No Reports",
            "Date Time": "",
            "Sea Water Temperature": "",
            "Hours": "",
            "ME Load": "",
            "Total AE kWh": "",
            "Av. AE kWh": "",
            "Number of AEs": "",
            "AE FOC": ""
        }])
        report_cache[cache_key] = empty_result
        return empty_result.copy(deep=False)

    seg_reports = apply_segment_intervals(seg_reports, segment_row)

    rows = []

    for _, row in seg_reports.iterrows():
        report_type = row["REPORT_TYPE"]
        report_dt = row["REPORT_DT_UTC"]
        interval_start = row["INTERVAL_START"]
        interval_end = row["INTERVAL_END"]

        hours_value = pd.to_numeric(
            row.get("INTERVAL_HOURS", np.nan),
            errors="coerce"
        )

        number_of_aes, number_of_aes_source = calculate_number_of_aes_for_interval(
            report_row=row,
            hours_value=hours_value,
            avg_df=avg_df,
            report_start=interval_start,
            report_end=interval_end,
            return_source=True
        )

        metrics = compute_report_interval_metrics(
            report_start=interval_start,
            report_end=interval_end,
            avg_df=avg_df,
            hours_value=hours_value,
            telemetry_calc_df=telemetry_calc_df,
            reports_view_row=row
        )

        report_row = {
            "Report": report_type,
            "Date Time": (
                pd.to_datetime(report_dt).strftime("%m/%d/%Y %H:%M")
                if pd.notna(report_dt)
                else ""
            ),
            "Sea Water Temperature": metrics["Sea Water Temperature"],
            "Hours": round(hours_value, 2) if pd.notna(hours_value) else "",
            "ME Load": metrics["ME Load"],
            "Total AE kWh": metrics["Total AE kWh"],
            "Av. AE kWh": metrics["Av. AE kWh"],
            "Number of AEs": number_of_aes,
            "AE FOC": metrics["AE FOC"],
            "__MAIN_KPI_REPORTS_VIEW_FALLBACK_USED": metrics[
                "__MAIN_KPI_REPORTS_VIEW_FALLBACK_USED"
            ]
        }

        for col_name in MAIN_KPI_METRIC_COLUMNS:
            marker_col = get_main_kpi_reports_view_fallback_marker(col_name)
            report_row[marker_col] = metrics.get(marker_col, False)

        report_row[get_main_kpi_reports_view_fallback_marker("Number of AEs")] = (
            number_of_aes_source == "REPORTS_VIEW"
        )

        rows.append(report_row)

    result_df = pd.DataFrame(rows)
    result_df = add_total_row_main_kpi(result_df)
    report_cache[cache_key] = result_df

    return result_df.copy(deep=False)


# =====================================================
# VFD DATE / COLUMN HELPERS
# =====================================================

def build_vfd_datetime_column(dt1, ctx=None):
    if dt1.empty:
        return dt1

    dt1 = dt1.copy()

    if "datetime" in dt1.columns:
        dt1.drop(columns=["datetime"], inplace=True, errors="ignore")

    if "Date" not in dt1.columns or "Time" not in dt1.columns:
        msg = "❌ Δεν βρέθηκαν οι στήλες Date και Time στο VFD table."

        if ctx is not None:
            ctx.log(msg)
            ctx.log(f"Available columns: {list(dt1.columns)}")
        else:
            print(msg)
            print("Available columns:", list(dt1.columns))

        dt1["DateTimeStamp"] = pd.NaT
        return dt1

    parsed_dates = coerce_excel_datetime_series(dt1["Date"])
    date_str = parsed_dates.dt.strftime("%Y-%m-%d")
    time_str = dt1["Time"].astype(str).str.strip()

    am_pm_mask = time_str.str.contains(
        r"\b(?:AM|PM)\b",
        case=False,
        regex=True,
        na=False
    )
    time_str.loc[am_pm_mask] = time_str.loc[am_pm_mask].str.replace(
        r"^0:",
        "12:",
        regex=True
    )

    combined = date_str + " " + time_str

    dt1["DateTimeStamp"] = pd.to_datetime(
        combined,
        format="%Y-%m-%d %I:%M:%S %p",
        errors="coerce"
    )

    mask = dt1["DateTimeStamp"].isna()
    if mask.any():
        dt1.loc[mask, "DateTimeStamp"] = pd.to_datetime(
            combined[mask],
            format="%Y-%m-%d %I:%M %p",
            errors="coerce"
        )

    mask = dt1["DateTimeStamp"].isna()
    if mask.any():
        dt1.loc[mask, "DateTimeStamp"] = pd.to_datetime(
            combined[mask],
            format="%Y-%m-%d %H:%M:%S",
            errors="coerce"
        )

    mask = dt1["DateTimeStamp"].isna()
    if mask.any():
        dt1.loc[mask, "DateTimeStamp"] = pd.to_datetime(
            combined[mask],
            format="%Y-%m-%d %H:%M",
            errors="coerce"
        )

    mask = dt1["DateTimeStamp"].isna()
    if mask.any():
        dt1.loc[mask, "DateTimeStamp"] = pd.to_datetime(
            combined[mask],
            errors="coerce"
        )

    # Excel stores time-only cells as a fraction of one day. Preserve the
    # existing text-time parsing above, then fill numeric Excel-time rows.
    numeric_time = pd.to_numeric(dt1["Time"], errors="coerce")
    numeric_time_mask = (
        dt1["DateTimeStamp"].isna()
        & parsed_dates.notna()
        & numeric_time.ge(0)
        & numeric_time.lt(1)
    )
    if numeric_time_mask.any():
        dt1.loc[numeric_time_mask, "DateTimeStamp"] = (
            parsed_dates.loc[numeric_time_mask].dt.normalize()
            + pd.to_timedelta(
                numeric_time.loc[numeric_time_mask],
                unit="D"
            )
        ).dt.round("s")

    if ctx is not None:
        ctx.log("✅ DateTimeStamp built from Date + Time")
        ctx.log(f"Total rows: {len(dt1)}")
        ctx.log(f"Non-null DateTimeStamp rows: {dt1['DateTimeStamp'].notna().sum()}")
        ctx.log(f"Null DateTimeStamp rows: {dt1['DateTimeStamp'].isna().sum()}")
    else:
        print("\n✅ DateTimeStamp built from Date + Time")
        print("Total rows:", len(dt1))
        print("Non-null DateTimeStamp rows:", dt1["DateTimeStamp"].notna().sum())
        print("Null DateTimeStamp rows:", dt1["DateTimeStamp"].isna().sum())

    return dt1


def resolve_vfd_columns(df, candidate_map):
    resolved = {}

    for equipment, metrics in candidate_map.items():
        resolved[equipment] = {}
        metric_candidates = dict(metrics)
        metric_candidates.setdefault(
            "total_consumption",
            build_total_consumption_candidates(equipment)
        )

        for metric_name, candidates in metric_candidates.items():
            resolved[equipment][metric_name] = find_first_existing_column(
                df,
                candidates
            )

    return resolved


def correct_fan_vfd_load_current_scale(dt1, resolved_map, ctx=None):
    if dt1.empty:
        return dt1

    corrected_by_equipment = {}

    for equipment, metrics in resolved_map.items():
        if not str(equipment).upper().startswith("FAN"):
            continue

        load_col = metrics.get("load")
        current_col = metrics.get("current")

        if load_col is None or load_col not in dt1.columns:
            continue

        dt1[load_col] = pd.to_numeric(dt1[load_col], errors="coerce").astype(float)

        if current_col is not None and current_col in dt1.columns:
            dt1[current_col] = (
                pd.to_numeric(dt1[current_col], errors="coerce").astype(float)
            )

        scale_mask = dt1[load_col] > FAN_VFD_SCALE_THRESHOLD
        corrected_count = int(scale_mask.sum())

        if corrected_count == 0:
            continue

        dt1.loc[scale_mask, load_col] = (
            dt1.loc[scale_mask, load_col] / FAN_VFD_SCALE_FACTOR
        )

        if current_col is not None and current_col in dt1.columns:
            dt1.loc[scale_mask, current_col] = (
                dt1.loc[scale_mask, current_col] / FAN_VFD_SCALE_FACTOR
            )

        corrected_by_equipment[equipment] = corrected_count

    if ctx is not None:
        if corrected_by_equipment:
            details = ", ".join(
                f"{equipment}: {count}"
                for equipment, count in corrected_by_equipment.items()
            )
            ctx.log(
                "Corrected fan VFD load/current scale where load > "
                f"{FAN_VFD_SCALE_THRESHOLD}: {details}"
            )
        else:
            ctx.log(
                "No fan VFD load/current scale correction needed "
                f"(threshold load > {FAN_VFD_SCALE_THRESHOLD})."
            )

    return dt1


def preprocess_vfd_data(dt1, resolved_map, vessel_info_row=None, ctx=None):
    if dt1.empty:
        return dt1

    dt1 = dt1.copy()
    dt1 = build_vfd_datetime_column(dt1, ctx)

    for equipment, metrics in resolved_map.items():
        for metric_name, col_name in metrics.items():
            if col_name is not None and col_name in dt1.columns:
                if metric_name in {"auto", "status", "bypass"}:
                    dt1[col_name] = coerce_binary_like_series(dt1[col_name])
                else:
                    dt1[col_name] = pd.to_numeric(dt1[col_name], errors="coerce")

    dt1 = correct_fan_vfd_load_current_scale(dt1, resolved_map, ctx)

    pid_set_point_col, configured_pid_set_point_col = (
        find_configured_pid_set_point_column(dt1, vessel_info_row)
    )

    if pid_set_point_col is not None:
        dt1[pid_set_point_col] = pd.to_numeric(
            dt1[pid_set_point_col],
            errors="coerce"
        )
        if ctx is not None:
            ctx.log(
                "Using PID set point variable from vessel info Excel: "
                f"{configured_pid_set_point_col} -> {pid_set_point_col}"
            )
    elif ctx is not None:
        if configured_pid_set_point_col:
            ctx.log(
                "PID set point variable from vessel info Excel was not found "
                f"in VFD data: {configured_pid_set_point_col}"
            )
        else:
            ctx.log(
                "No PID set point variable configured in vessel info Excel; "
                "PID_set_point will be shown as -."
            )

    if "DateTimeStamp" in dt1.columns:
        dt1 = dt1.sort_values("DateTimeStamp").reset_index(drop=True)

    return dt1


# =====================================================
# VFD CALCULATION HELPERS
# =====================================================

def get_binary_mask(df, col_name):
    if col_name is None or col_name not in df.columns:
        return pd.Series(False, index=df.index)

    values = coerce_binary_like_series(df[col_name]).fillna(0)

    return values > 0


def get_auto_mask(df, col_name):
    if col_name is None or col_name not in df.columns:
        return pd.Series(False, index=df.index)

    values = coerce_binary_like_series(df[col_name])
    col_upper = str(col_name).strip().upper()

    if "AUTOMAN" in col_upper:
        return values == 0

    return values == 1


def get_running_mask(df, resolved_metrics):
    status_col = resolved_metrics.get("status")

    if status_col is None or status_col not in df.columns:
        return pd.Series(False, index=df.index)

    values = coerce_binary_like_series(df[status_col])

    return values == 1


def calculate_vfd_sample_hours(dt1_interval, sample_minutes=VFD_SAMPLE_MINUTES):
    if (
        dt1_interval is None or
        dt1_interval.empty or
        "DateTimeStamp" not in dt1_interval.columns
    ):
        return 0

    timestamps = pd.to_datetime(
        dt1_interval["DateTimeStamp"],
        errors="coerce"
    ).dropna().drop_duplicates()

    return len(timestamps) * sample_minutes / 60


def get_max_running_motors(category_name, equipment_list=None):
    category_key = str(category_name).strip().upper()

    if category_key in MOTOR_CATEGORY_MAX_RUNNING:
        return MOTOR_CATEGORY_MAX_RUNNING[category_key]

    return len(equipment_list or [])


def calculate_number_of_running_motors(
    total_running_hours,
    report_hours,
    max_running_motors
):
    running_hours_value = pd.to_numeric(total_running_hours, errors="coerce")
    report_hours_value = pd.to_numeric(report_hours, errors="coerce")
    max_running_value = pd.to_numeric(max_running_motors, errors="coerce")

    if (
        pd.isna(running_hours_value) or
        running_hours_value <= 0 or
        pd.isna(report_hours_value) or
        report_hours_value <= 0 or
        pd.isna(max_running_value) or
        max_running_value <= 0
    ):
        return 0

    max_running_int = int(max_running_value)

    for running_motors in range(1, max_running_int):
        if running_hours_value <= (running_motors * report_hours_value) + 1:
            return running_motors

    return max_running_int


def format_running_motors_breakdown(parts):
    labels = []

    for category_name in ["SW", "FW", "FANS"]:
        if category_name not in parts:
            continue

        count = pd.to_numeric(parts.get(category_name), errors="coerce")

        if pd.isna(count):
            continue

        labels.append(f"{int(count)}{category_name}")

    return "-".join(labels)


def _precomputed_energy_column(equipment):
    return f"__{equipment}_VALID_KWH_INCREMENT"


def _calculate_full_series_counter_increments(
    dt1,
    total_consumption_col,
    running_mask,
    equipment_power,
    sample_minutes=3
):
    """Calculate accepted cumulative-counter increments once for the full VFD series."""
    result = pd.Series(np.nan, index=dt1.index, dtype="float64")

    if (
        dt1 is None or
        dt1.empty or
        total_consumption_col is None or
        total_consumption_col not in dt1.columns
    ):
        return result

    counter_all = pd.to_numeric(
        dt1[total_consumption_col],
        errors="coerce"
    ).to_numpy(dtype=float, copy=False)
    running_all = (
        pd.Series(running_mask, index=dt1.index)
        .fillna(False)
        .to_numpy(dtype=bool, copy=False)
    )

    finite_positions = np.flatnonzero(np.isfinite(counter_all))
    if finite_positions.size < 2:
        return result

    counter = counter_all[finite_positions]
    running = running_all[finite_positions]

    # Counter levels are often quantised. Process only level changes and use
    # running samples accumulated since the preceding level change.
    level_starts = np.flatnonzero(np.r_[True, counter[1:] != counter[:-1]])
    if level_starts.size < 2:
        return result

    levels = counter[level_starts]
    deltas = np.diff(levels)
    event_positions = level_starts[1:]
    previous_positions = level_starts[:-1]

    running_cumsum = np.r_[0, np.cumsum(running.astype(np.int64))]
    running_samples_between = (
        running_cumsum[event_positions + 1]
        - running_cumsum[previous_positions]
    )
    running_hours_between = running_samples_between * sample_minutes / 60

    positive = deltas > 0
    valid = positive & (running_samples_between > 0)

    equipment_power_value = pd.to_numeric(equipment_power, errors="coerce")
    if pd.notna(equipment_power_value) and equipment_power_value > 0 and positive.any():
        positive_steps = deltas[positive]
        if positive_steps.size <= 10:
            resolution = float(np.min(positive_steps))
        else:
            q_index = max(0, int(np.ceil(positive_steps.size * 0.10)) - 1)
            resolution = float(np.partition(positive_steps, q_index)[q_index])
        resolution = max(resolution, 0.001)

        max_allowed = (
            float(equipment_power_value) * running_hours_between * 1.35
            + resolution * 1.10
        )
        valid &= deltas <= max_allowed

    # A valid counter column with no accepted increment must not be interpreted
    # as measured zero consumption for an actively running motor.
    if not valid.any():
        return result

    increments = np.zeros(len(dt1), dtype=float)
    target_rows = finite_positions[event_positions[valid]]
    increments[target_rows] = deltas[valid]
    return pd.Series(increments, index=dt1.index, dtype="float64")


def precompute_vfd_counter_increments(
    dt1,
    resolved_map,
    vessel_info_row,
    sample_minutes=3,
    ctx=None
):
    """Precompute motor-energy increments once instead of once per report interval."""
    if dt1 is None or dt1.empty:
        return dt1

    for equipment, metrics in resolved_map.items():
        energy_col = _precomputed_energy_column(equipment)
        bypass_mask = get_binary_mask(dt1, metrics.get("bypass"))
        running_mask = get_running_mask(dt1, metrics) & ~bypass_mask
        equipment_power = get_equipment_power_from_excel(
            vessel_info_row,
            equipment
        )
        dt1[energy_col] = _calculate_full_series_counter_increments(
            dt1=dt1,
            total_consumption_col=metrics.get("total_consumption"),
            running_mask=running_mask,
            equipment_power=equipment_power,
            sample_minutes=sample_minutes
        )
        metrics["precomputed_energy"] = energy_col

    if ctx is not None:
        ctx.log("✅ VFD cumulative-counter increments precomputed once per motor")

    return dt1


def calculate_consumption_from_total_counter(
    dt1_interval,
    total_consumption_col,
    running_mask,
    equipment_power,
    sample_minutes=3
):
    """Fallback for callers without precomputed energy increments."""
    if (
        total_consumption_col is None or
        total_consumption_col not in dt1_interval.columns
    ):
        return np.nan

    running_values = (
        pd.Series(running_mask, index=dt1_interval.index)
        .fillna(False)
        .astype(bool)
    )
    if not running_values.any():
        return 0.0

    counter = pd.to_numeric(
        dt1_interval[total_consumption_col],
        errors="coerce"
    )
    valid = counter.notna()
    if valid.sum() < 2:
        return np.nan

    counter = counter.loc[valid]
    running_values = running_values.loc[valid]
    deltas = counter.diff()
    accepted = deltas[(deltas > 0) & running_values]
    return float(accepted.sum()) if not accepted.empty else np.nan

def compute_vfd_equipment_metrics_for_interval(
    dt1_interval,
    resolved_map,
    voltage,
    vessel_info_row,
    sample_minutes=3
):
    result = {}

    if dt1_interval.empty:
        for equipment in resolved_map.keys():
            result[f"{equipment}_RH"] = 0
            result[f"{equipment}_Auto"] = 0
            result[f"{equipment}_ByPass"] = 0
            result[f"{equipment}_Hz %"] = 0
            result[f"{equipment}_kWh"] = 0
            result[f"{equipment}_MAX"] = 0
            result[f"{equipment}_Save"] = 0

        return result

    hours_per_row = sample_minutes / 60

    for equipment, metrics in resolved_map.items():
        load_col = metrics.get("load")
        total_consumption_col = metrics.get("total_consumption")

        bypass_mask = get_binary_mask(dt1_interval, metrics.get("bypass"))
        running_mask = get_running_mask(dt1_interval, metrics) & ~bypass_mask
        running_samples = running_mask.sum()
        running_hours = running_samples * hours_per_row

        auto_mask = (
            running_mask &
            get_auto_mask(dt1_interval, metrics.get("auto"))
        )

        if load_col is not None and load_col in dt1_interval.columns:
            load_values = pd.to_numeric(
                dt1_interval[load_col],
                errors="coerce"
            )

            avg_load_running = load_values[running_mask].mean()

        else:
            avg_load_running = np.nan

        equipment_power = get_equipment_power_from_excel(
            vessel_info_row,
            equipment
        )

        precomputed_energy_col = metrics.get("precomputed_energy")
        if (
            precomputed_energy_col is not None and
            precomputed_energy_col in dt1_interval.columns
        ):
            if running_samples == 0:
                consumed_kwh = 0.0
            else:
                energy_values = pd.to_numeric(
                    dt1_interval[precomputed_energy_col],
                    errors="coerce"
                )
                consumed_kwh = energy_values.sum(min_count=1)
        else:
            consumed_kwh = calculate_consumption_from_total_counter(
                dt1_interval,
                total_consumption_col,
                running_mask,
                equipment_power,
                sample_minutes=sample_minutes
            )

        consumption_is_valid = pd.notna(consumed_kwh)

        if pd.notna(equipment_power):
            max_kwh_per_row = equipment_power * sample_minutes / 60
            max_kwh = running_samples * max_kwh_per_row
        else:
            max_kwh = np.nan

        # Never convert an unavailable/invalid consumption counter to zero.
        # Otherwise the report creates an artificial saving equal to 100% of
        # the fixed-speed baseline.
        if consumption_is_valid and pd.notna(max_kwh):
            save_kwh = max_kwh - consumed_kwh
        else:
            consumed_kwh = np.nan
            save_kwh = np.nan

        result[f"{equipment}_RH"] = round(
            running_hours,
            2
        )

        result[f"{equipment}_Auto"] = round(
            auto_mask.sum() * hours_per_row,
            2
        )

        result[f"{equipment}_ByPass"] = round(
            bypass_mask.sum() * hours_per_row,
            2
        )

        result[f"{equipment}_Hz %"] = (
            round(avg_load_running, 2)
            if pd.notna(avg_load_running)
            else 0
        )

        result[f"{equipment}_kWh"] = round(consumed_kwh, 2)
        result[f"{equipment}_MAX"] = round(max_kwh, 2)
        result[f"{equipment}_Save"] = round(save_kwh, 2)

    return result


def compute_pid_set_point_for_interval(
    dt1_interval,
    vessel_info_row=None,
    sample_minutes=3
):
    if dt1_interval.empty:
        return PID_SET_POINT_MISSING_DISPLAY, 0

    pid_set_point_col, _ = find_configured_pid_set_point_column(
        dt1_interval,
        vessel_info_row
    )

    if pid_set_point_col is None:
        return PID_SET_POINT_MISSING_DISPLAY, 0

    values = pd.to_numeric(
        dt1_interval[pid_set_point_col],
        errors="coerce"
    )

    valid_values = values.dropna()

    if valid_values.empty:
        return PID_SET_POINT_MISSING_DISPLAY, 0

    weight_per_row = sample_minutes / 60
    total_weight = len(valid_values) * weight_per_row
    weighted_avg = (valid_values * weight_per_row).sum() / total_weight

    return round(weighted_avg, 2), total_weight


def compute_candidate_average_for_interval(
    dt1_interval,
    candidates,
    sample_minutes=3
):
    if dt1_interval.empty:
        return np.nan, 0

    col_name = find_first_existing_column(dt1_interval, candidates)

    if col_name is None:
        return np.nan, 0

    values = pd.to_numeric(
        dt1_interval[col_name],
        errors="coerce"
    )

    valid_values = values.dropna()

    if valid_values.empty:
        return np.nan, 0

    weight_per_row = sample_minutes / 60
    total_weight = len(valid_values) * weight_per_row
    weighted_avg = (valid_values * weight_per_row).sum() / total_weight

    return round(weighted_avg, 2), total_weight


def format_er_range_temp(low_value, high_value):
    low_numeric = pd.to_numeric(low_value, errors="coerce")
    high_numeric = pd.to_numeric(high_value, errors="coerce")

    if pd.isna(low_numeric) or pd.isna(high_numeric):
        return ""

    return f"{low_numeric:.2f} - {high_numeric:.2f}"


def compute_er_range_temp_for_interval(dt1_interval, sample_minutes=3):
    low_value, low_weight = compute_candidate_average_for_interval(
        dt1_interval,
        Fan_PID_Temp_LOW_CANDIDATES,
        sample_minutes=sample_minutes
    )

    high_value, high_weight = compute_candidate_average_for_interval(
        dt1_interval,
        Fan_PID_Temp_HI_CANDIDATES,
        sample_minutes=sample_minutes
    )

    return {
        "display": format_er_range_temp(low_value, high_value),
        "low": low_value,
        "high": high_value,
        "low_weight": low_weight,
        "high_weight": high_weight
    }


def add_category_report_saving_columns(df, equipment_list):
    if df.empty:
        return df

    df = df.copy()

    total_save_list = []
    avg_save_list = []

    for _, row in df.iterrows():
        if str(row.get("Report", "")).strip().upper() == "TOTAL":
            total_save_list.append(np.nan)
            avg_save_list.append(np.nan)
            continue

        total_save = 0
        total_rh = 0

        for eq in equipment_list:
            save_col = f"{eq}_Save"
            rh_col = f"{eq}_RH"

            if save_col in df.columns:
                save_value = pd.to_numeric(
                    row.get(save_col, 0),
                    errors="coerce"
                )

                if pd.notna(save_value):
                    total_save += save_value

            if rh_col in df.columns:
                rh_value = pd.to_numeric(
                    row.get(rh_col, 0),
                    errors="coerce"
                )

                if pd.notna(rh_value):
                    total_rh += rh_value

        avg_save = total_save / total_rh if total_rh > 0 else 0

        total_save_list.append(round(total_save, 2))
        avg_save_list.append(round(avg_save, 2))

    df["Daily Save"] = total_save_list
    df["Avg Daily Save"] = avg_save_list

    return df

# =====================================================
# VFD TOTAL ROW / TABLE BUILDERS
# =====================================================

def add_total_row_vfd(df):
    if df.empty:
        return df

    df = df.copy()
    total_row = {}

    rh_cols = [col for col in df.columns if col.endswith("_RH")]

    if rh_cols:
        row_total_rh = (
            df[rh_cols]
            .apply(pd.to_numeric, errors="coerce")
            .fillna(0)
            .sum(axis=1)
        )
    else:
        row_total_rh = pd.Series(0, index=df.index)

    for col in df.columns:
        if col == "Report":
            total_row[col] = "TOTAL"

        elif col == "Date Time":
            total_row[col] = ""

        elif col == "PID_set_point":
            pid_vals = pd.to_numeric(df[col], errors="coerce")

            if "__PID_SET_POINT_WEIGHT" in df.columns:
                pid_weights = pd.to_numeric(
                    df["__PID_SET_POINT_WEIGHT"],
                    errors="coerce"
                )
            else:
                pid_weights = pd.to_numeric(row_total_rh, errors="coerce")

            mask = (
                pid_vals.notna() &
                pid_weights.notna() &
                (pid_weights > 0)
            )

            if mask.any():
                weighted_pid_set_point = (
                    pid_vals[mask] * pid_weights[mask]
                ).sum() / pid_weights[mask].sum()

                total_row[col] = round(weighted_pid_set_point, 2)
            else:
                total_row[col] = PID_SET_POINT_MISSING_DISPLAY

        elif col == "ER Range Temp.":
            low_vals = pd.to_numeric(
                df.get("__ER_RANGE_TEMP_LOW", pd.Series(np.nan, index=df.index)),
                errors="coerce"
            )
            high_vals = pd.to_numeric(
                df.get("__ER_RANGE_TEMP_HIGH", pd.Series(np.nan, index=df.index)),
                errors="coerce"
            )
            low_weights = pd.to_numeric(
                df.get("__ER_RANGE_TEMP_LOW_WEIGHT", pd.Series(np.nan, index=df.index)),
                errors="coerce"
            )
            high_weights = pd.to_numeric(
                df.get("__ER_RANGE_TEMP_HIGH_WEIGHT", pd.Series(np.nan, index=df.index)),
                errors="coerce"
            )

            low_mask = (
                low_vals.notna() &
                low_weights.notna() &
                (low_weights > 0)
            )
            high_mask = (
                high_vals.notna() &
                high_weights.notna() &
                (high_weights > 0)
            )

            total_low = (
                (low_vals[low_mask] * low_weights[low_mask]).sum() /
                low_weights[low_mask].sum()
                if low_mask.any()
                else np.nan
            )
            total_high = (
                (high_vals[high_mask] * high_weights[high_mask]).sum() /
                high_weights[high_mask].sum()
                if high_mask.any()
                else np.nan
            )

            total_row[col] = format_er_range_temp(total_low, total_high)

        elif col.endswith("Hz %"):
            eq_prefix = col.replace("_Hz %", "")
            rh_col = f"{eq_prefix}_RH"

            if rh_col in df.columns:
                hz_vals = pd.to_numeric(df[col], errors="coerce")
                rh_vals = pd.to_numeric(df[rh_col], errors="coerce")

                mask = (
                    hz_vals.notna() &
                    rh_vals.notna() &
                    (rh_vals > 0)
                )

                if mask.any():
                    weighted_avg_hz = (
                        hz_vals[mask] * rh_vals[mask]
                    ).sum() / rh_vals[mask].sum()

                    total_row[col] = round(weighted_avg_hz, 2)
                else:
                    total_row[col] = 0
            else:
                total_row[col] = 0

        elif col in ["Daily Save", "Avg Daily Save"]:
            avg_vals = pd.to_numeric(df[col], errors="coerce")
            rh_vals = pd.to_numeric(row_total_rh, errors="coerce")

            mask = (
                avg_vals.notna() &
                rh_vals.notna() &
                (rh_vals > 0)
            )

            if mask.any():
                weighted_avg = (
                    avg_vals[mask] * rh_vals[mask]
                ).sum() / rh_vals[mask].sum()

                total_row[col] = round(weighted_avg, 2)
            else:
                total_row[col] = 0

        elif col == "__PID_SET_POINT_WEIGHT":
            numeric_vals = pd.to_numeric(df[col], errors="coerce")
            total_row[col] = (
                numeric_vals.sum()
                if numeric_vals.notna().any()
                else 0
            )

        elif col == "__VFD_SAMPLE_HOURS":
            numeric_vals = pd.to_numeric(df[col], errors="coerce")
            total_row[col] = (
                numeric_vals.sum()
                if numeric_vals.notna().any()
                else 0
            )

        elif col == "__REPORT_INTERVAL_HOURS":
            numeric_vals = pd.to_numeric(df[col], errors="coerce")
            total_row[col] = (
                numeric_vals.sum()
                if numeric_vals.notna().any()
                else 0
            )

        elif str(col).startswith("__"):
            total_row[col] = 0

        else:
            numeric_vals = pd.to_numeric(df[col], errors="coerce")
            total_row[col] = (
                round(numeric_vals.sum(), 2)
                if numeric_vals.notna().any()
                else 0
            )

    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    return df


def slice_sorted_vfd_interval(dt1, interval_start, interval_end):
    """Return a timestamp interval without scanning the complete VFD dataframe."""
    if (
        dt1 is None or
        dt1.empty or
        "DateTimeStamp" not in dt1.columns or
        pd.isna(interval_start) or
        pd.isna(interval_end)
    ):
        return dt1.iloc[0:0]

    timestamps = dt1["DateTimeStamp"]
    left = int(timestamps.searchsorted(interval_start, side="right"))
    right = int(timestamps.searchsorted(interval_end, side="right"))
    return dt1.iloc[left:right]


def build_vfd_report_level_tables(
    segment_row,
    reports_df,
    dt1,
    resolved_map,
    voltage,
    vessel_info_row,
    ctx=None
):
    cache_key = (
        str(segment_row.get("segment_id", "")),
        str(segment_row.get("start_date", "")),
        str(segment_row.get("end_date", ""))
    )
    vfd_cache = dt1.attrs.setdefault("_vfd_report_level_cache", {})
    if cache_key in vfd_cache:
        cached = vfd_cache[cache_key]
        return tuple(frame.copy(deep=False) for frame in cached)

    seg_reports = get_reports_for_segment(segment_row, reports_df, ctx)

    if seg_reports.empty or dt1.empty or "DateTimeStamp" not in dt1.columns:
        empty_result = (pd.DataFrame(), pd.DataFrame(), pd.DataFrame())
        vfd_cache[cache_key] = empty_result
        return tuple(frame.copy(deep=False) for frame in empty_result)

    seg_reports = apply_segment_intervals(seg_reports, segment_row)

    sw_rows = []
    fw_rows = []
    fan_rows = []

    SW_EQUIP = ["SW1", "SW2", "SW3"]
    FW_EQUIP = ["FW1", "FW2", "FW3"]
    FAN_EQUIP = ["FAN1", "FAN2", "FAN3", "FAN4"]

    for _, row in seg_reports.iterrows():
        report_type = row["REPORT_TYPE"]
        report_dt = row["REPORT_DT_UTC"]
        interval_start = row["INTERVAL_START"]
        interval_end = row["INTERVAL_END"]
        report_interval_hours = pd.to_numeric(
            row.get("INTERVAL_HOURS", 0),
            errors="coerce"
        )

        if pd.isna(report_interval_hours):
            report_interval_hours = 0

        dt1_interval = slice_sorted_vfd_interval(
            dt1,
            interval_start,
            interval_end
        )
        vfd_sample_hours = calculate_vfd_sample_hours(
            dt1_interval,
            sample_minutes=VFD_SAMPLE_MINUTES
        )

        equipment_metrics = compute_vfd_equipment_metrics_for_interval(
            dt1_interval=dt1_interval,
            resolved_map=resolved_map,
            voltage=voltage,
            vessel_info_row=vessel_info_row,
            sample_minutes=VFD_SAMPLE_MINUTES
        )

        pid_set_point, pid_set_point_weight = compute_pid_set_point_for_interval(
            dt1_interval=dt1_interval,
            vessel_info_row=vessel_info_row,
            sample_minutes=VFD_SAMPLE_MINUTES
        )

        er_range_temp = compute_er_range_temp_for_interval(
            dt1_interval=dt1_interval,
            sample_minutes=VFD_SAMPLE_MINUTES
        )

        base_info = {
            "Report": report_type,
            "Date Time": (
                pd.to_datetime(report_dt).strftime("%m/%d/%Y %H:%M")
                if pd.notna(report_dt)
                else ""
            )
        }

        # ---------------- SW TABLE ----------------
        sw_row = base_info.copy()
        sw_row["PID_set_point"] = pid_set_point
        sw_row["__PID_SET_POINT_WEIGHT"] = pid_set_point_weight
        sw_row["__VFD_SAMPLE_HOURS"] = vfd_sample_hours
        sw_row["__REPORT_INTERVAL_HOURS"] = report_interval_hours

        for eq in SW_EQUIP:
            sw_row[f"{eq}_RH"] = equipment_metrics.get(f"{eq}_RH", 0)
            sw_row[f"{eq}_Auto"] = equipment_metrics.get(f"{eq}_Auto", 0)
            sw_row[f"{eq}_ByPass"] = equipment_metrics.get(f"{eq}_ByPass", 0)
            sw_row[f"{eq}_Hz %"] = equipment_metrics.get(f"{eq}_Hz %", 0)
            sw_row[f"{eq}_kWh"] = equipment_metrics.get(f"{eq}_kWh", 0)
            sw_row[f"{eq}_MAX"] = equipment_metrics.get(f"{eq}_MAX", 0)
            sw_row[f"{eq}_Save"] = equipment_metrics.get(f"{eq}_Save", 0)

        sw_rows.append(sw_row)

        # ---------------- FW TABLE ----------------
        fw_row = base_info.copy()
        fw_row["PID_set_point"] = pid_set_point
        fw_row["__PID_SET_POINT_WEIGHT"] = pid_set_point_weight
        fw_row["__VFD_SAMPLE_HOURS"] = vfd_sample_hours
        fw_row["__REPORT_INTERVAL_HOURS"] = report_interval_hours

        for eq in FW_EQUIP:
            fw_row[f"{eq}_RH"] = equipment_metrics.get(f"{eq}_RH", 0)
            fw_row[f"{eq}_Auto"] = equipment_metrics.get(f"{eq}_Auto", 0)
            fw_row[f"{eq}_ByPass"] = equipment_metrics.get(f"{eq}_ByPass", 0)
            fw_row[f"{eq}_Hz %"] = equipment_metrics.get(f"{eq}_Hz %", 0)
            fw_row[f"{eq}_kWh"] = equipment_metrics.get(f"{eq}_kWh", 0)
            fw_row[f"{eq}_MAX"] = equipment_metrics.get(f"{eq}_MAX", 0)
            fw_row[f"{eq}_Save"] = equipment_metrics.get(f"{eq}_Save", 0)

        fw_rows.append(fw_row)

        # ---------------- FAN TABLE ----------------
        fan_row = base_info.copy()
        fan_row["ER Range Temp."] = er_range_temp["display"]
        fan_row["__ER_RANGE_TEMP_LOW"] = er_range_temp["low"]
        fan_row["__ER_RANGE_TEMP_HIGH"] = er_range_temp["high"]
        fan_row["__ER_RANGE_TEMP_LOW_WEIGHT"] = er_range_temp["low_weight"]
        fan_row["__ER_RANGE_TEMP_HIGH_WEIGHT"] = er_range_temp["high_weight"]
        fan_row["__VFD_SAMPLE_HOURS"] = vfd_sample_hours
        fan_row["__REPORT_INTERVAL_HOURS"] = report_interval_hours

        for eq in FAN_EQUIP:
            fan_row[f"{eq}_RH"] = equipment_metrics.get(f"{eq}_RH", 0)
            fan_row[f"{eq}_Auto"] = equipment_metrics.get(f"{eq}_Auto", 0)
            fan_row[f"{eq}_ByPass"] = equipment_metrics.get(f"{eq}_ByPass", 0)
            fan_row[f"{eq}_Hz %"] = equipment_metrics.get(f"{eq}_Hz %", 0)
            fan_row[f"{eq}_kWh"] = equipment_metrics.get(f"{eq}_kWh", 0)
            fan_row[f"{eq}_MAX"] = equipment_metrics.get(f"{eq}_MAX", 0)
            fan_row[f"{eq}_Save"] = equipment_metrics.get(f"{eq}_Save", 0)

        fan_rows.append(fan_row)

    sw_df = pd.DataFrame(sw_rows)
    fw_df = pd.DataFrame(fw_rows)
    fan_df = pd.DataFrame(fan_rows)

    sw_df = add_category_report_saving_columns(sw_df, ["SW1", "SW2", "SW3"])
    fw_df = add_category_report_saving_columns(fw_df, ["FW1", "FW2", "FW3"])
    fan_df = add_category_report_saving_columns(fan_df, ["FAN1", "FAN2", "FAN3", "FAN4"])

    sw_df = add_total_row_vfd(sw_df)
    fw_df = add_total_row_vfd(fw_df)
    fan_df = add_total_row_vfd(fan_df)

    cached_result = (sw_df, fw_df, fan_df)
    vfd_cache[cache_key] = cached_result
    return tuple(frame.copy(deep=False) for frame in cached_result)


# =====================================================
# SUMMARY TABLES
# =====================================================

def sum_positive_report_column(df, col_name):
    if df is None or df.empty or col_name not in df.columns:
        return 0

    values = pd.to_numeric(df[col_name], errors="coerce").fillna(0)
    values = values.where(values > 0, 0)

    return values.sum()


def build_fuel_mix_from_reports_df(reports_df):
    fuel_mix_empty = {
        "hfo_share": np.nan,
        "uls_mgo_share": np.nan
    }

    if reports_df is None or reports_df.empty:
        return fuel_mix_empty.copy()

    hfo_gens = sum_positive_report_column(reports_df, "HFO_GENS")
    ls_hfo_gens = sum_positive_report_column(reports_df, "LS_HFO_GENS")
    uls_mgo_gens = sum_positive_report_column(reports_df, "ULS_MGO_GENS")

    hfo_group_foc = hfo_gens + ls_hfo_gens
    ae_foc = hfo_group_foc + uls_mgo_gens

    if ae_foc <= 0:
        ae_foc = sum_positive_report_column(reports_df, "FO_GENS_ALL")

    if ae_foc <= 0:
        return fuel_mix_empty.copy()

    return {
        "hfo_share": hfo_group_foc / ae_foc,
        "uls_mgo_share": uls_mgo_gens / ae_foc
    }


def build_fuel_mix_for_segment(segment_row, reports_df, ctx=None):
    seg_reports = get_reports_for_segment(segment_row, reports_df, ctx)
    return build_fuel_mix_from_reports_df(seg_reports)


def calculate_usd_savings_from_fuel_mix(fuel_mt, fuel_mix=None):
    fuel_value = pd.to_numeric(fuel_mt, errors="coerce")

    if pd.isna(fuel_value):
        return 0

    if fuel_mix is None:
        return fuel_value * DEFAULT_USD_PER_MT

    hfo_share = pd.to_numeric(fuel_mix.get("hfo_share", np.nan), errors="coerce")
    uls_mgo_share = pd.to_numeric(
        fuel_mix.get("uls_mgo_share", np.nan),
        errors="coerce"
    )

    if pd.isna(hfo_share) and pd.isna(uls_mgo_share):
        return fuel_value * DEFAULT_USD_PER_MT

    if pd.isna(hfo_share):
        hfo_share = 0

    if pd.isna(uls_mgo_share):
        uls_mgo_share = 0

    return (
        fuel_value * hfo_share * HFO_USD_PER_MT +
        fuel_value * uls_mgo_share * ULS_MGO_USD_PER_MT
    )


def build_motor_category_segment_summary(
    sw_df,
    fw_df,
    fan_df,
    fuel_mix=None,
    include_fw_pumps=True
):
    def summarize_category(df, equipment_list, category_name):
        max_running_motors = get_max_running_motors(category_name, equipment_list)

        if df.empty:
            return {
                "Category": category_name,
                "Total Running Hours": 0,
                "Number of Running Motors": 0,
                "Total Manual Hours": 0,
                "Total Bypass Hours": 0,
                "Weighted Avg Load %": 0,
                "Total AE kWh": 0,
                "Avg kW": 0,
                "Total Saved kWh": 0,
                "Fuel [m/t]": 0,
                "USD Savings": 0,
                "Avg Saved kW": 0,
                "__VFD_SAMPLE_HOURS": 0,
                "__REPORT_INTERVAL_HOURS": 0,
                "__MAX_RUNNING_MOTORS": max_running_motors
            }

        work_df = df.copy()

        if "Report" in work_df.columns:
            work_df = work_df[
                work_df["Report"].astype(str).str.strip().str.upper() != "TOTAL"
            ].copy()

        total_running_hours = 0
        total_auto_hours = 0
        total_bypass_hours = 0
        load_weighted_sum = 0
        total_kwh = 0
        total_saved_kwh = 0
        report_hours = (
            pd.to_numeric(
                work_df.get("__REPORT_INTERVAL_HOURS", pd.Series(0, index=work_df.index)),
                errors="coerce"
            )
            .fillna(0)
            .sum()
        )

        for eq in equipment_list:
            rh_col = f"{eq}_RH"
            auto_col = f"{eq}_Auto"
            bypass_col = f"{eq}_ByPass"
            load_col = f"{eq}_Hz %"
            kwh_col = f"{eq}_kWh"
            save_col = f"{eq}_Save"
            rh_values = pd.Series(0, index=work_df.index)

            if rh_col in work_df.columns:
                rh_values = pd.to_numeric(
                    work_df[rh_col],
                    errors="coerce"
                ).fillna(0)
                total_running_hours += rh_values.sum()

            if auto_col in work_df.columns:
                total_auto_hours += (
                    pd.to_numeric(work_df[auto_col], errors="coerce")
                    .fillna(0)
                    .sum()
                )

            if bypass_col in work_df.columns:
                total_bypass_hours += (
                    pd.to_numeric(work_df[bypass_col], errors="coerce")
                    .fillna(0)
                    .sum()
                )

            if load_col in work_df.columns:
                load_values = pd.to_numeric(
                    work_df[load_col],
                    errors="coerce"
                ).fillna(0)
                load_weighted_sum += (load_values * rh_values).sum()

            if kwh_col in work_df.columns:
                total_kwh += (
                    pd.to_numeric(work_df[kwh_col], errors="coerce")
                    .fillna(0)
                    .sum()
                )

            if save_col in work_df.columns:
                total_saved_kwh += (
                    pd.to_numeric(work_df[save_col], errors="coerce")
                    .fillna(0)
                    .sum()
                )

        avg_kw = total_kwh / total_running_hours if total_running_hours > 0 else 0
        avg_saved_kw = total_saved_kwh / total_running_hours if total_running_hours > 0 else 0
        total_manual_hours = max(total_running_hours - total_auto_hours, 0)
        number_of_running_motors = calculate_number_of_running_motors(
            total_running_hours,
            report_hours,
            max_running_motors
        )
        weighted_avg_load = (
            load_weighted_sum / total_running_hours
            if total_running_hours > 0
            else 0
        )
        fuel_mt = total_saved_kwh * SAVED_FUEL_MT_PER_KWH
        usd_savings = calculate_usd_savings_from_fuel_mix(fuel_mt, fuel_mix)

        return {
            "Category": category_name,
            "Total Running Hours": round(total_running_hours, 2),
            "Number of Running Motors": number_of_running_motors,
            "Total Manual Hours": round(total_manual_hours, 2),
            "Total Bypass Hours": round(total_bypass_hours, 2),
            "Weighted Avg Load %": round(weighted_avg_load, 2),
            "Total AE kWh": round(total_kwh, 2),
            "Avg kW": round(avg_kw, 2),
            "Total Saved kWh": round(total_saved_kwh, 2),
            "Fuel [m/t]": round(fuel_mt, 3),
            "USD Savings": round(usd_savings, 2),
            "Avg Saved kW": round(avg_saved_kw, 2),
            "__VFD_SAMPLE_HOURS": round(
                pd.to_numeric(
                    work_df.get("__VFD_SAMPLE_HOURS", pd.Series(0, index=work_df.index)),
                    errors="coerce"
                ).fillna(0).sum(),
                2
            ),
            "__REPORT_INTERVAL_HOURS": round(report_hours, 2),
            "__MAX_RUNNING_MOTORS": max_running_motors
        }

    rows = [
        summarize_category(sw_df, ["SW1", "SW2", "SW3"], "SW")
    ]

    if include_fw_pumps:
        rows.append(summarize_category(fw_df, ["FW1", "FW2", "FW3"], "FW"))

    rows.append(
        summarize_category(fan_df, ["FAN1", "FAN2", "FAN3", "FAN4"], "FANS")
    )

    summary_df = pd.DataFrame(rows)

    if not summary_df.empty:
        for col in [
            "Total Running Hours",
            "Number of Running Motors",
            "Total Manual Hours",
            "Total Bypass Hours",
            "Weighted Avg Load %",
            "Total AE kWh",
            "Avg kW",
            "Total Saved kWh",
            "Fuel [m/t]",
            "USD Savings",
            "Avg Saved kW",
            "__VFD_SAMPLE_HOURS",
            "__REPORT_INTERVAL_HOURS",
            "__MAX_RUNNING_MOTORS"
        ]:
            summary_df[col] = pd.to_numeric(summary_df[col], errors="coerce")

        total_running_hours = summary_df["Total Running Hours"].sum()
        total_sample_hours = summary_df["__VFD_SAMPLE_HOURS"].max()
        total_report_hours = summary_df["__REPORT_INTERVAL_HOURS"].max()
        total_max_running_motors = summary_df["__MAX_RUNNING_MOTORS"].sum()
        total_manual_hours = summary_df["Total Manual Hours"].sum()
        total_bypass_hours = summary_df["Total Bypass Hours"].sum()
        total_kwh = summary_df["Total AE kWh"].sum()
        total_saved_kwh = summary_df["Total Saved kWh"].sum()
        total_fuel_mt = total_saved_kwh * SAVED_FUEL_MT_PER_KWH
        total_usd_savings = summary_df["USD Savings"].sum()

        if total_running_hours > 0:
            avg_kw_total = (
                summary_df["Avg kW"] *
                summary_df["Total Running Hours"]
            ).sum() / total_running_hours

            avg_saved_kw_total = (
                summary_df["Avg Saved kW"] *
                summary_df["Total Running Hours"]
            ).sum() / total_running_hours

            weighted_avg_load_total = (
                summary_df["Weighted Avg Load %"] *
                summary_df["Total Running Hours"]
            ).sum() / total_running_hours
        else:
            avg_kw_total = 0
            avg_saved_kw_total = 0
            weighted_avg_load_total = 0

        number_of_running_motors_total = calculate_number_of_running_motors(
            total_running_hours,
            total_report_hours,
            total_max_running_motors
        )

        total_row = {
            "Category": "TOTAL",
            "Total Running Hours": round(total_running_hours, 2),
            "Number of Running Motors": number_of_running_motors_total,
            "Total Manual Hours": round(total_manual_hours, 2),
            "Total Bypass Hours": round(total_bypass_hours, 2),
            "Weighted Avg Load %": round(weighted_avg_load_total, 2),
            "Total AE kWh": round(total_kwh, 2),
            "Avg kW": round(avg_kw_total, 2),
            "Total Saved kWh": round(total_saved_kwh, 2),
            "Fuel [m/t]": round(total_fuel_mt, 3),
            "USD Savings": round(total_usd_savings, 2),
            "Avg Saved kW": round(avg_saved_kw_total, 2),
            "__VFD_SAMPLE_HOURS": round(total_sample_hours, 2),
            "__REPORT_INTERVAL_HOURS": round(total_report_hours, 2),
            "__MAX_RUNNING_MOTORS": total_max_running_motors
        }

        summary_df = pd.concat(
            [summary_df, pd.DataFrame([total_row])],
            ignore_index=True
        )

    return summary_df


def build_final_motor_summary_by_segment_type(
    segments_df,
    reports_df,
    dt1,
    resolved_map,
    voltage,
    vessel_info_row,
    ctx=None,
    avg_df=None,
    telemetry_calc_df=None
):
    rows = []

    segment_types = ["Port Stay", "Sea Passage"]

    if avg_df is None:
        avg_df = pd.DataFrame()

    sea_temp_by_segment_type = {}
    sea_temp_report_tables = []

    for seg_type in segment_types:
        seg_subset = segments_df[
            segments_df["segment_type"] == seg_type
        ].copy()
        segment_report_tables = []

        for _, segment_row in seg_subset.iterrows():
            report_df = build_report_level_table(
                segment_row=segment_row,
                reports_df=reports_df,
                avg_df=avg_df,
                telemetry_calc_df=telemetry_calc_df,
                ctx=ctx
            )

            if not report_df.empty:
                segment_report_tables.append(report_df)

        if segment_report_tables:
            combined_report_df = pd.concat(
                segment_report_tables,
                ignore_index=True
            )
        else:
            combined_report_df = pd.DataFrame()

        sea_temp_by_segment_type[seg_type] = (
            calculate_weighted_sea_water_temperature(combined_report_df)
        )

        if not combined_report_df.empty:
            sea_temp_report_tables.append(combined_report_df)

    sea_water_temp_total = calculate_weighted_sea_water_temperature(
        pd.concat(sea_temp_report_tables, ignore_index=True)
        if sea_temp_report_tables
        else pd.DataFrame()
    )

    category_map = {
        "SW": ["SW1", "SW2", "SW3"],
        "FANS": ["FAN1", "FAN2", "FAN3", "FAN4"]
    }

    if should_include_fw_pumps_table(vessel_info_row):
        category_map = {
            "SW": ["SW1", "SW2", "SW3"],
            "FW": ["FW1", "FW2", "FW3"],
            "FANS": ["FAN1", "FAN2", "FAN3", "FAN4"]
        }

    for category_name, equip_list in category_map.items():
        max_running_motors = get_max_running_motors(category_name, equip_list)

        for seg_type in segment_types:
            seg_subset = segments_df[
                segments_df["segment_type"] == seg_type
            ].copy()

            total_rh = 0
            total_auto = 0
            total_bypass = 0
            load_weighted_sum = 0
            total_kwh = 0
            total_save = 0
            sample_hours = 0
            report_hours = 0
            fuel_mix_report_tables = []

            for _, segment_row in seg_subset.iterrows():
                seg_reports_for_mix = get_reports_for_segment(
                    segment_row,
                    reports_df,
                    ctx
                )

                if not seg_reports_for_mix.empty:
                    fuel_mix_report_tables.append(seg_reports_for_mix)

                sw_df, fw_df, fan_df = build_vfd_report_level_tables(
                    segment_row=segment_row,
                    reports_df=reports_df,
                    dt1=dt1,
                    resolved_map=resolved_map,
                    voltage=voltage,
                    vessel_info_row=vessel_info_row,
                    ctx=ctx
                )

                if category_name == "SW":
                    work_df = sw_df.copy()
                elif category_name == "FW":
                    work_df = fw_df.copy()
                else:
                    work_df = fan_df.copy()

                if not work_df.empty and "Report" in work_df.columns:
                    work_df = work_df[
                        work_df["Report"].astype(str).str.strip().str.upper() != "TOTAL"
                    ].copy()

                if not work_df.empty:
                    sample_hours += (
                        pd.to_numeric(
                            work_df.get(
                                "__VFD_SAMPLE_HOURS",
                                pd.Series(0, index=work_df.index)
                            ),
                            errors="coerce"
                        )
                        .fillna(0)
                        .sum()
                    )
                    report_hours += (
                        pd.to_numeric(
                            work_df.get(
                                "__REPORT_INTERVAL_HOURS",
                                pd.Series(0, index=work_df.index)
                            ),
                            errors="coerce"
                        )
                        .fillna(0)
                        .sum()
                    )

                for eq in equip_list:
                    rh_col = f"{eq}_RH"
                    auto_col = f"{eq}_Auto"
                    bypass_col = f"{eq}_ByPass"
                    load_col = f"{eq}_Hz %"
                    kwh_col = f"{eq}_kWh"
                    save_col = f"{eq}_Save"
                    rh_values = pd.Series(0, index=work_df.index)

                    if rh_col in work_df.columns:
                        rh_values = pd.to_numeric(
                            work_df[rh_col],
                            errors="coerce"
                        ).fillna(0)
                        total_rh += rh_values.sum()

                    if auto_col in work_df.columns:
                        total_auto += (
                            pd.to_numeric(work_df[auto_col], errors="coerce")
                            .fillna(0)
                            .sum()
                        )

                    if bypass_col in work_df.columns:
                        total_bypass += (
                            pd.to_numeric(work_df[bypass_col], errors="coerce")
                            .fillna(0)
                            .sum()
                        )

                    if load_col in work_df.columns:
                        load_values = pd.to_numeric(
                            work_df[load_col],
                            errors="coerce"
                        ).fillna(0)
                        load_weighted_sum += (load_values * rh_values).sum()

                    if kwh_col in work_df.columns:
                        total_kwh += (
                            pd.to_numeric(work_df[kwh_col], errors="coerce")
                            .fillna(0)
                            .sum()
                        )

                    if save_col in work_df.columns:
                        total_save += (
                            pd.to_numeric(work_df[save_col], errors="coerce")
                            .fillna(0)
                            .sum()
                        )

            avg_consumed_kw = total_kwh / total_rh if total_rh > 0 else 0
            avg_saved_kw = total_save / total_rh if total_rh > 0 else 0
            total_manual = max(total_rh - total_auto, 0)
            number_of_running_motors = calculate_number_of_running_motors(
                total_rh,
                report_hours,
                max_running_motors
            )
            weighted_avg_load = (
                load_weighted_sum / total_rh
                if total_rh > 0
                else 0
            )
            fuel_mt = total_save * SAVED_FUEL_MT_PER_KWH

            if fuel_mix_report_tables:
                fuel_mix_reports_df = pd.concat(
                    fuel_mix_report_tables,
                    ignore_index=True
                )
            else:
                fuel_mix_reports_df = pd.DataFrame()

            fuel_mix = build_fuel_mix_from_reports_df(fuel_mix_reports_df)
            usd_savings = calculate_usd_savings_from_fuel_mix(fuel_mt, fuel_mix)

            rows.append({
                "Motor Category": category_name,
                "Segment Type": seg_type,
                "Total Running Hours": round(total_rh, 2),
                "Number of Running Motors": number_of_running_motors,
                "Total Manual Hours": round(total_manual, 2),
                "Total Bypass Hours": round(total_bypass, 2),
                "Avg Sea Water Temperature": round(
                    sea_temp_by_segment_type.get(seg_type, 0),
                    2
                ),
                "Weighted Avg Load %": round(weighted_avg_load, 2),
                "Total Consumed kWh": round(total_kwh, 2),
                "Total Saved kWh": round(total_save, 2),
                "Fuel [mt]": round(fuel_mt, 3),
                "USD Savings": round(usd_savings, 2),
                "Avg Consumed kW": round(avg_consumed_kw, 2),
                "Avg Saved kW": round(avg_saved_kw, 2),
                "__VFD_SAMPLE_HOURS": round(sample_hours, 2),
                "__REPORT_INTERVAL_HOURS": round(report_hours, 2),
                "__MAX_RUNNING_MOTORS": max_running_motors
            })

    summary_df = pd.DataFrame(rows)

    if not summary_df.empty:
        numeric_cols = [
            "Total Running Hours",
            "Number of Running Motors",
            "Total Manual Hours",
            "Total Bypass Hours",
            "Avg Sea Water Temperature",
            "Weighted Avg Load %",
            "Total Consumed kWh",
            "Total Saved kWh",
            "Fuel [mt]",
            "USD Savings",
            "Avg Consumed kW",
            "Avg Saved kW",
            "__VFD_SAMPLE_HOURS",
            "__REPORT_INTERVAL_HOURS",
            "__MAX_RUNNING_MOTORS"
        ]

        for col in numeric_cols:
            summary_df[col] = pd.to_numeric(summary_df[col], errors="coerce").fillna(0)

        total_running_hours = summary_df["Total Running Hours"].sum()
        sample_hours_by_segment_type = (
            summary_df
            .groupby("Segment Type", dropna=False)["__VFD_SAMPLE_HOURS"]
            .max()
        )
        total_sample_hours = sample_hours_by_segment_type.sum()
        report_hours_by_segment_type = (
            summary_df
            .groupby("Segment Type", dropna=False)["__REPORT_INTERVAL_HOURS"]
            .max()
        )
        total_report_hours = report_hours_by_segment_type.sum()
        total_max_running_motors = (
            summary_df
            .drop_duplicates("Motor Category")["__MAX_RUNNING_MOTORS"]
            .sum()
        )
        total_manual_hours = summary_df["Total Manual Hours"].sum()
        total_bypass_hours = summary_df["Total Bypass Hours"].sum()
        total_consumed_kwh = summary_df["Total Consumed kWh"].sum()
        total_saved_kwh = summary_df["Total Saved kWh"].sum()
        total_fuel_mt = total_saved_kwh * SAVED_FUEL_MT_PER_KWH
        total_usd_savings = summary_df["USD Savings"].sum()
        avg_consumed_kw_total = (
            total_consumed_kwh / total_running_hours
            if total_running_hours > 0
            else 0
        )
        avg_saved_kw_total = (
            total_saved_kwh / total_running_hours
            if total_running_hours > 0
            else 0
        )
        weighted_avg_load_total = (
            (
                summary_df["Weighted Avg Load %"] *
                summary_df["Total Running Hours"]
            ).sum() / total_running_hours
            if total_running_hours > 0
            else 0
        )
        number_of_running_motors_total = calculate_number_of_running_motors(
            total_running_hours,
            total_report_hours,
            total_max_running_motors
        )
        running_motors_breakdown = {}

        for category_name, category_df in summary_df.groupby("Motor Category", sort=False):
            category_name = str(category_name).strip().upper()

            if category_name == "TOTAL":
                continue

            category_running_hours = pd.to_numeric(
                category_df["Total Running Hours"],
                errors="coerce"
            ).fillna(0).sum()
            category_report_hours = pd.to_numeric(
                category_df["__REPORT_INTERVAL_HOURS"],
                errors="coerce"
            ).fillna(0).sum()
            category_max_running_motors = pd.to_numeric(
                category_df["__MAX_RUNNING_MOTORS"],
                errors="coerce"
            ).fillna(0).max()

            running_motors_breakdown[category_name] = calculate_number_of_running_motors(
                category_running_hours,
                category_report_hours,
                category_max_running_motors
            )

        total_row = {
            "Motor Category": "TOTAL",
            "Segment Type": "",
            "Total Running Hours": round(total_running_hours, 2),
            "Number of Running Motors": format_running_motors_breakdown(
                running_motors_breakdown
            ),
            "Total Manual Hours": round(total_manual_hours, 2),
            "Total Bypass Hours": round(total_bypass_hours, 2),
            "Avg Sea Water Temperature": round(sea_water_temp_total, 2),
            "Weighted Avg Load %": round(weighted_avg_load_total, 2),
            "Total Consumed kWh": round(total_consumed_kwh, 2),
            "Total Saved kWh": round(total_saved_kwh, 2),
            "Fuel [mt]": round(total_fuel_mt, 3),
            "USD Savings": round(total_usd_savings, 2),
            "Avg Consumed kW": round(avg_consumed_kw_total, 2),
            "Avg Saved kW": round(avg_saved_kw_total, 2),
            "__VFD_SAMPLE_HOURS": round(total_sample_hours, 2),
            "__REPORT_INTERVAL_HOURS": round(total_report_hours, 2),
            "__MAX_RUNNING_MOTORS": total_max_running_motors
        }

        summary_df = pd.concat(
            [summary_df, pd.DataFrame([total_row])],
            ignore_index=True
        )

    return summary_df


def build_final_main_kpi_summary_by_segment_type(
    segments_df,
    reports_df,
    avg_df,
    telemetry_calc_df=None,
    ctx=None
):
    rows = []

    segment_types = ["Port Stay", "Sea Passage"]

    for seg_type in segment_types:
        seg_subset = segments_df[
            segments_df["segment_type"] == seg_type
        ].copy()

        all_report_tables = []

        for _, segment_row in seg_subset.iterrows():
            report_df = build_report_level_table(
                segment_row=segment_row,
                reports_df=reports_df,
                avg_df=avg_df,
                telemetry_calc_df=telemetry_calc_df,
                ctx=ctx
            )

            if not report_df.empty and "Report" in report_df.columns:
                report_df = report_df[
                    report_df["Report"].astype(str).str.strip().str.upper() != "TOTAL"
                ].copy()

            if not report_df.empty:
                all_report_tables.append(report_df)

        if all_report_tables:
            combined = pd.concat(all_report_tables, ignore_index=True)
        else:
            combined = pd.DataFrame(
                columns=["Hours", "Total AE kWh", "Av. AE kWh", "AE FOC"]
            )

        for col in ["Hours", "Total AE kWh", "Av. AE kWh", "AE FOC"]:
            if col in combined.columns:
                combined[col] = pd.to_numeric(combined[col], errors="coerce")

        total_hours = (
            combined["Hours"].sum(skipna=True)
            if "Hours" in combined.columns
            else 0
        )

        total_kwh = (
            combined["Total AE kWh"].sum(skipna=True)
            if "Total AE kWh" in combined.columns
            else 0
        )

        total_foc = (
            combined["AE FOC"].sum(skipna=True)
            if "AE FOC" in combined.columns
            else 0
        )

        avg_kw_total = total_kwh / total_hours if total_hours > 0 else 0

        rows.append({
            "Segment Type": seg_type,
            "Total Hours": round(total_hours, 2),
            "Total AE kWh": round(total_kwh, 2),
            "Avg kW": round(avg_kw_total, 2),
            "Total AE FOC": round(total_foc, 4)
        })

    return pd.DataFrame(rows)


def add_total_usd_savings_to_main_kpi_summary(main_kpi_df, motor_summary_df):
    if main_kpi_df is None or main_kpi_df.empty:
        return main_kpi_df

    result_df = main_kpi_df.copy()
    result_df["Total USD Savings"] = 0.0
    result_df["Number of Running Motors"] = 0

    if (
        motor_summary_df is None or
        motor_summary_df.empty or
        "Segment Type" not in motor_summary_df.columns or
        "USD Savings" not in motor_summary_df.columns
    ):
        return result_df

    work_df = motor_summary_df.copy()

    if "Motor Category" in work_df.columns:
        work_df = work_df[
            work_df["Motor Category"].astype(str).str.strip().str.upper() != "TOTAL"
        ].copy()

    work_df["USD Savings"] = pd.to_numeric(
        work_df["USD Savings"],
        errors="coerce"
    ).fillna(0)

    usd_by_segment_type = (
        work_df
        .groupby("Segment Type", dropna=False)["USD Savings"]
        .sum()
        .to_dict()
    )

    result_df["Total USD Savings"] = result_df["Segment Type"].map(
        usd_by_segment_type
    ).fillna(0).round(2)

    if (
        "Total Running Hours" in work_df.columns and
        "__REPORT_INTERVAL_HOURS" in work_df.columns
    ):
        running_motors_by_segment_type = {}

        for segment_type, segment_df in work_df.groupby("Segment Type", dropna=False):
            total_running_hours = pd.to_numeric(
                segment_df["Total Running Hours"],
                errors="coerce"
            ).fillna(0).sum()
            report_hours = pd.to_numeric(
                segment_df["__REPORT_INTERVAL_HOURS"],
                errors="coerce"
            ).fillna(0).max()

            if "__MAX_RUNNING_MOTORS" in segment_df.columns:
                max_running_df = (
                    segment_df.drop_duplicates("Motor Category")
                    if "Motor Category" in segment_df.columns
                    else segment_df
                )
                max_running_motors = pd.to_numeric(
                    max_running_df["__MAX_RUNNING_MOTORS"],
                    errors="coerce"
                ).fillna(0).sum()
            else:
                max_running_motors = np.nan

            running_motors_by_segment_type[segment_type] = (
                calculate_number_of_running_motors(
                    total_running_hours,
                    report_hours,
                    max_running_motors
                )
            )

        result_df["Number of Running Motors"] = result_df["Segment Type"].map(
            running_motors_by_segment_type
        ).fillna(0).astype(int)

    return result_df


def get_total_vfd_savings_usd(motor_summary_df):
    if (
        motor_summary_df is None or
        motor_summary_df.empty or
        "USD Savings" not in motor_summary_df.columns
    ):
        return 0

    work_df = motor_summary_df.copy()
    work_df["USD Savings"] = pd.to_numeric(
        work_df["USD Savings"],
        errors="coerce"
    ).fillna(0)

    if "Motor Category" in work_df.columns:
        total_mask = (
            work_df["Motor Category"]
            .astype(str)
            .str.strip()
            .str.upper()
            == "TOTAL"
        )

        if total_mask.any():
            return work_df.loc[total_mask, "USD Savings"].sum()

    return work_df["USD Savings"].sum()


def calculate_roi_percent(vfd_savings_usd, retrofit_cost):
    savings_numeric = pd.to_numeric(vfd_savings_usd, errors="coerce")
    cost_numeric = pd.to_numeric(retrofit_cost, errors="coerce")

    if (
        pd.isna(savings_numeric) or
        pd.isna(cost_numeric) or
        cost_numeric <= 0
    ):
        return np.nan

    return savings_numeric / cost_numeric * 100


def format_percent_value(value):
    numeric_value = pd.to_numeric(value, errors="coerce")

    if pd.isna(numeric_value):
        return ""

    return f"{numeric_value:,.2f}%"


def calculate_expected_vfd_observations(
    start_date,
    end_date,
    sample_minutes=VFD_SAMPLE_MINUTES
):
    start_ts = pd.to_datetime(start_date, errors="coerce")
    end_ts = pd.to_datetime(end_date, errors="coerce")
    sample_minutes_numeric = pd.to_numeric(sample_minutes, errors="coerce")

    if (
        pd.isna(start_ts) or
        pd.isna(end_ts) or
        pd.isna(sample_minutes_numeric) or
        sample_minutes_numeric <= 0 or
        end_ts < start_ts
    ):
        return 0

    end_exclusive = end_ts + pd.Timedelta(seconds=1)
    duration_minutes = (end_exclusive - start_ts).total_seconds() / 60

    if duration_minutes <= 0:
        return 0

    return int(np.ceil(duration_minutes / sample_minutes_numeric))


def count_actual_vfd_observations(dt1, start_date, end_date):
    if dt1 is None or dt1.empty or "DateTimeStamp" not in dt1.columns:
        return 0

    start_ts = pd.to_datetime(start_date, errors="coerce")
    end_ts = pd.to_datetime(end_date, errors="coerce")

    if pd.isna(start_ts) or pd.isna(end_ts) or end_ts < start_ts:
        return 0

    timestamps = pd.to_datetime(dt1["DateTimeStamp"], errors="coerce")
    in_range_timestamps = timestamps[
        timestamps.notna() &
        (timestamps >= start_ts) &
        (timestamps <= end_ts)
    ]

    return int(in_range_timestamps.drop_duplicates().shape[0])


def calculate_vfd_missing_data_stats(
    dt1,
    start_date,
    end_date,
    sample_minutes=VFD_SAMPLE_MINUTES
):
    expected_observations = calculate_expected_vfd_observations(
        start_date,
        end_date,
        sample_minutes=sample_minutes
    )
    actual_observations = count_actual_vfd_observations(
        dt1,
        start_date,
        end_date
    )
    missing_observations = max(expected_observations - actual_observations, 0)
    missing_percent = (
        np.nan
        if expected_observations == 0
        else missing_observations / expected_observations * 100
    )

    return {
        "expected_observations": expected_observations,
        "actual_observations": actual_observations,
        "missing_observations": missing_observations,
        "missing_percent": missing_percent
    }


def build_vfd_segment_summary_tables(sw_df, fw_df, fan_df):
    def summarize_group(df, equipment_list):
        rows = []

        for eq in equipment_list:
            rh_col = f"{eq}_RH"
            auto_col = f"{eq}_Auto"
            bypass_col = f"{eq}_ByPass"
            hz_col = f"{eq}_Hz %"
            kwh_col = f"{eq}_kWh"
            max_col = f"{eq}_MAX"
            save_col = f"{eq}_Save"

            if df.empty or rh_col not in df.columns:
                rows.append({
                    "Equipment": eq,
                    "Total Running Hours": 0,
                    "Total Auto Hours": 0,
                    "Total Bypass Hours": 0,
                    "Avg Hz %": 0,
                    "Total Consumed kWh": 0,
                    "Total Max kWh (100% load)": 0,
                    "Total Saved kWh": 0
                })
                continue

            work_df = df.copy()

            numeric_cols = [
                rh_col,
                auto_col,
                bypass_col,
                hz_col,
                kwh_col,
                max_col,
                save_col
            ]

            for col in numeric_cols:
                if col in work_df.columns:
                    work_df[col] = (
                        pd.to_numeric(work_df[col], errors="coerce")
                        .fillna(0)
                    )

            total_rh = round(work_df[rh_col].sum(), 2) if rh_col in work_df.columns else 0
            total_auto = round(work_df[auto_col].sum(), 2) if auto_col in work_df.columns else 0
            total_bypass = round(work_df[bypass_col].sum(), 2) if bypass_col in work_df.columns else 0
            total_kwh = round(work_df[kwh_col].sum(), 2) if kwh_col in work_df.columns else 0
            total_max = round(work_df[max_col].sum(), 2) if max_col in work_df.columns else 0
            total_save = round(work_df[save_col].sum(), 2) if save_col in work_df.columns else 0

            if hz_col in work_df.columns and rh_col in work_df.columns:
                running_mask = work_df[rh_col] > 0
                avg_hz = work_df.loc[running_mask, hz_col].mean()
                avg_hz = round(avg_hz, 2) if pd.notna(avg_hz) else 0
            else:
                avg_hz = 0

            rows.append({
                "Equipment": eq,
                "Total Running Hours": total_rh,
                "Total Auto Hours": total_auto,
                "Total Bypass Hours": total_bypass,
                "Avg Hz %": avg_hz,
                "Total Consumed kWh": total_kwh,
                "Total Max kWh (100% load)": total_max,
                "Total Saved kWh": total_save
            })

        return pd.DataFrame(rows)

    sw_summary = summarize_group(sw_df, ["SW1", "SW2", "SW3"])
    fw_summary = summarize_group(fw_df, ["FW1", "FW2", "FW3"])
    fan_summary = summarize_group(fan_df, ["FAN1", "FAN2", "FAN3", "FAN4"])

    return sw_summary, fw_summary, fan_summary


# =====================================================
# EXCEL EXPORT
# =====================================================

def build_segment_description(segment_row):
    segment_type = str(segment_row.get("segment_type", "Segment")).strip()
    return segment_type if segment_type in {"Port Stay", "Sea Passage"} else "Segment"


def is_negative_save_cell(column_name, value):
    numeric_value = pd.to_numeric(value, errors="coerce")
    return str(column_name).endswith("_Save") and pd.notna(numeric_value) and numeric_value < 0


def is_high_hz_cell(column_name, value):
    numeric_value = pd.to_numeric(value, errors="coerce")
    return str(column_name).endswith("_Hz %") and pd.notna(numeric_value) and numeric_value > 100


def is_auto_less_than_rh_cell(column_name, row):
    column_name = str(column_name)
    if not column_name.endswith("_Auto"):
        return False

    rh_column = column_name[:-5] + "_RH"
    auto_value = pd.to_numeric(row.get(column_name), errors="coerce")
    rh_value = pd.to_numeric(row.get(rh_column), errors="coerce")
    return pd.notna(auto_value) and pd.notna(rh_value) and auto_value < rh_value


def sanitize_excel_sheet_name(name):
    name = str(name)
    name = re.sub(r'[\[\]\:\*\?\/\\]', "_", name)
    name = name.strip()

    if len(name) == 0:
        name = "Sheet"

    return name[:31]


def remove_total_rows_for_excel_sheet(df):
    if df is None or df.empty:
        return df

    df_without_totals = df.copy()

    for possible_total_col in ["Report", "Category", "Segment Type", "Motor Category"]:
        if possible_total_col in df_without_totals.columns:
            total_mask = (
                df_without_totals[possible_total_col]
                .astype(str)
                .str.strip()
                .str.upper()
                == "TOTAL"
            )
            df_without_totals = df_without_totals[~total_mask].copy()

    return df_without_totals.reset_index(drop=True)


def combine_report_tables_horizontally(main_df, sw_df, fw_df, fans_df):
    """Join the KPI and equipment tables into one row per report."""
    tables = [("", main_df), ("SW", sw_df), ("FW", fw_df), ("FAN", fans_df)]
    combined = None

    for prefix, table in tables:
        if table is None or table.empty:
            continue

        current = remove_total_rows_for_excel_sheet(table).reset_index(drop=True)

        if combined is None:
            combined = current.copy()
            continue

        current = current.drop(columns=["Report", "Date Time"], errors="ignore")
        current = current[
            [column for column in current.columns if not str(column).startswith("__")]
        ]
        current = current.rename(
            columns={column: f"{prefix} - {column}" for column in current.columns}
        )
        combined = pd.concat([combined, current], axis=1)

    return combined if combined is not None else pd.DataFrame()


def create_stacked_excel_sheet(writer, workbook, sheet_name, title, title_format):
    safe_sheet_name = sanitize_excel_sheet_name(sheet_name)
    worksheet = workbook.add_worksheet(safe_sheet_name)
    writer.sheets[safe_sheet_name] = worksheet
    worksheet.hide_gridlines(2)
    worksheet.merge_range(0, 0, 0, 8, title, title_format)
    worksheet.freeze_panes(2, 0)

    return worksheet, 2


def normalize_metadata_rows(metadata):
    if not metadata:
        return []

    if isinstance(metadata, dict):
        return [[(key, value)] for key, value in metadata.items()]

    normalized_rows = []

    for item in metadata:
        if isinstance(item, dict):
            normalized_rows.append(list(item.items()))
        elif (
            isinstance(item, tuple) and
            len(item) == 2 and
            not isinstance(item[0], (list, tuple, dict))
        ):
            normalized_rows.append([item])
        else:
            normalized_rows.append(list(item))

    return normalized_rows


def write_table_to_single_excel_sheet(
    worksheet,
    workbook,
    df,
    startrow,
    title=None,
    metadata=None
):
    title_format = workbook.add_format({
        "bold": True,
        "font_size": 13,
        "font_color": "#C00000"
    })

    metadata_label_format = workbook.add_format({
        "bold": True,
        "bg_color": "#EEF3F7",
        "border": 1
    })

    metadata_value_format = workbook.add_format({
        "border": 1
    })

    header_format = workbook.add_format({
        "bold": True,
        "bg_color": "#DDE6EE",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    cell_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    total_format = workbook.add_format({
        "bold": True,
        "bg_color": "#DBE5F1",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    number_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.00"
    })

    total_number_format = workbook.add_format({
        "bold": True,
        "bg_color": "#DBE5F1",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.00"
    })

    integer_number_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0"
    })

    total_integer_number_format = workbook.add_format({
        "bold": True,
        "bg_color": "#DBE5F1",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0"
    })

    fuel_number_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.###"
    })

    total_fuel_number_format = workbook.add_format({
        "bold": True,
        "bg_color": "#DBE5F1",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.###"
    })

    negative_save_format = workbook.add_format({
        "bg_color": "#FF0000",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    negative_save_number_format = workbook.add_format({
        "bg_color": "#FF0000",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.00"
    })

    total_negative_save_format = workbook.add_format({
        "bold": True,
        "bg_color": "#FF0000",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    total_negative_save_number_format = workbook.add_format({
        "bold": True,
        "bg_color": "#FF0000EF",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.00"
    })

    high_hz_format = workbook.add_format({
        "bg_color": "#F8D7DA",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    high_hz_number_format = workbook.add_format({
        "bg_color": "#F8D7DA",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.00"
    })

    total_high_hz_format = workbook.add_format({
        "bold": True,
        "bg_color": "#F8D7DA",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    total_high_hz_number_format = workbook.add_format({
        "bold": True,
        "bg_color": "#F8D7DA",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.00"
    })

    auto_less_than_rh_format = workbook.add_format({
        "bg_color": "#FDE9D9",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    auto_less_than_rh_number_format = workbook.add_format({
        "bg_color": "#FDE9D9",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.00"
    })

    total_auto_less_than_rh_format = workbook.add_format({
        "bold": True,
        "bg_color": "#FDE9D9",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    total_auto_less_than_rh_number_format = workbook.add_format({
        "bold": True,
        "bg_color": "#FDE9D9",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.00"
    })

    reports_view_fallback_format = workbook.add_format({
        "bg_color": "#FFF2CC",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    reports_view_fallback_number_format = workbook.add_format({
        "bg_color": "#FFF2CC",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.00"
    })

    reports_view_fallback_fuel_number_format = workbook.add_format({
        "bg_color": "#FFF2CC",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "#,##0.###"
    })

    if title:
        title_last_col = 7

        if df is not None and not df.empty:
            visible_title_columns = [
                col for col in df.columns
                if not str(col).startswith("__")
            ]
            title_last_col = max(title_last_col, len(visible_title_columns) - 1)

        worksheet.merge_range(startrow, 0, startrow, title_last_col, title, title_format)
        startrow += 2

    if metadata:
        for metadata_row in normalize_metadata_rows(metadata):
            for pair_index, (key, value) in enumerate(metadata_row):
                col_offset = pair_index * 2
                worksheet.write(startrow, col_offset, key, metadata_label_format)
                worksheet.write(startrow, col_offset + 1, value, metadata_value_format)

            startrow += 1

        startrow += 1

    if df is None or df.empty:
        worksheet.write(startrow, 0, "No data available.", cell_format)
        return startrow + 3

    df_to_write = df.copy()
    visible_columns = [
        col for col in df_to_write.columns
        if not str(col).startswith("__")
    ]

    if not visible_columns:
        worksheet.write(startrow, 0, "No data available.", cell_format)
        return startrow + 3

    for col_num, col_name in enumerate(visible_columns):
        worksheet.write(startrow, col_num, col_name, header_format)

    first_data_row = startrow + 1

    for row_idx, (_, row) in enumerate(df_to_write.iterrows()):
        excel_row = first_data_row + row_idx

        is_total_row = False

        for possible_total_col in ["Report", "Category", "Segment Type", "Motor Category"]:
            if possible_total_col in df_to_write.columns:
                value_for_total_check = row.get(possible_total_col, "")

                if str(value_for_total_check).strip().upper() == "TOTAL":
                    is_total_row = True
                    break

        merge_total_first_two_cols = (
            is_total_row and
            len(visible_columns) >= 2 and
            list(visible_columns[:2]) == ["Motor Category", "Segment Type"]
        )

        for col_num, col_name in enumerate(visible_columns):
            if merge_total_first_two_cols and col_num == 0:
                worksheet.merge_range(
                    excel_row,
                    0,
                    excel_row,
                    1,
                    "TOTAL",
                    total_format
                )
                continue

            if merge_total_first_two_cols and col_num == 1:
                continue

            value = row[col_name]

            if pd.isna(value):
                value = ""

            is_negative_save = is_negative_save_cell(col_name, value)
            is_auto_less_than_rh = is_auto_less_than_rh_cell(col_name, row)
            is_high_hz = is_high_hz_cell(col_name, value)
            reports_view_fallback_cell = (
                not is_total_row and
                is_reports_view_fallback_cell(row, col_name)
            )

            if isinstance(value, pd.Timestamp):
                worksheet.write(
                    excel_row,
                    col_num,
                    value.strftime("%Y-%m-%d %H:%M"),
                    total_format if is_total_row else cell_format
                )

            elif isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
                if is_negative_save:
                    value_format = (
                        total_negative_save_number_format
                        if is_total_row
                        else negative_save_number_format
                    )
                elif col_name in ["Fuel [mt]", "Fuel [m/t]"]:
                    value_format = (
                        total_fuel_number_format
                        if is_total_row
                        else reports_view_fallback_fuel_number_format
                        if reports_view_fallback_cell
                        else fuel_number_format
                    )
                elif col_name == "Number of Running Motors":
                    value_format = (
                        total_integer_number_format
                        if is_total_row
                        else integer_number_format
                    )
                elif is_high_hz:
                    value_format = (
                        total_high_hz_number_format
                        if is_total_row
                        else high_hz_number_format
                    )
                else:
                    if is_auto_less_than_rh:
                        value_format = (
                            total_auto_less_than_rh_number_format
                            if is_total_row
                            else auto_less_than_rh_number_format
                        )
                    elif reports_view_fallback_cell:
                        value_format = reports_view_fallback_number_format
                    else:
                        value_format = (
                            total_number_format
                            if is_total_row
                            else number_format
                        )

                worksheet.write(
                    excel_row,
                    col_num,
                    value,
                    value_format
                )

            else:
                if is_negative_save:
                    value_format = (
                        total_negative_save_format
                        if is_total_row
                        else negative_save_format
                    )
                elif is_high_hz:
                    value_format = (
                        total_high_hz_format
                        if is_total_row
                        else high_hz_format
                    )
                else:
                    if is_auto_less_than_rh:
                        value_format = (
                            total_auto_less_than_rh_format
                            if is_total_row
                            else auto_less_than_rh_format
                        )
                    elif reports_view_fallback_cell:
                        value_format = reports_view_fallback_format
                    else:
                        value_format = total_format if is_total_row else cell_format

                worksheet.write(
                    excel_row,
                    col_num,
                    value,
                    value_format
                )

    last_data_row = first_data_row + len(df_to_write) - 1

    worksheet.autofilter(
        startrow,
        0,
        last_data_row,
        len(visible_columns) - 1
    )

    for col_num, col_name in enumerate(visible_columns):
        max_len = len(str(col_name))

        for value in df_to_write[col_name].head(300):
            if pd.notna(value):
                max_len = max(max_len, len(str(value)))

        width = min(max(max_len + 2, 12), 28)
        worksheet.set_column(col_num, col_num, width)

    return last_data_row + 3


def build_full_excel(
    vessel_name,
    segments_df,
    reports_df,
    avg_df,
    telemetry_calc_df,
    dt1,
    resolved_map,
    voltage,
    vessel_info_row,
    output_excel,
    ctx
):
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Combined_Report")
        writer.sheets["Combined_Report"] = worksheet

        worksheet.hide_gridlines(2)

        info_label_format = workbook.add_format({
            "bold": True,
            "bg_color": "#DDE6EE",
            "border": 1
        })

        info_value_format = workbook.add_format({
            "border": 1
        })

        include_fw_pumps_table = should_include_fw_pumps_table(vessel_info_row)

        combined_report_tables = []

        if segments_df.empty:
            worksheet.write(0, 0, "No segments available.")

        else:
            for _, segment_row in segments_df.iterrows():
                segment_title = build_segment_description(segment_row)

                report_level_df = build_report_level_table(
                    segment_row=segment_row,
                    reports_df=reports_df,
                    avg_df=avg_df,
                    telemetry_calc_df=telemetry_calc_df,
                    ctx=ctx
                )

                sw_vfd_df, fw_vfd_df, fan_vfd_df = build_vfd_report_level_tables(
                    segment_row=segment_row,
                    reports_df=reports_df,
                    dt1=dt1,
                    resolved_map=resolved_map,
                    voltage=voltage,
                    vessel_info_row=vessel_info_row,
                    ctx=ctx
                )

                combined_segment_df = combine_report_tables_horizontally(
                    main_df=report_level_df,
                    sw_df=sw_vfd_df,
                    fw_df=fw_vfd_df if include_fw_pumps_table else None,
                    fans_df=fan_vfd_df
                )
                combined_segment_df = combined_segment_df.drop(
                    columns=["Report"],
                    errors="ignore"
                )
                combined_segment_df.insert(0, "Segment Description", segment_title)
                combined_report_tables.append(combined_segment_df)

            combined_report_df = pd.concat(
                combined_report_tables,
                ignore_index=True
            )
            write_table_to_single_excel_sheet(
                worksheet=worksheet,
                workbook=workbook,
                df=combined_report_df,
                startrow=0
            )

        worksheet.freeze_panes(1, 0)


def build_full_excel_bytes(
    vessel_name,
    segments_df,
    reports_df,
    avg_df,
    telemetry_calc_df,
    dt1,
    resolved_map,
    voltage,
    vessel_info_row,
    ctx
):
    output_stream = io.BytesIO()

    build_full_excel(
        vessel_name=vessel_name,
        segments_df=segments_df,
        reports_df=reports_df,
        avg_df=avg_df,
        telemetry_calc_df=telemetry_calc_df,
        dt1=dt1,
        resolved_map=resolved_map,
        voltage=voltage,
        vessel_info_row=vessel_info_row,
        output_excel=output_stream,
        ctx=ctx
    )

    output_stream.seek(0)
    return output_stream.getvalue()


def build_summary_excel(
    vessel_name,
    segments_df,
    reports_df,
    avg_df,
    telemetry_calc_df,
    dt1,
    resolved_map,
    voltage,
    vessel_info_row,
    output_excel,
    ctx
):
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Final_Summary")
        writer.sheets["Final_Summary"] = worksheet

        worksheet.hide_gridlines(2)

        main_title_format = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "font_color": "#B30000"
        })

        info_label_format = workbook.add_format({
            "bold": True,
            "bg_color": "#DDE6EE",
            "border": 1
        })

        info_value_format = workbook.add_format({
            "border": 1
        })

        row = 0

        worksheet.merge_range(
            row,
            0,
            row,
            8,
            f"{vessel_name} - VFD SYSTEM REPORT - FINAL SUMMARY",
            main_title_format
        )

        row += 2

        final_motor_summary_df = build_final_motor_summary_by_segment_type(
            segments_df=segments_df,
            reports_df=reports_df,
            dt1=dt1,
            resolved_map=resolved_map,
            voltage=voltage,
            vessel_info_row=vessel_info_row,
            ctx=ctx,
            avg_df=avg_df,
            telemetry_calc_df=telemetry_calc_df
        )

        final_main_kpi_summary_df = build_final_main_kpi_summary_by_segment_type(
            segments_df=segments_df,
            reports_df=reports_df,
            avg_df=avg_df,
            telemetry_calc_df=telemetry_calc_df,
            ctx=ctx
        )

        final_main_kpi_summary_df = add_total_usd_savings_to_main_kpi_summary(
            final_main_kpi_summary_df,
            final_motor_summary_df
        )

        installation_date = get_vessel_installation_date(vessel_info_row)
        retrofit_cost_value = get_vessel_retrofit_cost_value(vessel_info_row)
        retrofit_cost = get_vessel_retrofit_cost(vessel_info_row)
        total_vfd_savings_usd = get_total_vfd_savings_usd(final_motor_summary_df)
        roi_percent = format_percent_value(
            calculate_roi_percent(total_vfd_savings_usd, retrofit_cost_value)
        )
        vfd_missing_data_stats = calculate_vfd_missing_data_stats(
            dt1,
            ctx.global_start_date,
            ctx.global_end_date,
            sample_minutes=VFD_SAMPLE_MINUTES
        )
        vfd_missing_data_percent = format_percent_value(
            vfd_missing_data_stats["missing_percent"]
        )

        general_info = {
            "Start Date": ctx.global_start_date.strftime("%Y-%m-%d"),
            "End Date": ctx.global_end_date.strftime("%Y-%m-%d"),
            "VFD Sample Minutes": VFD_SAMPLE_MINUTES,
            "Installation Date": installation_date,
            "VFD Retrofit Cost": retrofit_cost,
            "ROI(%)": roi_percent,
            "Missing VFD Data(%)": vfd_missing_data_percent,
            "Number of Segments": len(segments_df)
        }

        for key, value in general_info.items():
            worksheet.write(row, 0, key, info_label_format)
            worksheet.write(row, 1, value, info_value_format)
            row += 1

        row += 2

        row = write_table_to_single_excel_sheet(
            worksheet=worksheet,
            workbook=workbook,
            df=final_motor_summary_df,
            startrow=row,
            title="Final Summary - Motor Category Summary by Segment Type"
        )

        write_table_to_single_excel_sheet(
            worksheet=worksheet,
            workbook=workbook,
            df=final_main_kpi_summary_df,
            startrow=row,
            title="Final Summary - Main Report KPIs Summary by Segment Type"
        )

        worksheet.freeze_panes(6, 0)


def build_summary_excel_bytes(
    vessel_name,
    segments_df,
    reports_df,
    avg_df,
    telemetry_calc_df,
    dt1,
    resolved_map,
    voltage,
    vessel_info_row,
    ctx
):
    output_stream = io.BytesIO()

    build_summary_excel(
        vessel_name=vessel_name,
        segments_df=segments_df,
        reports_df=reports_df,
        avg_df=avg_df,
        telemetry_calc_df=telemetry_calc_df,
        dt1=dt1,
        resolved_map=resolved_map,
        voltage=voltage,
        vessel_info_row=vessel_info_row,
        output_excel=output_stream,
        ctx=ctx
    )

    output_stream.seek(0)
    return output_stream.getvalue()


# =====================================================
# DATA LOADING - EXCEL INPUTS
# =====================================================

def load_all_data(ctx, vessel_info_row=None):
    voyages = pd.DataFrame()
    reports = pd.DataFrame()
    avg_data = pd.DataFrame()
    dt1 = pd.DataFrame()
    telemetry_calc = pd.DataFrame()

    average_table_name = None
    vfd_table = None
    vfd_resolved_columns = {}

    ctx.log(f"Input data directory: {DATA_DIR}")

    raw_workbook = find_vessel_raw_workbook(ctx.vessel_name)
    ctx.log(f"Matched vessel raw workbook: {raw_workbook.name}")

    voyages_path = resolve_input_excel(
        VOYAGES_FILE_CANDIDATES,
        keyword_groups=[["voyage"]],
        required=False
    )
    reports_path = resolve_input_excel(
        REPORTS_FILE_CANDIDATES,
        keyword_groups=[["report", "viewer"], ["reports", "view"]],
        required=False
    )
    telemetry_calc_path = resolve_input_excel(
        TELEMETRY_CALC_FILE_CANDIDATES,
        keyword_groups=[["consumption"], ["midday", "report"]],
        required=False
    )

    # ---------------- VOYAGES ----------------
    voyages, voyages_source = load_shared_or_embedded_sheet(
        vessel_name=ctx.vessel_name,
        common_path=voyages_path,
        common_sheet_candidates=["Sheet1", "VOYAGES"],
        raw_workbook=raw_workbook,
        embedded_sheet_candidates=["VOYAGES"],
        ctx=ctx
    )

    if not voyages.empty:
        voyage_rename = {}
        normalized_cols = {str(c).strip().lower(): c for c in voyages.columns}
        aliases = {
            "vessel name": ["vessel name", "vessel_name", "vessel"],
            "arrival date": ["arrival date", "arrival_date"],
            "departure date": ["departure date", "departure_date"],
            "port name": ["port name", "port_name", "port"]
        }
        target_names = {
            "vessel name": "Vessel Name",
            "arrival date": "Arrival Date",
            "departure date": "Departure Date",
            "port name": "Port Name"
        }

        for logical_name, candidates in aliases.items():
            for candidate in candidates:
                actual = normalized_cols.get(candidate)
                if actual is not None:
                    voyage_rename[actual] = target_names[logical_name]
                    break

        voyages = voyages.rename(columns=voyage_rename)

        if "Vessel Name" not in voyages.columns:
            voyages["Vessel Name"] = ctx.vessel_name
        if "Port Name" not in voyages.columns:
            voyages["Port Name"] = "UNKNOWN PORT"

        for required_col in ["Arrival Date", "Departure Date"]:
            if required_col not in voyages.columns:
                voyages[required_col] = pd.NaT
            voyages[required_col] = coerce_excel_datetime_series(voyages[required_col])

        voyages["Port Name"] = voyages["Port Name"].fillna("UNKNOWN PORT")
        voyages.loc[
            voyages["Port Name"].astype(str).str.strip().eq(""),
            "Port Name"
        ] = "UNKNOWN PORT"

        voyages = voyages[
            (voyages["Arrival Date"] <= ctx.global_end_date) &
            (
                voyages["Departure Date"].isna() |
                (voyages["Departure Date"] >= ctx.global_start_date)
            )
        ].copy()
        voyages = voyages.sort_values("Arrival Date").reset_index(drop=True)

    ctx.log(f"Loaded voyages: {len(voyages)} rows from {voyages_source or 'no source'}")

    # ---------------- REPORTS VIEW ----------------
    reports, reports_source = load_shared_or_embedded_sheet(
        vessel_name=ctx.vessel_name,
        common_path=reports_path,
        common_sheet_candidates=["Sheet1", "REPORTS_VIEW"],
        raw_workbook=raw_workbook,
        embedded_sheet_candidates=["REPORTS_VIEW"],
        ctx=ctx
    )

    if not reports.empty:
        reports.columns = [str(c).strip() for c in reports.columns]

        if "REPORT_DT_UTC" not in reports.columns:
            raise KeyError("REPORT_DT_UTC column is missing from the reports input.")
        if "REPORT_TYPE" not in reports.columns:
            raise KeyError("REPORT_TYPE column is missing from the reports input.")

        reports["REPORT_DT_UTC"] = coerce_excel_datetime_series(
            reports["REPORT_DT_UTC"]
        )
        allowed_report_types = ["Arrival", "Port", "Departure", "Noon", "Shift", "Drift"]
        reports = reports[
            reports["REPORT_TYPE"].astype(str).str.strip().isin(allowed_report_types) &
            (reports["REPORT_DT_UTC"] >= ctx.global_start_date) &
            (reports["REPORT_DT_UTC"] <= ctx.global_end_date)
        ].copy()
        reports = preprocess_reports(reports)

    ctx.log(f"Loaded reports: {len(reports)} rows from {reports_source or 'no source'}")

    # ---------------- TELEMETRY CALCULATIONS / MIDDAY REPORTS ----------------
    telemetry_calc, telemetry_source = load_shared_or_embedded_sheet(
        vessel_name=ctx.vessel_name,
        common_path=telemetry_calc_path,
        common_sheet_candidates=["Sheet1", "MIDDAY REPORTS"],
        raw_workbook=raw_workbook,
        embedded_sheet_candidates=["MIDDAY REPORTS"],
        ctx=ctx
    )

    if not telemetry_calc.empty:
        telemetry_calc.columns = [str(c).strip() for c in telemetry_calc.columns]

        if "REPORT_DT_UTC" in telemetry_calc.columns:
            telemetry_calc["REPORT_DT_UTC"] = coerce_excel_datetime_series(
                telemetry_calc["REPORT_DT_UTC"]
            )
            telemetry_calc = telemetry_calc[
                (telemetry_calc["REPORT_DT_UTC"] >= ctx.global_start_date) &
                (telemetry_calc["REPORT_DT_UTC"] <= ctx.global_end_date)
            ].copy()

        for col in telemetry_calc.columns:
            if col not in COMMON_VESSEL_COLUMN_CANDIDATES + ["REPORT_TYPE", "REPORT_DT_UTC"]:
                telemetry_calc[col] = pd.to_numeric(
                    telemetry_calc[col],
                    errors="coerce"
                )

        telemetry_calc = telemetry_calc.sort_values("REPORT_DT_UTC").reset_index(drop=True)

    ctx.log(
        f"Loaded telemetry calculations: {len(telemetry_calc)} rows "
        f"from {telemetry_source or 'no source'}"
    )

    # ---------------- HIGH-FREQUENCY AVERAGES / TELEMETRY ----------------
    average_sheet = first_existing_sheet(raw_workbook, ["AVERAGES", "TELEMETRY"])
    if average_sheet is None:
        ctx.log(f"No AVERAGES or TELEMETRY sheet found in {raw_workbook.name}")
    else:
        avg_data = read_excel_sheet(
            raw_workbook,
            average_sheet,
            ctx=ctx,
            date_filter=build_excel_date_filter(
                column="DateTimeStamp",
                start_date=ctx.global_start_date,
                end_date=ctx.global_end_date,
                kind="excel_datetime"
            )
        )
        avg_data.columns = [str(c).strip() for c in avg_data.columns]

        if "DateTimeStamp" not in avg_data.columns:
            datetime_col = find_first_existing_column(
                avg_data,
                ["DateTimeStamp", "DATETIMESTAMP", "Date Time", "Datetime"]
            )
            if datetime_col is not None:
                avg_data = avg_data.rename(columns={datetime_col: "DateTimeStamp"})

        if "DateTimeStamp" not in avg_data.columns:
            raise KeyError(
                f"DateTimeStamp column is missing from {raw_workbook.name} [{average_sheet}]."
            )

        avg_data["DateTimeStamp"] = coerce_excel_datetime_series(
            avg_data["DateTimeStamp"]
        )
        avg_data = avg_data[
            (avg_data["DateTimeStamp"] >= ctx.global_start_date) &
            (avg_data["DateTimeStamp"] <= ctx.global_end_date)
        ].copy()

        avg_resolved_columns = resolve_average_columns(
            avg_data,
            AVERAGE_CANDIDATES
        )
        avg_data = rename_average_columns_to_standard(
            avg_data,
            avg_resolved_columns
        )
        avg_data = avg_data.sort_values("DateTimeStamp").reset_index(drop=True)
        average_table_name = f"{raw_workbook.name}::{average_sheet}"
        ctx.log(f"Loaded average/telemetry data: {len(avg_data)} rows from {average_table_name}")

    # ---------------- VFD RAW DATA ----------------
    vfd_sheet = first_existing_sheet(raw_workbook, ["VFD_RAW", "VFD"])
    if vfd_sheet is None:
        raise ValueError(f"No VFD_RAW sheet found in {raw_workbook.name}.")

    dt1 = read_excel_sheet(
        raw_workbook,
        vfd_sheet,
        ctx=ctx,
        date_filter=build_excel_date_filter(
            column="Date",
            start_date=ctx.global_start_date,
            end_date=ctx.global_end_date,
            kind="date_only"
        )
    )
    dt1.columns = [str(c).strip() for c in dt1.columns]
    vfd_resolved_columns = resolve_vfd_columns(dt1, VFD_CANDIDATES)
    dt1 = preprocess_vfd_data(
        dt1,
        vfd_resolved_columns,
        vessel_info_row=vessel_info_row,
        ctx=ctx
    )

    if not dt1.empty and "DateTimeStamp" in dt1.columns:
        ctx.log("=== BEFORE VFD DATE FILTER ===")
        ctx.log(f"Rows: {len(dt1)}")
        ctx.log(f"Min DateTimeStamp: {dt1['DateTimeStamp'].min()}")
        ctx.log(f"Max DateTimeStamp: {dt1['DateTimeStamp'].max()}")

        dt1 = dt1[
            (dt1["DateTimeStamp"] >= ctx.global_start_date) &
            (dt1["DateTimeStamp"] <= ctx.global_end_date)
        ].copy()
        dt1 = dt1.sort_values("DateTimeStamp").reset_index(drop=True)

        ctx.log("=== AFTER VFD DATE FILTER ===")
        ctx.log(f"Rows: {len(dt1)}")
        if not dt1.empty:
            ctx.log(f"Min DateTimeStamp: {dt1['DateTimeStamp'].min()}")
            ctx.log(f"Max DateTimeStamp: {dt1['DateTimeStamp'].max()}")

        dt1 = precompute_vfd_counter_increments(
            dt1=dt1,
            resolved_map=vfd_resolved_columns,
            vessel_info_row=vessel_info_row,
            sample_minutes=VFD_SAMPLE_MINUTES,
            ctx=ctx
        )

    vfd_table = f"{raw_workbook.name}::{vfd_sheet}"

    return {
        "voyages": voyages,
        "reports": reports,
        "avg_data": avg_data,
        "dt1": dt1,
        "telemetry_calc": telemetry_calc,
        "average_table_name": average_table_name,
        "vfd_table": vfd_table,
        "vfd_resolved_columns": vfd_resolved_columns,
        "raw_workbook": str(raw_workbook),
        "voyages_source": voyages_source,
        "reports_source": reports_source,
        "telemetry_source": telemetry_source
    }


# =====================================================
# MAIN REPORT GENERATOR
# =====================================================

def generate_report(vessel_name, start_date, end_date, log_callback=None):
    vessel_name_clean = str(vessel_name).strip()

    # ---------------- VESSEL INFO EXCEL ----------------
    # Loaded before ReportContext so blank date inputs can use vessel defaults.
    vessel_info_path = resolve_input_excel(
        VESSEL_INFO_FILE_CANDIDATES,
        keyword_groups=[["power", "voltage", "sw", "fans"], ["vessel", "info"]],
        required=True
    )
    vessel_info_df = load_vessel_info_excel(vessel_info_path)
    vessel_info_row = get_vessel_info_row(vessel_info_df, vessel_name_clean)

    start_date_was_blank = is_blank_input(start_date)
    end_date_was_blank = is_blank_input(end_date)
    resolved_start_date, resolved_end_date = resolve_report_date_inputs(
        start_date=start_date,
        end_date=end_date,
        vessel_info_row=vessel_info_row
    )

    ctx = ReportContext(
        vessel_name=vessel_name_clean,
        start_date=resolved_start_date,
        end_date=resolved_end_date,
        log_callback=log_callback
    )

    ctx.log("🚀 Starting report generation")
    ctx.log(f"Vessel: {ctx.vessel_name}")
    ctx.log(
        f"Date range: {ctx.global_start_date.strftime('%Y-%m-%d')} "
        f"to {ctx.global_end_date.strftime('%Y-%m-%d')}"
    )

    if start_date_was_blank:
        ctx.log(
            "Start Date not provided. Using Installation Date: "
            f"{ctx.global_start_date.strftime('%Y-%m-%d')}"
        )

    if end_date_was_blank:
        ctx.log(
            "End Date not provided. Using MC Assessment data cutoff: "
            f"{resolved_end_date.strftime('%Y-%m-%d')}"
        )

    vessel_voltage = get_vessel_voltage(
        vessel_info_row,
        ctx=ctx
    )

    ctx.log(f"✅ Vessel Voltage: {vessel_voltage}")

    # ---------------- LOAD ALL DATA ----------------
    data = load_all_data(ctx, vessel_info_row=vessel_info_row)

    voyages = data["voyages"]
    reports = data["reports"]
    avg_data = data["avg_data"]
    dt1 = data["dt1"]
    telemetry_calc = data["telemetry_calc"]
    average_table_name = data["average_table_name"]
    vfd_table = data["vfd_table"]
    vfd_resolved_columns = data["vfd_resolved_columns"]

    # ---------------- BUILD SEGMENTS ----------------
    segments_df = build_segments(
        voyages_df=voyages,
        reports_df=reports,
        avg_df=avg_data,
        dt1=dt1,
        ctx=ctx
    )

    if not segments_df.empty:
        ctx.log(f"✅ Built segments: {len(segments_df)} rows")
    else:
        ctx.log("❌ No segments were built")

    output_excel = build_output_paths(ctx.vessel_name)

    # ---------------- EXPORT EXCEL ----------------
    build_full_excel(
        vessel_name=ctx.vessel_name,
        segments_df=segments_df,
        reports_df=reports,
        avg_df=avg_data,
        telemetry_calc_df=telemetry_calc,
        dt1=dt1,
        resolved_map=vfd_resolved_columns,
        voltage=vessel_voltage,
        vessel_info_row=vessel_info_row,
        output_excel=output_excel,
        ctx=ctx
    )
    ctx.log(f"✅ Excel report created: {output_excel}")

    return {
        "excel_path": str(output_excel),
        "segments_count": len(segments_df),
        "average_table_name": average_table_name,
        "vfd_table": vfd_table
    }


# =====================================================
# OPTIONAL COMMAND LINE TEST
# =====================================================

if __name__ == "__main__":
    print(f"Input data directory: {DATA_DIR}")
    print(f"Per-vessel data directory: {RAW_DATA_DIR}")
    print("Enter a vessel name (for example Vessel 1) or ALL for every vessel.")
    vessel_name = input("Enter Vessel Name or ALL: ").strip()
    start_date = input(
        "Enter Start Date (YYYY-MM-DD, blank = Installation Date): "
    ).strip()
    end_date = input(
        "Enter End Date (YYYY-MM-DD, blank = 2026-04-01 data cutoff): "
    ).strip()

    if vessel_name.upper() == "ALL":
        vessels = list_available_vessels()
        if not vessels:
            raise FileNotFoundError(
                f"No vessel raw Excel workbooks were found in {RAW_DATA_DIR}."
            )

        completed = 0
        failures = []
        print(f"\nGenerating reports for {len(vessels)} vessels...")

        for vessel in vessels:
            print(f"\n==================== {vessel} ====================")
            try:
                result = generate_report(
                    vessel_name=vessel,
                    start_date=start_date,
                    end_date=end_date
                )
                completed += 1
            except Exception as exc:
                failures.append((vessel, str(exc)))
                print(f"Report failed for {vessel}: {exc}")

        print("\n==================== ALL VESSELS RESULT ====================")
        print(f"Completed: {completed}")
        print(f"Failed: {len(failures)}")

        for failed_vessel, error in failures:
            print(f"- {failed_vessel}: {error}")
    else:
        result = generate_report(
            vessel_name=vessel_name,
            start_date=start_date,
            end_date=end_date
        )

        print("\n==================== RESULT ====================")
        print(result)
