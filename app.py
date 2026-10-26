from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime
import io
import math
import os
import re
import sys
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
from PIL import Image
import pypdf
import streamlit as st

# ============================================================
# 1. Streamlit Page Config
# ============================================================
st.set_page_config(
    page_title="WIN-SQUARE | BOQ & Measurement Sheet Engine",
    layout="wide",
    page_icon="🪟",
    initial_sidebar_state="expanded",
)

# ============================================================
# 2. FIX CSS: Header and Sidebar Toggle Button Styling
# ============================================================
st.markdown(
    """
    <style>
    header[data-testid="stHeader"] {
        z-index: 99999 !important;
        background: transparent !important;
    }

    button[data-testid="stSidebarCollapsedControl"],
    button[data-testid="stSidebarNavCollapseButton"] {
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        background-color: #FF4B4B !important;
        color: white !important;
        border-radius: 8px !important;
        position: fixed !important;
        top: 12px !important;
        left: 12px !important;
        z-index: 999999 !important;
        box-shadow: 0px 3px 8px rgba(0,0,0,0.3) !important;
    }

    button[data-testid="stSidebarCollapsedControl"] svg,
    button[data-testid="stSidebarNavCollapseButton"] svg {
        fill: white !important;
        color: white !important;
        width: 22px !important;
        height: 22px !important;
    }

    [data-testid="stStatusWidget"],
    #MainMenu, 
    footer {
        display: none !important;
        visibility: hidden !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# State Management
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# ============================================================
# 3. UI Layout & Fonts CSS
# ============================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #f4f6f9;
        color: #334155;
    }

    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 98%;
    }

    [data-testid="stSidebar"] {
        background-color: #f1f5f9;
        border-right: 1px solid #e2e8f0;
    }
    
    .quick-guide-title {
        font-size: 15px;
        font-weight: 700;
        color: #0f172a;
        margin-top: 15px;
        margin-bottom: 12px;
    }
    
    .quick-guide-step {
        font-size: 13px;
        color: #475569;
        margin-bottom: 10px;
        line-height: 1.4;
    }

    .hero-container {
        background: #ffffff;
        border-radius: 16px;
        padding: 24px 30px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        margin-bottom: 24px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    .hero-title-text {
        font-size: 22px;
        font-weight: 800;
        color: #0f172a;
        margin: 0;
    }

    .hero-sub-text {
        font-size: 13px;
        color: #64748b;
        margin-top: 4px;
    }

    .step-title {
        font-size: 16px;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .kpi-card-box {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }

    .kpi-title-lbl {
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .kpi-val-lbl {
        font-size: 24px;
        font-weight: 800;
        color: #0f172a;
        margin-top: 6px;
    }

    div.stButton > button[kind="primary"] {
        background-color: #2563eb !important;
        background: #2563eb !important;
        border: 1px solid #2563eb !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        height: 38px !important;
        padding: 0 16px !important;
        box-shadow: 0 1px 2px rgba(37, 99, 235, 0.2) !important;
    }

    div.stButton > button[kind="secondary"] {
        background-color: #dc2626 !important;
        background: #dc2626 !important;
        border: 1px solid #dc2626 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        height: 38px !important;
        padding: 0 16px !important;
        box-shadow: 0 1px 2px rgba(220, 38, 38, 0.2) !important;
    }

    div.stDownloadButton > button {
        background-color: #059669 !important;
        background: #059669 !important;
        border: 1px solid #059669 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        height: 38px !important;
        padding: 0 16px !important;
        box-shadow: 0 1px 2px rgba(5, 150, 105, 0.2) !important;
    }

    div.stButton > button p, div.stButton > button span,
    div.stDownloadButton > button p, div.stDownloadButton > button span {
        color: #ffffff !important;
        font-weight: 500 !important;
        font-size: 13px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def get_image_path(filename: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    logo_file = get_image_path("logo.png")
    if os.path.exists(logo_file):
        col_s1, col_s2, col_s3 = st.columns([1, 2, 1])
        with col_s2:
            st.image(Image.open(logo_file), width=110)
    else:
        st.markdown(
            "<h2 style='text-align: center; color:#1e293b;'><b>win square</b></h2>",
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.markdown(
        "<div class='quick-guide-title'>💡 Quick Guide</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class='quick-guide-step'><b>1.</b> Select app mode: <b>Glass Summary Engine</b> or <b>Horizontal MEASUREMENT Sheet</b>.</div>
        <div class='quick-guide-step'><b>2.</b> Upload multi-sheet Excel BOQ or PDF files.</div>
        <div class='quick-guide-step'><b>3.</b> Process files to automatically extract glass, specs, drawings & Qty splits (W1 -> W1A, W1B...).</div>
        <div class='quick-guide-step'><b>4.</b> Review live preview & download the final formatted Excel sheet.</div>
        """,
        unsafe_allow_html=True,
    )

# =========================================================
# HEADER HERO BANNER
# =========================================================
st.markdown(
    """
    <div class="hero-container">
        <div>
            <div class="hero-title-text">Requirement & Measurement Engine</div>
            <div class="hero-sub-text">Enterprise BOQ Extraction, Quantity Splitting, Drawing Mapping & Measurement Generator</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# Global Helper Functions & Parsing Utilities
# ============================================================

HEADER_SCAN_LIMIT = 200
HEADER_REMOVE_PATTERN = r"[^A-Z0-9]"

KEYWORDS = {
    "CODE": ["CODE"],
    "WIDTH": [
        ["GL", "W"],
        ["GLS", "W"],
        ["GLASS", "W"],
        ["GL", "WIDTH"],
        ["GLS", "WIDTH"],
        ["GLASS", "WIDTH"],
        ["S", "GLS", "W"],
        ["SGLS", "W"],
        ["S", "GL", "W"],
    ],
    "HEIGHT": [
        ["GL", "H"],
        ["GLS", "H"],
        ["GLASS", "H"],
        ["GL", "HEIGHT"],
        ["GLS", "HEIGHT"],
        ["GLASS", "HEIGHT"],
        ["S", "GLS", "H"],
        ["SGLS", "H"],
        ["S", "GL", "H"],
    ],
    "QTY": ["QTY"],
    "GLASS": [["GLASS"], ["DESP"]],
}


def get_code_suffix(index: int) -> str:
    """Generates A, B, C... Z, AA, AB... for quantity splitting."""
    result = ""
    while True:
        result = chr(65 + (index % 26)) + result
        index = (index // 26) - 1
        if index < 0:
            break
    return result


def extract_glass_thickness(glass_text: Any) -> str:
    """Extracts glass thickness and formats as '20 MM', '5 MM', '11.52 MM', etc."""
    if not glass_text or pd.isna(glass_text):
        return ""
    text = str(glass_text).strip()

    match = re.search(r"(\d+(?:\.\d+)?)\s*(MM|mm)", text)
    if match:
        return f"{match.group(1)} MM"

    match_lam = re.search(r"\)(\d+(?:\.\d+)?)\s*(MM|mm)?", text, re.IGNORECASE)
    if match_lam:
        return f"{match_lam.group(1)} MM"

    return ""


def standardize_glass_spec(val: Any) -> str:
    if pd.isna(val) or not str(val).strip():
        return "NOT SPECIFIED"
    text = str(val).strip()
    if text.lower() == "nan" or not text:
        return "NOT SPECIFIED"
    return re.sub(r"\s+", " ", text).strip()


@dataclass(slots=True)
class HeaderInfo:
    row_index: int
    code_col: Optional[int] = None
    width_col: Optional[int] = None
    height_col: Optional[int] = None
    qty_col: Optional[int] = None
    glass_col: Optional[int] = None
    columns: Dict[str, Optional[int]] = field(default_factory=dict)


@dataclass(slots=True)
class HeaderBlock:
    header: HeaderInfo
    start_row: int
    end_row: int


@dataclass(slots=True)
class GlassRecord:
    WindowCode: str
    Width: int
    Height: int
    Qty: int
    GlassType: str
    SourceFile: str
    SheetName: str


def normalize_header(text: Any) -> str:
    if pd.isna(text):
        return ""
    text = str(text).upper().strip()
    return re.sub(HEADER_REMOVE_PATTERN, "", text)


def normalize_header_row(row: pd.Series) -> List[str]:
    return [normalize_header(val) for val in row.tolist()]


def contains_keywords(text: str, keyword_groups: List[Any]) -> bool:
    if not text:
        return False
    text = normalize_header(text)
    text = re.sub(r"[^A-Z0-9]", "", text)

    for group in keyword_groups:
        if isinstance(group, str):
            group = [group]
        matched = True
        for keyword in group:
            key = re.sub(r"[^A-Z0-9]", "", normalize_header(keyword))
            if key not in text:
                matched = False
                break
        if matched:
            return True
    return False


def detect_column(header_row: List[str], keyword_groups: List[Any]) -> Optional[int]:
    for index, value in enumerate(header_row):
        text = normalize_header(value)
        text = re.sub(r"[^A-Z0-9]", "", text)
        text = (
            text.replace("GLASS", "GL")
            .replace("GLAZING", "GL")
            .replace("SGLS", "GLS")
            .replace("SGL", "GL")
            .replace("GLS.", "GLS")
        )
        for group in keyword_groups:
            if isinstance(group, str):
                group = [group]
            matched = True
            for keyword in group:
                key = re.sub(r"[^A-Z0-9]", "", normalize_header(keyword))
                if key not in text:
                    matched = False
                    break
            if matched:
                return index
    return None


def detect_header_columns(header_row: pd.Series) -> Dict[str, Optional[int]]:
    normalized = normalize_header_row(header_row)
    columns = {
        "code": detect_column(normalized, KEYWORDS["CODE"]),
        "width": detect_column(normalized, KEYWORDS["WIDTH"]),
        "height": detect_column(normalized, KEYWORDS["HEIGHT"]),
        "qty": detect_column(normalized, KEYWORDS["QTY"]),
        "glass": detect_column(normalized, KEYWORDS["GLASS"]),
    }

    if columns["width"] is None:
        for kw in [
            ["S", "GLS", "W"],
            ["S", "GL", "W"],
            ["GLS", "W"],
            ["GL", "W"],
            ["S", "GZ", "W"],
            ["GZ", "W"],
            ["FWIDTH"],
            ["F", "WIDTH"],
            ["SWIDTH"],
            ["S", "WIDTH"],
        ]:
            col = detect_column(normalized, kw)
            if col is not None:
                columns["width"] = col
                break

    if columns["height"] is None:
        for kw in [
            ["S", "GLS", "H"],
            ["S", "GL", "H"],
            ["GLS", "H"],
            ["GL", "H"],
            ["S", "GZ", "H"],
            ["GZ", "H"],
            ["FHEIGHT"],
            ["F", "HEIGHT"],
            ["SHEIGHT"],
            ["S", "HEIGHT"],
        ]:
            col = detect_column(normalized, kw)
            if col is not None:
                columns["height"] = col
                break

    return columns


def is_business_header(row: pd.Series) -> bool:
    normalized = normalize_header_row(row)
    has_code = False
    has_qty = False
    has_glass = False

    for value in normalized:
        if contains_keywords(value, KEYWORDS["CODE"]):
            has_code = True
        if contains_keywords(value, KEYWORDS["QTY"]):
            has_qty = True
        if (
            contains_keywords(value, KEYWORDS["GLASS"])
            or contains_keywords(value, KEYWORDS["WIDTH"])
            or contains_keywords(value, KEYWORDS["HEIGHT"])
        ):
            has_glass = True

    return has_code and has_qty and has_glass


def find_header_blocks(dataframe: pd.DataFrame) -> List[HeaderInfo]:
    headers: List[HeaderInfo] = []
    rows = min(len(dataframe), HEADER_SCAN_LIMIT)

    for row_number in range(rows):
        row = dataframe.iloc[row_number]
        if not is_business_header(row):
            continue

        columns = detect_header_columns(row)
        header = HeaderInfo(
            row_index=row_number,
            code_col=columns["code"],
            width_col=columns["width"],
            height_col=columns["height"],
            qty_col=columns["qty"],
            glass_col=columns["glass"],
            columns=columns,
        )
        headers.append(header)

    return headers


def build_header_blocks(
    dataframe: pd.DataFrame, headers: List[HeaderInfo]
) -> List[HeaderBlock]:
    blocks: List[HeaderBlock] = []
    if not headers:
        return blocks

    headers = sorted(headers, key=lambda h: h.row_index)
    for i, header in enumerate(headers):
        start = header.row_index + 1
        end = (
            len(dataframe) - 1
            if i == len(headers) - 1
            else headers[i + 1].row_index - 1
        )
        blocks.append(HeaderBlock(header=header, start_row=start, end_row=end))

    return blocks


def score_business_sheet(df: pd.DataFrame) -> int:
    score = 0
    scan_rows = min(len(df), HEADER_SCAN_LIMIT)

    for r in range(scan_rows):
        row = normalize_header_row(df.iloc[r])
        text = " ".join(row)
        if "CODE" in text:
            score += 5
        if "QTY" in text:
            score += 8
        if "GLASS" in text:
            score += 8
        if re.search(r"\bGL.*W\b", text):
            score += 10
        if re.search(r"\bGL.*H\b", text):
            score += 10

    return score


def find_business_sheets(
    workbook: Dict[str, pd.DataFrame],
) -> List[Tuple[str, pd.DataFrame]]:
    business_sheets = []
    threshold = 35

    for sheet_name, df in workbook.items():
        score = score_business_sheet(df)
        accept_sheet = False

        if score >= threshold:
            accept_sheet = True
        else:
            header_blocks = find_header_blocks(df)
            if header_blocks:
                accept_sheet = True

        if accept_sheet:
            business_sheets.append((sheet_name, df))

    return business_sheets


def safe_numeric(value: Any) -> Optional[int]:
    if pd.isna(value):
        return None
    try:
        val = float(value)
        if val <= 0:
            return None
        return int(math.floor(val + 0.5))
    except Exception:
        return None


def build_window_code(row: pd.Series, header: HeaderInfo) -> Optional[str]:
    if header.code_col is None:
        return None

    val = row.iloc[header.code_col]
    if pd.isna(val):
        return None

    code = str(val).strip()
    if code == "" or code.lower() == "nan":
        return None

    code = re.sub(r"\.0$", "", code)
    code = re.sub(r"\s+", " ", code).strip()

    if header.code_col + 1 < len(row):
        next_val = row.iloc[header.code_col + 1]
        if not pd.isna(next_val):
            next_str = str(next_val).strip()
            next_str = re.sub(r"\.0$", "", next_str)

            if next_str and next_str.lower() != "nan":
                invalid_headers = [
                    "FWIDTH",
                    "FHEIGHT",
                    "GLSW",
                    "GLSH",
                    "GLASS",
                    "QTY",
                    "DESCRIPTION",
                    "WIDTH",
                    "HEIGHT",
                ]
                is_numeric_dim = next_str.isdigit() or re.match(
                    r"^\d+(\.\d+)?$", next_str
                )
                is_valid_location = (
                    not is_numeric_dim
                    and next_str.upper() not in invalid_headers
                )

                if is_valid_location:
                    code = f"{code} {next_str}"

    return re.sub(r"\s+", " ", code).strip()


def is_record_start(row: pd.Series, header: HeaderInfo) -> bool:
    try:
        code_parts = []
        for col in range(min(3, len(row))):
            value = row.iloc[col]
            if pd.isna(value):
                continue
            val_str = str(value).strip()
            if val_str == "":
                continue
            code_parts.append(normalize_header(val_str))

        if not code_parts:
            return False
        code = " ".join(code_parts).strip()

        invalid = (
            "CODE",
            "FWIDTH",
            "FHEIGHT",
            "GLSW",
            "GLSH",
            "GLASS",
            "QTY",
            "DESCRIPTION",
        )
        if code in invalid or any(inv == code for inv in invalid):
            return False

        upper = code.upper()
        if any(w in upper for w in ["FRAME", "WINDOW", "SECTION", "PROFILE"]):
            return False
        if not re.search(r"[A-Z0-9]", upper):
            return False

        return True
    except Exception:
        return False


def build_record_buffers(
    dataframe: pd.DataFrame, block: HeaderBlock
) -> List[pd.DataFrame]:
    buffers: List[pd.DataFrame] = []
    current_rows = []

    for row_no in range(block.start_row, block.end_row + 1):
        row = dataframe.iloc[row_no]
        if is_record_start(row, block.header):
            if current_rows:
                buffers.append(pd.DataFrame(current_rows))
                current_rows = []
        current_rows.append(row)

    if current_rows:
        buffers.append(pd.DataFrame(current_rows))

    return buffers


def collect_numeric_from_buffer(
    buffer: pd.DataFrame, column_index: Optional[int]
) -> Optional[int]:
    if column_index is None:
        return None
    for _, row in buffer.iterrows():
        if column_index >= len(row):
            continue
        val = safe_numeric(row.iloc[column_index])
        if val is not None:
            return val
    return None


def collect_glass(buffer: pd.DataFrame, header: HeaderInfo) -> Optional[str]:
    parts = []
    seen = set()

    if header.glass_col is not None:
        for _, row in buffer.iterrows():
            if header.glass_col >= len(row):
                continue
            val = row.iloc[header.glass_col]
            if pd.isna(val):
                continue
            val_str = re.sub(r"\s+", " ", str(val).strip()).strip()
            if val_str == "" or val_str.upper() in seen:
                continue
            seen.add(val_str.upper())
            parts.append(val_str)

    if parts:
        return " ".join(parts)

    for _, row in buffer.iterrows():
        for cell in row.values:
            if pd.isna(cell):
                continue
            cell_str = str(cell).strip()
            upper_cell = cell_str.upper()
            if any(
                k in upper_cell
                for k in [
                    "LAMINATED",
                    "TOUGHENED",
                    "DGU",
                    "SGU",
                    "CLEAR",
                    "FROSTED",
                    "MM",
                ]
            ):
                if not any(
                    k in upper_cell
                    for k in [
                        "CODE",
                        "HANDLE",
                        "LOCK",
                        "COLOR",
                        "FRAME",
                        "WIDTH",
                        "HEIGHT",
                    ]
                ):
                    return cell_str

    return None


def parse_header_block(
    dataframe: pd.DataFrame,
    block: HeaderBlock,
    source_file: str,
    sheet_name: str,
) -> List[GlassRecord]:
    records: List[GlassRecord] = []
    buffers = build_record_buffers(dataframe, block)

    for buffer in buffers:
        first_row = buffer.iloc[0]
        window = build_window_code(first_row, block.header)

        width = collect_numeric_from_buffer(buffer, block.header.width_col)
        height = collect_numeric_from_buffer(buffer, block.header.height_col)
        qty = collect_numeric_from_buffer(buffer, block.header.qty_col)
        glass_raw = collect_glass(buffer, block.header)

        if not window or width is None or height is None:
            continue
        if qty is None:
            qty = 1

        # ACCURATE FROSTED FILTER: Only ignore 'FROSTED' / '5MM FROSTED' when 'TOUGHENED' is NOT present
        if glass_raw:
            glass_clean = re.sub(r"\s+", " ", str(glass_raw)).strip().upper()
            if "FROSTED" in glass_clean and "TOUGHENED" not in glass_clean:
                continue

        glass = standardize_glass_spec(glass_raw)

        records.append(
            GlassRecord(
                WindowCode=window,
                Width=width,
                Height=height,
                Qty=qty,
                GlassType=glass,
                SourceFile=source_file,
                SheetName=sheet_name,
            )
        )

    return records


def parse_business_sheet(
    dataframe: pd.DataFrame, source_file: str, sheet_name: str
) -> List[GlassRecord]:
    headers = find_header_blocks(dataframe)
    blocks = build_header_blocks(dataframe, headers)
    all_records: List[GlassRecord] = []

    for block in blocks:
        all_records.extend(
            parse_header_block(dataframe, block, source_file, sheet_name)
        )

    return all_records


def load_excel_with_calculated_values(file: Any) -> Dict[str, pd.DataFrame]:
    file_bytes = io.BytesIO(file.read())
    file.seek(0)
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    workbook_dict = {}

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        data = sheet.values
        cols = next(data, None)
        if cols is None:
            continue

        data_rows = list(data)
        if cols:
            data_rows.insert(0, cols)

        df = pd.DataFrame(data_rows)
        workbook_dict[sheet_name] = df

    return workbook_dict


def process_uploaded_files(uploaded_files: List[Any]) -> pd.DataFrame:
    all_records = []

    for file in uploaded_files:
        try:
            workbook_dict = load_excel_with_calculated_values(file)
            filtered_workbook = {}
            for s_name, df in workbook_dict.items():
                s_name_upper = s_name.upper().strip()
                if any(k in s_name_upper for k in ["CLIENT", "DETAIL", "QUOTE"]):
                    continue
                filtered_workbook[s_name] = df

            business_sheets = find_business_sheets(filtered_workbook)
            for sheet_name, df in business_sheets:
                records = parse_business_sheet(df, file.name, sheet_name)
                all_records.extend(records)

        except Exception as e:
            st.error(f"Error processing file {file.name}: {e}")

    return (
        pd.DataFrame([asdict(r) for r in all_records]).reset_index(drop=True)
        if all_records
        else pd.DataFrame()
    )


# ============================================================
# HORIZONTAL MEASUREMENT SHEET PARSING (EXCEL + PDF)
# ============================================================

def parse_pdf_quotation(file_bytes: io.BytesIO) -> List[Dict[str, Any]]:
    """PDF parsing engine with image extraction and Qty split (W1 -> W1A, W1B...)."""
    reader = pypdf.PdfReader(file_bytes)
    extracted_items = []

    for page in reader.pages:
        page_text = page.extract_text() or ""

        page_images = []
        try:
            for img_obj in page.images:
                page_images.append(img_obj.data)
        except Exception:
            pass

        code_blocks = re.split(r"(Code\s*:\s*)", page_text, flags=re.IGNORECASE)

        for i in range(1, len(code_blocks), 2):
            block_text = code_blocks[i] + code_blocks[i + 1]

            code_match = re.search(r"Code\s*:\s*([^\n\r]+)", block_text)
            name_match = re.search(r"Name\s*:\s*([^\n\r]+)", block_text)
            loc_match = re.search(r"Location\s*:\s*([^\n\r]+)", block_text)
            size_match = re.search(
                r"Size\s*:\s*W\s*=\s*([\d\.]+);\s*H\s*=\s*([\d\.]+)", block_text
            )
            prof_match = re.search(r"Profile System\s*:\s*([^\n\r]+)", block_text)
            glass_match = re.search(r"Glass\s*:\s*([^\n\r]+)", block_text)
            qty_match = re.search(
                r"Quantity\s*:\s*(\d+)", block_text, flags=re.IGNORECASE
            )

            base_code = code_match.group(1).strip() if code_match else ""
            name = name_match.group(1).strip() if name_match else ""
            location = loc_match.group(1).strip() if loc_match else ""
            w_val = float(size_match.group(1)) if size_match else ""
            h_val = float(size_match.group(2)) if size_match else ""
            prof_sys = prof_match.group(1).strip() if prof_match else ""
            glass = glass_match.group(1).strip() if glass_match else ""
            qty_val = int(qty_match.group(1)) if qty_match else 1

            matched_img_data = page_images[0] if page_images else None

            if base_code:
                if qty_val > 1:
                    for q in range(qty_val):
                        suffix = get_code_suffix(q)
                        split_code = f"{base_code}{suffix}"
                        extracted_items.append({
                            "code": split_code,
                            "name": name,
                            "location": location,
                            "w_val": w_val,
                            "h_val": h_val,
                            "qty": 1,
                            "prof_sys": prof_sys,
                            "glass": glass,
                            "mesh": "NO",
                            "lock": "MULTI POINT",
                            "remarks": "",
                            "image_data": matched_img_data,
                        })
                else:
                    extracted_items.append({
                        "code": base_code,
                        "name": name,
                        "location": location,
                        "w_val": w_val,
                        "h_val": h_val,
                        "qty": 1,
                        "prof_sys": prof_sys,
                        "glass": glass,
                        "mesh": "NO",
                        "lock": "MULTI POINT",
                        "remarks": "",
                        "image_data": matched_img_data,
                    })

    return extracted_items


def parse_excel_quotation(file_bytes: io.BytesIO) -> List[Dict[str, Any]]:
    """Excel quotation parsing engine with Qty split (W1 -> W1A, W1B...) and drawing image capture."""
    wb_source = openpyxl.load_workbook(file_bytes, data_only=True)
    extracted_items = []

    for sheetname in wb_source.sheetnames:
        src_ws = wb_source[sheetname]
        images = getattr(src_ws, "_images", [])

        for r in range(1, src_ws.max_row + 1):
            c1_val = str(src_ws.cell(r, 1).value or "").strip()

            if c1_val == "Code :":
                base_code = str(src_ws.cell(r, 5).value or "").strip()
                size_str = str(
                    src_ws.cell(r, 12).value or src_ws.cell(r, 13).value or ""
                )
                name = str(src_ws.cell(r + 1, 5).value or "").strip()
                prof_sys = str(
                    src_ws.cell(r + 1, 12).value or src_ws.cell(r + 1, 13).value or ""
                ).strip()
                location = str(src_ws.cell(r + 2, 5).value or "").strip()
                glass = str(
                    src_ws.cell(r + 2, 12).value or src_ws.cell(r + 2, 13).value or ""
                ).strip()

                w_val, h_val = "", ""
                if "W =" in size_str and ";" in size_str:
                    try:
                        parts = size_str.split(";")
                        w_val = float(parts[0].split("=")[1].strip())
                        h_val = float(parts[1].split("=")[1].strip())
                    except Exception:
                        pass

                qty_val = 1
                for search_qty_r in range(r, min(r + 20, src_ws.max_row + 1)):
                    lbl_qty = str(src_ws.cell(search_qty_r, 8).value or "").strip()
                    if lbl_qty == "Quantity":
                        try:
                            qty_val = int(
                                src_ws.cell(search_qty_r, 11).value
                                or src_ws.cell(search_qty_r, 12).value
                                or 1
                            )
                        except Exception:
                            qty_val = 1
                        break

                mesh_type, lock_type, remarks_desc = "NO", "", ""

                for offset in range(10, 22):
                    v_mesh = str(src_ws.cell(r + offset, 8).value or "")
                    if "MeshType :" in v_mesh:
                        mesh_type = v_mesh.replace("MeshType :", "").strip() or "NO"

                    v_lock = str(
                        src_ws.cell(r + offset, 13).value
                        or src_ws.cell(r + offset, 14).value
                        or ""
                    )
                    if "Locking :" in v_lock:
                        lock_type = v_lock.replace("Locking :", "").strip()

                for search_r in range(r, min(r + 25, src_ws.max_row + 1)):
                    cell_rem = str(src_ws.cell(search_r, 8).value or "")
                    if "Remarks :" in cell_rem:
                        remarks_desc = cell_rem.replace("Remarks :", "").strip()

                matched_img_data = None
                for img in images:
                    anchor = img.anchor
                    if hasattr(anchor, "_from"):
                        img_row = anchor._from.row + 1
                        if r <= img_row <= r + 22 and img_row > 5:
                            matched_img_data = img._data()
                            break

                if qty_val > 1:
                    for q in range(qty_val):
                        suffix = get_code_suffix(q)
                        split_code = f"{base_code}{suffix}"
                        extracted_items.append({
                            "code": split_code,
                            "name": name,
                            "location": location,
                            "w_val": w_val,
                            "h_val": h_val,
                            "qty": 1,
                            "prof_sys": prof_sys,
                            "glass": glass,
                            "mesh": mesh_type,
                            "lock": lock_type,
                            "remarks": remarks_desc,
                            "image_data": matched_img_data,
                        })
                else:
                    extracted_items.append({
                        "code": base_code,
                        "name": name,
                        "location": location,
                        "w_val": w_val,
                        "h_val": h_val,
                        "qty": 1,
                        "prof_sys": prof_sys,
                        "glass": glass,
                        "mesh": mesh_type,
                        "lock": lock_type,
                        "remarks": remarks_desc,
                        "image_data": matched_img_data,
                    })

    return extracted_items


def build_measurement_excel(items: List[Dict[str, Any]]) -> Tuple[io.BytesIO, int]:
    """Generates horizontal MEASUREMENT sheet with formatted Glass Thickness (e.g., '20 MM')."""
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "MEASUREMENT"
    ws_out.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(
        start_color="A6A085", end_color="A6A085", fill_type="solid"
    )
    font_header = Font(name="Arial", size=10, bold=True, color="000000")
    font_data = Font(name="Arial", size=10, bold=False, color="000000")

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws_out.row_dimensions[1].height = 28
    ws_out.row_dimensions[2].height = 20

    headers_m = [
        ("A1:A2", "S.NO"),
        ("B1:B2", "FLAT NO"),
        ("C1:C2", "WINDOW CODE"),
        ("D1:D2", "WINDOW TYPE"),
        ("E1:E2", "LOCATION"),
        ("F1:F2", "PROFILE"),
        ("G1:H1", "Dimension in mm"),
        ("I1:I2", "Area in Sq. Ft"),
        ("J1:J2", "Mosquito Mesh"),
        ("K1:K2", "Lock Type"),
        ("L1:M1", "Glass / FLYMESH"),
        ("N1:N2", "Remarks"),
    ]

    for cell_range, text in headers_m:
        if ":" in cell_range:
            ws_out.merge_cells(cell_range)
            ws_out[cell_range.split(":")[0]].value = text
        else:
            ws_out[cell_range].value = text

    ws_out["G2"] = "Width"
    ws_out["H2"] = "Height"
    ws_out["L2"] = "Thickness"
    ws_out["M2"] = "Spece."

    for r in range(1, 3):
        for c in range(1, 15):
            cell = ws_out.cell(row=r, column=c)
            cell.fill = header_fill
            cell.font = font_header
            cell.alignment = align_center
            cell.border = thin_border

    curr_row = 3
    s_no = 1

    for item in items:
        ws_out.row_dimensions[curr_row].height = 24

        ws_out.cell(row=curr_row, column=1, value=s_no)
        ws_out.cell(row=curr_row, column=2, value="")
        ws_out.cell(row=curr_row, column=3, value=item["code"])
        ws_out.cell(row=curr_row, column=4, value=item["name"])
        ws_out.cell(row=curr_row, column=5, value=item["location"])
        ws_out.cell(row=curr_row, column=6, value=item["prof_sys"])
        ws_out.cell(row=curr_row, column=7, value=item["w_val"])
        ws_out.cell(row=curr_row, column=8, value=item["h_val"])

        c_sq = ws_out.cell(
            row=curr_row,
            column=9,
            value=f"=ROUND((G{curr_row}/304.8)*(H{curr_row}/304.8), 2)",
        )
        c_sq.number_format = "0.00"

        glass_thick = extract_glass_thickness(item["glass"])

        ws_out.cell(row=curr_row, column=10, value=item["mesh"])
        ws_out.cell(row=curr_row, column=11, value=item["lock"])
        ws_out.cell(row=curr_row, column=12, value=glass_thick)
        ws_out.cell(row=curr_row, column=13, value=item["glass"])
        ws_out.cell(row=curr_row, column=14, value=item["remarks"])

        for c in range(1, 15):
            cell = ws_out.cell(row=curr_row, column=c)
            cell.font = font_data
            cell.border = thin_border
            if c in [1, 2, 3, 7, 8, 9, 10, 11, 12]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

        if item["image_data"]:
            try:
                img_bytes = io.BytesIO(item["image_data"])
                img_copy = openpyxl.drawing.image.Image(img_bytes)
                ws_out.add_image(img_copy, f"C{curr_row + 2}")
            except Exception:
                pass

        curr_row += 18
        s_no += 1

    col_widths = {
        "A": 6,
        "B": 10,
        "C": 16,
        "D": 22,
        "E": 18,
        "F": 18,
        "G": 10,
        "H": 10,
        "I": 14,
        "J": 15,
        "K": 15,
        "L": 12,
        "M": 25,
        "N": 25,
    }
    for col_letter, width in col_widths.items():
        ws_out.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb_out.save(output)
    output.seek(0)
    return output, len(items)


# ============================================================
# APP NAVIGATION & MAIN CONTROLLER
# ============================================================

app_mode = st.radio(
    "📌 Select Application Workflow:",
    [
        "📊 Glass Summary Engine (Extract BOQ Glass Records)",
        "📐 Horizontal MEASUREMENT Sheet Engine (Excel & PDF Support)",
    ],
    horizontal=True,
)

st.markdown("<br>", unsafe_allow_html=True)

if app_mode == "📊 Glass Summary Engine (Extract BOQ Glass Records)":
    st.markdown(
        "<div class='step-title'>📁 Step 1: Upload BOQ Excel Files</div>",
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Upload BOQ Excel Files",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key=f"boq_uploader_{st.session_state['uploader_key']}",
    )

    st.markdown("<br>", unsafe_allow_html=True)

    btn_col1, btn_col2, _ = st.columns([1, 1, 6])

    with btn_col1:
        if st.button("Merge & Process Files", type="primary", use_container_width=False):
            if uploaded_files:
                with st.spinner("Extracting & Merging Records..."):
                    df_merged = process_uploaded_files(uploaded_files)
                    if not df_merged.empty:
                        st.session_state["merged_df"] = df_merged
                        st.toast(
                            f"Successfully Extracted {len(df_merged)} Non-Frosted Glass Records!",
                            icon="✅",
                        )
                    else:
                        st.error("⚠️ No valid non-frosted glass records found.")
            else:
                st.warning("Please upload Excel file(s) first!")

    with btn_col2:
        if st.button("Reset Data", type="secondary", use_container_width=False):
            for key in ["merged_df", "req_df_preview", "req_bytes", "req_generated"]:
                if key in st.session_state:
                    del st.session_state[key]
            st.session_state["uploader_key"] += 1
            st.rerun()

    if "merged_df" in st.session_state:
        df_merged = st.session_state["merged_df"]

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            "<div class='step-title'>📋 Extracted Glass Records (Frosted Excluded)</div>",
            unsafe_allow_html=True,
        )

        f_col1, f_col2 = st.columns([2, 2])
        with f_col1:
            search_query = st.text_input(
                "🔍 Quick Search (Window Code / Glass Spec)",
                placeholder="Type to filter...",
            )
        with f_col2:
            glass_types = ["ALL"] + sorted(list(df_merged["GlassType"].unique()))
            selected_glass = st.selectbox("Filter by Glass Spec", glass_types)

        filtered_df = df_merged.copy()
        if search_query:
            filtered_df = filtered_df[
                filtered_df["WindowCode"].str.contains(
                    search_query, case=False, na=False
                )
                | filtered_df["GlassType"].str.contains(
                    search_query, case=False, na=False
                )
            ]
        if selected_glass != "ALL":
            filtered_df = filtered_df[filtered_df["GlassType"] == selected_glass]

        filtered_display_df = filtered_df.copy()
        if "Sr. No." not in filtered_display_df.columns:
            filtered_display_df.insert(
                0, "Sr. No.", range(1, len(filtered_display_df) + 1)
            )

        st.dataframe(
            filtered_display_df,
            use_container_width=True,
            height=280,
            hide_index=True,
        )
        st.caption(
            f"Showing {len(filtered_df)} of {len(df_merged)} extracted non-frosted records"
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            "<div class='step-title'>⚡ Step 2: Generate Official Requirement Sheet</div>",
            unsafe_allow_html=True,
        )

        if st.button(
            "⚡ GENERATE REQUIREMENT SHEET (MEASUREMENTS)",
            type="primary",
            use_container_width=False,
        ):
            with st.spinner("Calculating SQFT..."):
                df_req_preview = df_merged.copy()
                df_req_preview["SQFT"] = (
                    (df_req_preview["Width"] * df_req_preview["Height"]) / 92903.04
                ).round(2)
                df_req_preview["TTL SQFT"] = (
                    df_req_preview["SQFT"] * df_req_preview["Qty"]
                ).round(2)

                df_req_preview.insert(0, "Sr.No", range(1, len(df_req_preview) + 1))
                df_req_preview = df_req_preview.rename(
                    columns={
                        "WindowCode": "WINDOW CODE",
                        "Width": "WIDTH",
                        "Height": "HEIGHT",
                        "Qty": "QTY",
                        "GlassType": "REMARKS",
                    }
                )

                preview_cols = [
                    "Sr.No",
                    "WINDOW CODE",
                    "WIDTH",
                    "HEIGHT",
                    "SQFT",
                    "QTY",
                    "TTL SQFT",
                    "REMARKS",
                ]
                df_req_preview = df_req_preview[preview_cols]
                st.session_state["req_df_preview"] = df_req_preview

                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Sheet1"
                ws.views.sheetView[0].showGridLines = True

                title_font = Font(name="Calibri", size=16, bold=True, color="000000")
                header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                data_font = Font(name="Calibri", size=11, bold=False, color="000000")
                total_font = Font(name="Calibri", size=11, bold=True, color="000000")

                title_fill = PatternFill(
                    start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"
                )
                header_fill = PatternFill(
                    start_color="1F497D", end_color="1F497D", fill_type="solid"
                )

                thin_border = Border(
                    left=Side(style="thin", color="000000"),
                    right=Side(style="thin", color="000000"),
                    top=Side(style="thin", color="000000"),
                    bottom=Side(style="thin", color="000000"),
                )

                today_str = datetime.now().strftime("%d %b %Y").upper()
                title_text = f"1 WIN-SQUARE {today_str}"
                st.session_state["generated_title_name"] = title_text

                ws.merge_cells("A1:H1")
                ws.row_dimensions[1].height = 28
                ws["A1"].value = title_text
                ws["A1"].font = title_font
                ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

                for col_i in range(1, 9):
                    cell = ws.cell(row=1, column=col_i)
                    cell.fill = title_fill
                    cell.border = thin_border

                headers = [
                    "Sr.No",
                    "WINDOW CODE",
                    "WIDTH",
                    "HEIGHT",
                    "SQFT",
                    "QTY",
                    "TTL SQFT",
                    "REMARKS",
                ]
                ws.row_dimensions[2].height = 22
                for col_i, h_text in enumerate(headers, 1):
                    cell = ws.cell(row=2, column=col_i, value=h_text)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border

                for idx, row in df_merged.iterrows():
                    r_idx = idx + 3
                    sqft_formula = f"=ROUND((C{r_idx}*D{r_idx})/92903.04, 2)"
                    ttl_sqft_formula = f"=E{r_idx}*F{r_idx}"

                    row_data = [
                        idx + 1,
                        row["WindowCode"],
                        row["Width"],
                        row["Height"],
                        sqft_formula,
                        row["Qty"],
                        ttl_sqft_formula,
                        row["GlassType"],
                    ]

                    ws.row_dimensions[r_idx].height = 20

                    for col_i, val in enumerate(row_data, 1):
                        cell = ws.cell(row=r_idx, column=col_i, value=val)
                        cell.font = data_font
                        cell.border = thin_border

                        if col_i in [1, 2, 8]:
                            cell.alignment = Alignment(
                                horizontal="center", vertical="center"
                            )
                        elif col_i in [3, 4, 6]:
                            cell.number_format = "0"
                            cell.alignment = Alignment(
                                horizontal="center", vertical="center"
                            )
                        elif col_i in [5, 7]:
                            cell.number_format = "0.00"
                            cell.alignment = Alignment(
                                horizontal="center", vertical="center"
                            )

                tot_row = len(df_merged) + 3
                ws.row_dimensions[tot_row].height = 22

                for c in range(1, 9):
                    ws.cell(row=tot_row, column=c).border = thin_border

                cell_tot = ws.cell(row=tot_row, column=5, value="TOTAL")
                cell_tot.font = total_font
                cell_tot.alignment = Alignment(horizontal="right", vertical="center")

                qty_sum = ws.cell(row=tot_row, column=6, value=f"=SUM(F3:F{tot_row-1})")
                qty_sum.font = total_font
                qty_sum.number_format = "0"
                qty_sum.alignment = Alignment(horizontal="center", vertical="center")

                ttl_sqft_sum = ws.cell(
                    row=tot_row, column=7, value=f"=SUM(G3:G{tot_row-1})"
                )
                ttl_sqft_sum.font = total_font
                ttl_sqft_sum.number_format = "0.00"
                ttl_sqft_sum.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

                col_widths = {
                    "A": 8,
                    "B": 16,
                    "C": 10,
                    "D": 10,
                    "E": 10,
                    "F": 8,
                    "G": 12,
                    "H": 32,
                }
                for col_letter, width in col_widths.items():
                    ws.column_dimensions[col_letter].width = width

                output = io.BytesIO()
                wb.save(output)
                st.session_state["req_bytes"] = output.getvalue()
                st.session_state["req_generated"] = True

        if st.session_state.get("req_generated"):
            st.markdown("<br>", unsafe_allow_html=True)

            req_df = st.session_state["req_df_preview"]
            tot_items = len(req_df)
            tot_qty = req_df["QTY"].sum()
            tot_area = req_df["TTL SQFT"].sum().round(2)

            k1, k2, k3 = st.columns(3)
            with k1:
                st.markdown(
                    f"<div class='kpi-card-box'><div class='kpi-title-lbl'>TOTAL ITEMS</div><div class='kpi-val-lbl'>{tot_items}</div></div>",
                    unsafe_allow_html=True,
                )
            with k2:
                st.markdown(
                    f"<div class='kpi-card-box'><div class='kpi-title-lbl'>TOTAL GLASS QUANTITY</div><div class='kpi-val-lbl'>{tot_qty} Pcs</div></div>",
                    unsafe_allow_html=True,
                )
            with k3:
                st.markdown(
                    f"<div class='kpi-card-box'><div class='kpi-title-lbl'>TOTAL GLASS SQFT (NON-FROSTED)</div><div class='kpi-val-lbl'>{tot_area:,.2f} Sq.Ft</div></div>",
                    unsafe_allow_html=True,
                )

            st.markdown("<br>", unsafe_allow_html=True)

            tab1, tab2, tab3 = st.tabs([
                "📄 MEASUREMENTS Live Preview",
                "📊 OC Wise Summary (Glass SQFT)",
                "🧩 Glass Type Breakdown",
            ])

            with tab1:
                st.dataframe(
                    req_df.style.format({"SQFT": "{:.2f}", "TTL SQFT": "{:.2f}"}),
                    use_container_width=True,
                    height=350,
                    hide_index=True,
                )

            with tab2:
                df_merged_copy = df_merged.copy()
                df_merged_copy["Total_SQFT"] = (
                    (df_merged_copy["Width"] * df_merged_copy["Height"]) / 92903.04
                ) * df_merged_copy["Qty"]

                def clean_glass_name(text: Any) -> str:
                    text_str = str(text).upper().strip()
                    text_str = re.sub(r"\s+", " ", text_str)
                    return text_str.replace("TOUGHENED", "THGN").replace(
                        "TOUGHNED", "THGN"
                    )

                df_merged_copy["CleanGlassType"] = df_merged_copy["GlassType"].apply(
                    clean_glass_name
                )

                def make_glass_string(group: pd.DataFrame) -> str:
                    summary = group.groupby("CleanGlassType")["Qty"].sum()
                    details = [
                        f"{g_type} - {qty}"
                        for g_type, qty in summary.items()
                        if g_type != "NOT SPECIFIED"
                    ]
                    return ", ".join(details) if details else "-"

                glass_details_series = (
                    df_merged_copy.groupby("SourceFile")
                    .apply(make_glass_string, include_groups=False)
                    .reset_index(name="GLASS DETAILS")
                )

                oc_summary = df_merged_copy.groupby("SourceFile", as_index=False).agg(
                    Qty=("Qty", "sum"), Total_SQFT=("Total_SQFT", "sum")
                )

                oc_summary = pd.merge(oc_summary, glass_details_series, on="SourceFile")
                oc_summary["Total_SQFT"] = oc_summary["Total_SQFT"].round(2)

                oc_summary.columns = [
                    "SourceFile (OC Name)",
                    "Qty (Pcs)",
                    "Total Glass SQFT",
                    "GLASS DETAILS",
                ]

                if "Sr. No." not in oc_summary.columns:
                    oc_summary.insert(0, "Sr. No.", range(1, len(oc_summary) + 1))

                st.dataframe(oc_summary, use_container_width=True, hide_index=True)

            with tab3:
                df_glass_copy = df_merged.copy()

                def clean_glass_name_tab3(text: Any) -> str:
                    text_str = str(text).upper().strip()
                    text_str = re.sub(r"\s+", " ", text_str)
                    return text_str.replace("TOUGHNED", "TOUGHENED")

                df_glass_copy["CleanGlassType"] = df_glass_copy["GlassType"].apply(
                    clean_glass_name_tab3
                )
                df_glass_filtered = df_glass_copy[
                    df_glass_copy["CleanGlassType"] != "NOT SPECIFIED"
                ]

                glass_breakdown = (
                    df_glass_filtered.groupby("CleanGlassType", as_index=False)["Qty"]
                    .sum()
                    .sort_values(by="Qty", ascending=False)
                )

                glass_breakdown.columns = ["GlassType", "Qty"]
                glass_breakdown.insert(0, "Sr. No.", range(1, len(glass_breakdown) + 1))

                st.dataframe(glass_breakdown, use_container_width=True, hide_index=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.success("✅ Requirement Excel Sheet Ready!")

            file_download_name = (
                f"{st.session_state.get('generated_title_name', '1 WIN-SQUARE')}.xlsx"
            )
            st.download_button(
                label=f"📥 DOWNLOAD OFFICIAL REQUIREMENT SHEET ({file_download_name})",
                data=st.session_state["req_bytes"],
                file_name=file_download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=False,
            )

else:
    st.markdown(
        "<div class='step-title'>📐 Upload Quotation File (.xlsx / .pdf)</div>",
        unsafe_allow_html=True,
    )

    uploaded_h_file = st.file_uploader(
        "Upload File (.xlsx, .xlsm, .pdf)",
        type=["xlsx", "xlsm", "pdf"],
        key="horizontal_uploader",
    )

    if st.button("🚀 Generate MEASUREMENT Sheet", type="primary"):
        if not uploaded_h_file:
            st.error("Please upload an Excel or PDF Quotation file first!")
        else:
            file_bytes = io.BytesIO(uploaded_h_file.read())
            uploaded_h_file.seek(0)

            with st.spinner("Processing file, splitting Qty & formatting thickness..."):
                if uploaded_h_file.name.endswith(".pdf"):
                    extracted_items = parse_pdf_quotation(file_bytes)
                else:
                    extracted_items = parse_excel_quotation(file_bytes)

                if extracted_items:
                    excel_out, count = build_measurement_excel(extracted_items)
                    st.success(
                        f"Successfully generated MEASUREMENT sheet ({count} total unique items split)!"
                    )

                    out_filename = (
                        f"MEASUREMENT_{uploaded_h_file.name.rsplit('.', 1)[0]}.xlsx"
                    )
                    st.download_button(
                        label="📥 Download MEASUREMENT Excel Sheet",
                        data=excel_out,
                        file_name=out_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=False,
                    )
                else:
                    st.warning("No valid Window items found in the uploaded file.")
