
import re
import io
import os
import calendar
import logging
import warnings
from datetime import datetime, timedelta
from collections import defaultdict

import streamlit as st
import pdfplumber

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment
)
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# ============================================================
# 1. APPLICATION CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="SPRAC | Attendance Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Suppress repetitive, non-fatal PDF font descriptor warnings.
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfplumber").setLevel(logging.ERROR)
warnings.filterwarnings(
    "ignore",
    message="Could not get FontBBox.*"
)

# SPRAC brand palette
NAVY = "#16324F"
TEAL = "#087E8B"
ACCENT = "#12A4A6"
WHITE = "#FFFFFF"


# ============================================================
# 2. THEME MANAGEMENT
# ============================================================

if "app_theme" not in st.session_state:
    st.session_state.app_theme = "Light"

with st.sidebar:
    st.markdown("## SPRAC ATTENDANCE")
    st.caption("Workforce Analytics Portal")
    st.divider()

    st.markdown("### Display")
    theme = st.radio(
        "Choose appearance",
        ["Light", "Dark"],
        horizontal=True,
        index=0 if st.session_state.app_theme == "Light" else 1,
        label_visibility="collapsed"
    )

    st.session_state.app_theme = theme

DARK = st.session_state.app_theme == "Dark"

if DARK:
    BG = "#0B1220"
    PANEL = "#131F30"
    PANEL_ALT = "#1A2A3D"
    TEXT = "#F1F5F9"
    MUTED = "#A6B5C8"
    BORDER = "#30445D"
    INPUT_BG = "#17263A"
    HERO_START = "#16324F"
    HERO_END = "#087E8B"
else:
    BG = "#F2F5F9"
    PANEL = "#FFFFFF"
    PANEL_ALT = "#F5F8FC"
    TEXT = "#172B43"
    MUTED = "#61748B"
    BORDER = "#D9E3EF"
    INPUT_BG = "#FFFFFF"
    HERO_START = "#16324F"
    HERO_END = "#087E8B"


# ============================================================
# 3. CUSTOM STYLING
# ============================================================

st.markdown(
    f"""
    <style>
    :root {{
        --app-bg: {BG};
        --panel: {PANEL};
        --panel-alt: {PANEL_ALT};
        --text: {TEXT};
        --muted: {MUTED};
        --border: {BORDER};
        --teal: {TEAL};
        --navy: {NAVY};
    }}

    .stApp {{
        background: var(--app-bg);
        color: var(--text);
    }}

    [data-testid="stHeader"] {{
        background: var(--app-bg);
    }}

    [data-testid="stSidebar"] {{
        background: var(--panel);
        border-right: 1px solid var(--border);
    }}

    [data-testid="stSidebar"] * {{
        color: var(--text);
    }}

    .block-container {{
        padding-top: 1.6rem;
        padding-bottom: 3rem;
        max-width: 1600px;
    }}

    h1, h2, h3, h4 {{
        color: var(--text) !important;
        letter-spacing: -0.5px;
    }}

    p, label, li {{
        color: var(--text);
    }}

    [data-testid="stCaptionContainer"] p {{
        color: var(--muted) !important;
    }}

    .hero {{
        background: linear-gradient(
            115deg,
            {HERO_START},
            {HERO_END}
        );
        padding: 34px 38px;
        border-radius: 16px;
        margin-bottom: 22px;
        box-shadow: 0 8px 25px rgba(10, 30, 55, 0.12);
    }}

    .hero-eyebrow {{
        color: #BDE9F1;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: 2px;
        margin-bottom: 15px;
    }}

    .hero h1 {{
        color: #FFFFFF !important;
        font-size: clamp(27px, 3vw, 39px);
        font-weight: 750;
        margin: 0 0 12px 0;
    }}

    .hero p {{
        color: #E7F1F8 !important;
        font-size: 15px;
        line-height: 1.7;
        max-width: 850px;
        margin: 0;
    }}

    .hero-footer {{
        color: #C7E6EF;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 1.4px;
        margin-top: 22px;
    }}

    .section-heading {{
        font-size: 21px;
        font-weight: 750;
        color: var(--text);
        margin: 20px 0 5px 0;
    }}

    .section-subtitle {{
        color: var(--muted);
        font-size: 13px;
        margin-bottom: 16px;
    }}

    .upload-panel {{
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 22px;
        margin: 8px 0 20px 0;
    }}

    .upload-title {{
        font-size: 19px;
        font-weight: 750;
        color: var(--text);
        margin-bottom: 5px;
    }}

    .upload-description {{
        color: var(--muted);
        font-size: 13px;
        margin-bottom: 15px;
    }}

    [data-testid="stFileUploader"] {{
        background: var(--panel-alt);
        border: 1.5px dashed {TEAL};
        border-radius: 11px;
        padding: 12px;
    }}

    [data-testid="stFileUploader"] section {{
        background: var(--panel-alt);
    }}

    [data-testid="stFileUploader"] button {{
        background: {NAVY} !important;
        color: #FFFFFF !important;
        border: 1px solid {NAVY} !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        padding: 8px 20px !important;
    }}

    [data-testid="stFileUploader"] button:hover {{
        background: {TEAL} !important;
        border-color: {TEAL} !important;
        color: #FFFFFF !important;
    }}

    [data-testid="stFileUploader"] *,
    [data-testid="stFileUploaderDropzone"] * {{
        color: var(--text);
    }}

    [data-testid="stMetric"] {{
        background: var(--panel);
        border: 1px solid var(--border);
        padding: 16px;
        border-radius: 12px;
    }}

    [data-testid="stMetricLabel"] {{
        color: var(--muted) !important;
    }}

    [data-testid="stMetricValue"] {{
        color: var(--text) !important;
    }}

    .step-card {{
        background: var(--panel);
        border: 1px solid var(--border);
        padding: 17px;
        border-radius: 12px;
        height: 100%;
        min-height: 140px;
    }}

    .step-number {{
        color: {TEAL};
        font-size: 11px;
        font-weight: 800;
        letter-spacing: 1px;
    }}

    .step-title {{
        color: var(--text);
        font-weight: 750;
        font-size: 15px;
        margin: 10px 0 7px;
    }}

    .step-description {{
        color: var(--muted);
        font-size: 12px;
        line-height: 1.6;
    }}

    .info-panel {{
        background: var(--panel);
        border: 1px solid var(--border);
        padding: 17px;
        border-radius: 12px;
        margin: 10px 0;
    }}

    .footer {{
        border-top: 1px solid var(--border);
        margin-top: 40px;
        padding-top: 18px;
        color: var(--muted);
        font-size: 11px;
        display: flex;
        justify-content: space-between;
    }}

    .stButton > button {{
        background: {NAVY};
        color: white;
        border: 1px solid {NAVY};
        border-radius: 8px;
        font-weight: 650;
    }}

    .stButton > button:hover {{
        background: {TEAL};
        color: white;
        border-color: {TEAL};
    }}

    .stDownloadButton > button {{
        background: {TEAL} !important;
        color: white !important;
        border: 1px solid {TEAL} !important;
        border-radius: 8px;
        font-weight: 700;
        width: 100%;
        min-height: 45px;
    }}

    .stDownloadButton > button:hover {{
        background: {NAVY} !important;
        border-color: {NAVY} !important;
    }}

    [data-testid="stDataFrame"] {{
        border: 1px solid var(--border);
        border-radius: 10px;
    }}

    div[data-testid="stAlert"] {{
        border-radius: 10px;
    }}

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 4. SIDEBAR GUIDE
# ============================================================

with st.sidebar:

    st.divider()

    st.markdown("### How to use this portal")

    guide = [
        (
            "01",
            "Select your reports",
            "Choose the employee attendance PDFs."
        ),
        (
            "02",
            "Upload one month",
            "All reports must belong to the same month."
        ),
        (
            "03",
            "Review the results",
            "Check employee summaries and import warnings."
        ),
        (
            "04",
            "Download your workbook",
            "Export the consolidated Excel report."
        )
    ]

    for number, title, description in guide:
        st.markdown(
            f"""
            <div style="margin:15px 0;">
                <div style="color:{TEAL};font-size:10px;
                font-weight:800;">{number}</div>
                <div style="font-weight:700;margin:4px 0;">
                    {title}
                </div>
                <div style="color:{MUTED};font-size:11px;">
                    {description}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    st.markdown("### Before uploading")

    st.markdown(
        """
        - Use original attendance-system PDF reports.
        - Each employee should have their own report.
        - Upload reports for the same month together.
        - If the PDF does not contain an employee ID,
          its filename will be used.
        - Uploaded files are processed for this session.
        """
    )

    st.divider()

    st.caption("SPRAC SERVICES | INTERNAL HR OPERATIONS")


# ============================================================
# 5. HERO SECTION
# ============================================================

st.markdown(
    f"""
    <div class="hero">
        <div class="hero-eyebrow">
            SPRAC SERVICES &nbsp; | &nbsp; PEOPLE OPERATIONS
        </div>
        <h1>Workforce Attendance Analytics</h1>
        <p>
            A centralized workspace for monthly employee attendance
            reporting, workforce insights and management-ready
            Excel reports.
        </p>
        <div class="hero-footer">
            HUMAN RESOURCES &nbsp; • &nbsp;
            WORKFORCE REPORTING &nbsp; • &nbsp;
            OPERATIONAL INSIGHTS
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 6. PDF PARSING
# ============================================================

TIME_PATTERN = re.compile(
    r"(?<!\d)([01]?\d|2[0-3]):([0-5]\d)(?::([0-5]\d))?(?!\d)"
)


def normalize_time(value):
    """Convert a time string to HH:MM:SS."""
    match = TIME_PATTERN.search(value)

    if not match:
        return None

    hour = int(match.group(1))
    minute = int(match.group(2))
    second = int(match.group(3) or 0)

    return f"{hour:02d}:{minute:02d}:{second:02d}"


def find_month_year(text):
    """Find month/year from common attendance PDF formats."""

    patterns = [
        r"\uf073\s*(\d{1,2})-(20\d{2})",
        r"(?<!\d)(0?[1-9]|1[0-2])-(20\d{2})(?!\d)",
        r"(?<!\d)(0?[1-9]|1[0-2])/(20\d{2})(?!\d)"
    ]

    for pattern in patterns:
        match = re.search(pattern, text)

        if match:
            month = int(match.group(1))
            year = int(match.group(2))

            if 1 <= month <= 12:
                return year, month

    # Support month names, e.g. January 2026.
    match = re.search(
        r"\b(January|February|March|April|May|June|July|"
        r"August|September|October|November|December)\s+(20\d{2})\b",
        text,
        re.IGNORECASE
    )

    if match:
        month = list(calendar.month_name).index(
            match.group(1).capitalize()
        )
        return int(match.group(2)), month

    raise ValueError("Month and year could not be identified.")


def find_employee_code(text, filename):
    """Extract employee ID, otherwise use the filename."""

    patterns = [
        r"(?:Employee\s*(?:Code|ID|No\.?|Number)|"
        r"Emp\s*(?:Code|ID|No\.?))\s*[:#\-]?\s*"
        r"([A-Za-z0-9_-]+)",
        r"\bStaff\s*(?:ID|Code)\s*[:#\-]?\s*"
        r"([A-Za-z0-9_-]+)"
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            return match.group(1).strip()

    return os.path.splitext(filename)[0].strip()


def parse_pdf(file_bytes, filename):
    """
    Extract monthly attendance records.

    Expected row structure:
    Serial number + date + attendance details + two or more times.

    The last two detected times in a row are treated as
    In Time and Out Time, matching the original parser logic.
    """

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:

        pages = [page.extract_text() or "" for page in pdf.pages]

    all_text = "\n".join(pages)

    if not all_text.strip():
        raise ValueError(
            "No selectable text found. This may be a scanned PDF."
        )

    year, month = find_month_year(all_text)

    employee_code = find_employee_code(all_text, filename)

    days_in_month = calendar.monthrange(year, month)[1]

    records = {}

    row_start_pattern = re.compile(
        r"^\s*\d{1,2}\s+(\d{1,2})[-/]",
    )

    for page in pages:

        lines = page.splitlines()

        starts = [
            index
            for index, line in enumerate(lines)
            if row_start_pattern.match(line)
        ]

        for index, start in enumerate(starts):

            end = (
                starts[index + 1]
                if index + 1 < len(starts)
                else len(lines)
            )

            chunk = " ".join(lines[start:end])

            row = row_start_pattern.match(chunk)

            if not row:
                continue

            day = int(row.group(1))

            if not 1 <= day <= days_in_month:
                continue

            times = TIME_PATTERN.findall(chunk)

            if len(times) < 2:
                continue

            normalized = []

            for h, m, s in times:
                normalized.append(
                    f"{int(h):02d}:{int(m):02d}:{int(s or 0):02d}"
                )

            # Retain original parser's last-two-times logic.
            records[day] = (
                normalized[-2],
                normalized[-1]
            )

    if not records:
        raise ValueError(
            "No valid attendance rows were detected. "
            "Please verify that this is an original attendance PDF."
        )

    return {
        "employee_code": employee_code,
        "year": year,
        "month": month,
        "records": records,
        "filename": filename
    }


# ============================================================
# 7. TIME CALCULATIONS
# ============================================================

def time_to_seconds(time_string):
    t = datetime.strptime(time_string, "%H:%M:%S")
    return t.hour * 3600 + t.minute * 60 + t.second


def calculate_stay(in_time, out_time):
    """Calculate attendance duration, including overnight shifts."""

    seconds = (
        time_to_seconds(out_time)
        - time_to_seconds(in_time)
    )

    if seconds < 0:
        seconds += 86400

    return seconds


def seconds_to_excel_time(seconds):
    return seconds / 86400


def average_clock_time(times):
    """
    Calculate average clock time using a circular mean.
    This avoids incorrect averages around midnight.
    """

    if not times:
        return None

    import math

    angles = [
        (time_to_seconds(t) / 86400) * 2 * math.pi
        for t in times
    ]

    sin_avg = sum(math.sin(a) for a in angles) / len(angles)
    cos_avg = sum(math.cos(a) for a in angles) / len(angles)

    if abs(sin_avg) < 1e-10 and abs(cos_avg) < 1e-10:
        # Fallback for a mathematically undefined circular mean.
        return (
            sum(time_to_seconds(t) for t in times)
            / len(times)
        ) / 86400

    angle = math.atan2(sin_avg, cos_avg)

    if angle < 0:
        angle += 2 * math.pi

    return (angle / (2 * math.pi)) % 1


def employee_statistics(records):

    in_times = []
    out_times = []
    stays = []

    for day, (in_time, out_time) in records.items():

        in_times.append(in_time)
        out_times.append(out_time)

        stays.append(
            calculate_stay(in_time, out_time)
        )

    return {
        "avg_in": average_clock_time(in_times),
        "avg_out": average_clock_time(out_times),
        "avg_stay": (
            seconds_to_excel_time(
                sum(stays) / len(stays)
            ) if stays else None
        ),
        "worked_days": len(records),
        "total_seconds": sum(stays)
    }


# ============================================================
# 8. EXCEL GENERATION
# ============================================================

def style_header(ws, row=1):

    for cell in ws[row]:

        cell.fill = PatternFill(
            "solid",
            fgColor=NAVY.replace("#", "")
        )

        cell.font = Font(
            name="Aptos",
            size=10,
            bold=True,
            color="FFFFFF"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.row_dimensions[row].height = 32


def apply_body_style(ws):

    for row in ws.iter_rows(min_row=2):

        for cell in row:

            cell.font = Font(
                name="Aptos",
                size=10,
                color="24364B"
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            if cell.row % 2 == 0:
                cell.fill = PatternFill(
                    "solid",
                    fgColor="F0F5F9"
                )


def add_excel_table(ws, name):

    if ws.max_row < 2 or ws.max_column < 1:
        return

    ref = (
        f"A1:{get_column_letter(ws.max_column)}"
        f"{ws.max_row}"
    )

    table = Table(displayName=name, ref=ref)

    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    ws.add_table(table)


def make_excel(employees, year, month, import_log):

    wb = Workbook()

    # --------------------------------------------------------
    # SHEET 1: MONTHLY SUMMARY
    # --------------------------------------------------------

    ws = wb.active
    ws.title = "Monthly Summary"

    days = calendar.monthrange(year, month)[1]

    daily_start_col = 7

    headers = [
        "Employee ID",
        "Avg In",
        "Avg Out",
        "Avg Stay",
        "Worked Days",
        "Total Hours"
    ]

    headers.extend([
        f"Day {day:02d}"
        for day in range(1, days + 1)
    ])

    ws.append(headers)

    sorted_employees = sorted(
        employees.items(),
        key=lambda item: item[0].lower()
    )

    for employee_code, data in sorted_employees:

        records = data["records"]
        stats = employee_statistics(records)

        row = [
            employee_code,
            stats["avg_in"],
            stats["avg_out"],
            stats["avg_stay"],
            stats["worked_days"],
            seconds_to_excel_time(stats["total_seconds"])
        ]

        for day in range(1, days + 1):

            if day in records:
                in_time, out_time = records[day]

                stay = calculate_stay(
                    in_time,
                    out_time
                )

                row.append(seconds_to_excel_time(stay))
            else:
                row.append(None)

        ws.append(row)

    style_header(ws)
    apply_body_style(ws)

    ws.freeze_panes = "G2"
    ws.auto_filter.ref = ws.dimensions

    ws.column_dimensions["A"].width = 20

    for col in range(2, 5):
        ws.column_dimensions[
            get_column_letter(col)
        ].width = 15

    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 17

    for col in range(daily_start_col, daily_start_col + days):
        ws.column_dimensions[
            get_column_letter(col)
        ].width = 11

    for row in range(2, ws.max_row + 1):

        for col in [2, 3]:
            ws.cell(row, col).number_format = "hh:mm AM/PM"

        for col in [4, 6]:
            ws.cell(row, col).number_format = "[h]:mm"

        for col in range(daily_start_col, daily_start_col + days):
            ws.cell(row, col).number_format = "[h]:mm"

    # --------------------------------------------------------
    # SHEET 2: DAILY ATTENDANCE
    # --------------------------------------------------------

    daily_ws = wb.create_sheet("Daily Attendance")

    daily_ws.append([
        "Employee ID",
        "Date",
        "Day",
        "In Time",
        "Out Time",
        "Stay Duration"
    ])

    for employee_code, data in sorted_employees:

        for day in sorted(data["records"]):

            in_time, out_time = data["records"][day]

            daily_ws.append([
                employee_code,
                datetime(year, month, day),
                calendar.day_name[
                    datetime(year, month, day).weekday()
                ],
                datetime.strptime(in_time, "%H:%M:%S"),
                datetime.strptime(out_time, "%H:%M:%S"),
                seconds_to_excel_time(
                    calculate_stay(in_time, out_time)
                )
            ])

    style_header(daily_ws)
    apply_body_style(daily_ws)

    daily_ws.freeze_panes = "A2"
    daily_ws.auto_filter.ref = daily_ws.dimensions

    for col, width in {
        "A": 20,
        "B": 16,
        "C": 17,
        "D": 16,
        "E": 16,
        "F": 18
    }.items():
        daily_ws.column_dimensions[col].width = width

    for row in range(2, daily_ws.max_row + 1):
        daily_ws.cell(row, 2).number_format = "dd-mmm-yyyy"
        daily_ws.cell(row, 4).number_format = "hh:mm AM/PM"
        daily_ws.cell(row, 5).number_format = "hh:mm AM/PM"
        daily_ws.cell(row, 6).number_format = "[h]:mm"

    # --------------------------------------------------------
    # SHEET 3: IMPORT LOG
    # --------------------------------------------------------

    log_ws = wb.create_sheet("Import Log")

    log_ws.append([
        "File Name",
        "Employee ID",
        "Status",
        "Message"
    ])

    for entry in import_log:

        log_ws.append([
            entry.get("filename", ""),
            entry.get("employee_code", ""),
            entry.get("status", ""),
            entry.get("message", "")
        ])

    style_header(log_ws)
    apply_body_style(log_ws)

    log_ws.freeze_panes = "A2"
    log_ws.column_dimensions["A"].width = 40
    log_ws.column_dimensions["B"].width = 20
    log_ws.column_dimensions["C"].width = 18
    log_ws.column_dimensions["D"].width = 65

    # --------------------------------------------------------
    # WORKBOOK FORMATTING
    # --------------------------------------------------------

    for sheet in wb.worksheets:

        sheet.sheet_view.showGridLines = False

        sheet.sheet_properties.pageSetUpPr.fitToPage = True

        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.page_setup.orientation = "landscape"

        sheet.sheet_properties.outlinePr.summaryBelow = True

        sheet.sheet_view.zoomScale = 90

    output = io.BytesIO()
    wb.save(output)

    return output.getvalue()


# ============================================================
# 9. MAIN UPLOAD SECTION
# ============================================================

st.markdown(
    '<div class="section-heading">Create a monthly attendance report</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-subtitle">'
    'Upload employee attendance reports for the same month. '
    'The portal will automatically analyse and consolidate the records.'
    '</div>',
    unsafe_allow_html=True
)

# Upload appears immediately at the top of the application.
st.markdown(
    """
    <div class="upload-panel">
        <div class="upload-title">
            📁 Upload employee attendance reports
        </div>
        <div class="upload-description">
            Select multiple PDF files together. You can upload
            10–20 employee reports or more in one batch.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

uploaded_files = st.file_uploader(
    "Browse or drag and drop employee PDF reports",
    type=["pdf"],
    accept_multiple_files=True,
    help="Select all employee attendance PDFs for one month.",
    key="attendance_upload"
)


# ============================================================
# 10. PROCESSING AND VALIDATION
# ============================================================

if uploaded_files:

    st.markdown("### Uploaded files")

    st.success(
        f"{len(uploaded_files)} PDF file(s) selected."
    )

    if st.button(
        "⚡ Process attendance reports",
        type="primary",
        use_container_width=True
    ):

        employees = {}
        import_log = []

        selected_month = None
        duplicate_ids = set()

        progress = st.progress(0)
        status_text = st.empty()

        for index, uploaded_file in enumerate(uploaded_files):

            status_text.write(
                f"Processing {uploaded_file.name}..."
            )

            try:

                result = parse_pdf(
                    uploaded_file.getvalue(),
                    uploaded_file.name
                )

                employee_code = result["employee_code"]

                report_month = (
                    result["year"],
                    result["month"]
                )

                # Ensure every report belongs to the same month.
                if selected_month is None:
                    selected_month = report_month

                elif report_month != selected_month:

                    raise ValueError(
                        "This report belongs to a different month. "
                        "Upload reports for one month together."
                    )

                # Do not silently overwrite duplicate employee IDs.
                if employee_code in employees:

                    duplicate_ids.add(employee_code)

                    import_log.append({
                        "filename": uploaded_file.name,
                        "employee_code": employee_code,
                        "status": "Duplicate",
                        "message": (
                            "Another report with this employee ID "
                            "has already been processed. This file "
                            "was not imported."
                        )
                    })

                else:

                    employees[employee_code] = result

                    import_log.append({
                        "filename": uploaded_file.name,
                        "employee_code": employee_code,
                        "status": "Success",
                        "message": (
                            f"{len(result['records'])} attendance "
                            "days extracted successfully."
                        )
                    })

            except Exception as error:

                import_log.append({
                    "filename": uploaded_file.name,
                    "employee_code": "",
                    "status": "Failed",
                    "message": str(error)
                })

            progress.progress(
                (index + 1) / len(uploaded_files)
            )

        status_text.empty()
        progress.empty()

        if employees and selected_month:

            st.session_state.processed_employees = employees
            st.session_state.processed_log = import_log
            st.session_state.processed_month = selected_month

            st.success(
                "Attendance processing completed."
            )

        else:

            st.session_state.processed_employees = {}
            st.session_state.processed_log = import_log
            st.session_state.processed_month = None

            st.error(
                "No valid employee reports could be processed. "
                "Please check the Import Log."
            )


# ============================================================
# 11. RESULTS DASHBOARD
# ============================================================

if st.session_state.get("processed_employees"):

    employees = st.session_state.processed_employees
    import_log = st.session_state.processed_log

    year, month = st.session_state.processed_month

    successful = len(employees)

    failed = sum(
        1 for item in import_log
        if item["status"] == "Failed"
    )

    duplicates = sum(
        1 for item in import_log
        if item["status"] == "Duplicate"
    )

    st.divider()

    st.markdown("## Monthly attendance overview")

    st.caption(
        f"Reporting period: {calendar.month_name[month]} {year}"
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Employees processed", successful)
    c2.metric("Total PDF files", len(import_log))
    c3.metric("Import failures", failed)
    c4.metric("Duplicate reports", duplicates)

    # --------------------------------------------------------
    # EMPLOYEE PREVIEW
    # --------------------------------------------------------

    st.markdown("### Employee attendance summary")

    preview = []

    for employee_code, data in sorted(employees.items()):

        stats = employee_statistics(data["records"])

        avg_in = (
            datetime.strptime(
                "00:00:00", "%H:%M:%S"
            ) if stats["avg_in"] is None
            else stats["avg_in"]
        )

        preview.append({
            "Employee ID": employee_code,
            "Avg In": (
                datetime(1899, 12, 30)
                + timedelta(days=avg_in)
            ).strftime("%I:%M %p"),
            "Avg Out": (
                datetime(1899, 12, 30)
                + timedelta(days=stats["avg_out"])
            ).strftime("%I:%M %p"),
            "Avg Stay": (
                f"{stats['total_seconds'] / stats['worked_days'] / 3600:.2f} hrs"
            ),
            "Worked Days": stats["worked_days"],
            "Total Hours": (
                f"{stats['total_seconds'] / 3600:.2f} hrs"
            )
        })

    st.dataframe(
        preview,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------------
    # IMPORT LOG
    # --------------------------------------------------------

    with st.expander("View detailed import log"):

        st.dataframe(
            import_log,
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # EXCEL EXPORT
    # --------------------------------------------------------

    st.divider()

    st.markdown("### Download your consolidated workbook")

    st.caption(
        "Your Excel workbook contains a monthly summary, "
        "daily attendance details and a complete import log."
    )

    excel_data = make_excel(
        employees,
        year,
        month,
        import_log
    )

    filename = (
        f"SPRAC_Attendance_{calendar.month_name[month]}_{year}.xlsx"
    )

    st.download_button(
        label="⬇ Download Monthly Attendance Excel",
        data=excel_data,
        file_name=filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    st.caption(
        "The workbook is generated locally for this session. "
        "Review the import log and verify the records before "
        "using them for official HR decisions."
    )


# ============================================================
# 12. USER INSTRUCTIONS
# ============================================================

else:

    st.markdown("### Your reporting workflow")

    step1, step2, step3, step4 = st.columns(4)

    steps = [
        (
            step1,
            "STEP 01",
            "Select reports",
            "Choose the employee attendance PDFs from your computer."
        ),
        (
            step2,
            "STEP 02",
            "Upload one month",
            "All employee reports must belong to the same reporting month."
        ),
        (
            step3,
            "STEP 03",
            "Review results",
            "Check employee summaries and resolve any import warnings."
        ),
        (
            step4,
            "STEP 04",
            "Export workbook",
            "Download the consolidated Excel report for HR operations."
        )
    ]

    for column, number, title, description in steps:

        with column:

            st.markdown(
                f"""
                <div class="step-card">
                    <div class="step-number">{number}</div>
                    <div class="step-title">{title}</div>
                    <div class="step-description">
                        {description}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

    with st.expander("What will the Excel workbook contain?"):

        a, b, c = st.columns(3)

        with a:
            st.markdown("**01 · Monthly Summary**")
            st.caption(
                "Average in time, average out time, "
                "average stay, worked days, total hours "
                "and day-wise attendance."
            )

        with b:
            st.markdown("**02 · Daily Attendance**")
            st.caption(
                "A detailed employee-wise daily record "
                "of entry time, exit time and duration."
            )

        with c:
            st.markdown("**03 · Import Log**")
            st.caption(
                "A record of successful imports, "
                "duplicates and failed reports."
            )


# ============================================================
# 13. FOOTER
# ============================================================

st.markdown(
    f"""
    <div class="footer">
        <div>
            <b style="color:{TEXT};">SPRAC SERVICES</b><br>
            Workforce Attendance Analytics
        </div>
        <div style="text-align:right;">
            Internal HR Operations Portal<br>
            Monthly Attendance Reporting
        </div>
    </div>
    """,
    unsafe_allow_html=True
)
