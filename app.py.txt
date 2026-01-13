import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Font
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re
import os
import streamlit as st
from io import BytesIO # Required for in-memory file operations
import random # Added import for random, as it was used in original code

# --- Streamlit App Setup ---
st.set_page_config(layout="wide")
st.title("Επεξεργασία Ανθρωπομηνών Excel")

TEMPLATE_FILE = "AM TEST 1.xlsx"

# Check if TEMPLATE_FILE exists. If not, prompt the user to upload it.
# Streamlit's file_uploader works differently from Colab's files.upload()
if not os.path.exists(TEMPLATE_FILE):
    st.warning(f"Το αρχείο template '{TEMPLATE_FILE}' δεν βρέθηκε στον κατάλογο. Παρακαλώ ανεβάστε το.")
    uploaded_template_file = st.file_uploader(
        "Ανεβάστε το αρχείο template AM TEST 1.xlsx",
        type=["xlsx"],
        key="template_upload"
    )
    if uploaded_template_file is not None:
        # Save the uploaded template file temporarily to disk to be opened by openpyxl
        with open(TEMPLATE_FILE, "wb") as f:
            f.write(uploaded_template_file.getbuffer())
        st.success(f"Το αρχείο template '{TEMPLATE_FILE}' ανέβηκε επιτυχώς.")
    else:
        st.stop() # Stop execution until template is uploaded


yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# Define thin border style
thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

# ------------------------------------------------
# Συναρτήσεις ημερομηνιών
# ------------------------------------------------
def parse_date(text, is_start=True):
    text = text.strip()
    if re.match(r"\d{2}/\d{4}", text):
        if is_start:
            return datetime.strptime("01/" + text, "%d/%m/%Y")
        else:
            d = datetime.strptime("01/" + text, "%d/%m/%Y")
            return d + relativedelta(months=1) - relativedelta(days=1)
    else:
        return datetime.strptime(text, "%d/%m/%Y")

def parse_period(p):
    # Replace en-dash with a regular hyphen for consistent splitting
    p_cleaned = p.replace("–", "-")
    a, b = p_cleaned.split("-")
    return parse_date(a, True), parse_date(b, False)

def month_range(start, end):
    current = datetime(start.year, start.month, 1)
    end = datetime(end.year, end.month, 1)
    out = []
    while current <= end:
        out.append((current.year, current.month))
        current += relativedelta(months=1)
    return out

# Function to determine if a color is light or dark (for text readability)
def is_light_color(hex_color):
    hex_color = hex_color.lstrip('#')
    rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    # Calculate luminance (Y = 0.299R + 0.587G + 0.114B)
    luminance = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255
    return luminance > 0.5 # Threshold can be adjusted


# ------------------------------------------------
# Upload Input file in Streamlit
# ------------------------------------------------
uploaded_input_file = st.file_uploader(
    "Ανεβάστε το αρχείο INPUT excel (μόνο 2 στήλες)",
    type=["xlsx"],
    key="input_upload"
)

if uploaded_input_file is not None:
    # Use BytesIO to read the uploaded file directly without saving to disk
    wb_in = openpyxl.load_workbook(uploaded_input_file)
    ws_in = wb_in.active

    headers = {}
    for c in range(1, ws_in.max_column + 1):
        val = str(ws_in.cell(1,c).value).strip()
        headers[val] = c

    if "ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ" not in headers or "ΑΝΘΡΩΠΟΜΗΝΕΣ" not in headers:
        st.error("Το input πρέπει να έχει στήλες: ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ και ΑΝΘΡΩΠΟΜΗΝΕΣ")
        st.stop() # Stop execution if headers are missing

    PERIOD_COL = headers["ΧΡΟΝΙΚΟ ΔΙΑΣΤΗΜΑ"]
    AM_COL = headers["ΑΝΘΡΩΠΟΜΗΝΕΣ"]

    data = []
    all_months = set()

    for r in range(2, ws_in.max_row + 1):
        period = ws_in.cell(r, PERIOD_COL).value
        am_raw = ws_in.cell(r, AM_COL).value
        try:
            am = int(am_raw) if am_raw is not None else 0
        except (ValueError, TypeError):
            am = 0

        if not period:
            continue
        start, end = parse_period(str(period))
        months = month_range(start, end)
        data.append((period, am, months))
        for m in months:
            all_months.add(m)

    all_months = sorted(all_months)
    years = sorted(set(y for y,m in all_months))

    # ------------------------------------------------
    # Ανοίγουμε TEMPLATE
    # ------------------------------------------------
    wb = openpyxl.load_workbook(TEMPLATE_FILE)
    ws = wb.active

    # Adjusted START_ROW, YEAR_ROW, MONTH_ROW
    START_ROW = 4 # Data starts here
    YEAR_ROW = 2  # Years go here
    MONTH_ROW = 3 # Months go here
    START_COL = 5 # Month 1 of first year starts in column E

    # ------------------------------------------------
    # Καθαρισμός παλιάς περιοχής
    # ------------------------------------------------
    merged_cells_to_unmerge = []
    for cell_range_str in list(ws.merged_cells.ranges):
        min_col_mc, min_row_mc, max_col_mc, max_row_mc = openpyxl.utils.cell.range_boundaries(str(cell_range_str))
        if (min_row_mc <= YEAR_ROW <= max_row_mc) or \
           (min_row_mc <= MONTH_ROW <= max_col_mc) or \
           (min_row_mc <= START_ROW <= max_row_mc) or \
           (min_row_mc <= START_ROW + 1 <= max_row_mc):
            merged_cells_to_unmerge.append(cell_range_str)

    for cell_range_str in merged_cells_to_unmerge:
        ws.unmerge_cells(str(cell_range_str))

    max_col_to_clear = max(START_COL + len(years) * 12, ws.max_column + 1)

    rows_to_clear_completely = [YEAR_ROW, START_ROW, START_ROW + 1]

    for r_clear in rows_to_clear_completely:
        for c_clear in range(1, max_col_to_clear):
            ws.cell(r_clear, c_clear).value = None
            ws.cell(r_clear, c_clear).fill = PatternFill()

    ws.cell(MONTH_ROW, 1).value = None
    ws.cell(MONTH_ROW, 1).fill = PatternFill()
    for c_clear in range(START_COL, max_col_to_clear):
        ws.cell(MONTH_ROW, c_clear).value = None
        ws.cell(MONTH_ROW, c_clear).fill = PatternFill()

    for r_clear in range(START_ROW, ws.max_row + 1):
        for c_clear in range(1, max_col_to_clear):
            ws.cell(r_clear,c_clear).value = None
            ws.cell(r_clear,c_clear).fill = PatternFill()

    # ------------------------------------------------
    # Χτίσιμο ετών & μηνών
    # ------------------------------------------------
    col = START_COL
    month_col_map = {}

    for y in years:
        year_start_col = col
        year_header_cell = ws.cell(YEAR_ROW, col)

        r_color_func = lambda: random.randint(0,255)
        random_color_hex = '%02X%02X%02X' % (r_color_func(), r_color_func(), r_color_func())
        year_header_cell.fill = PatternFill(start_color=random_color_hex, end_color=random_color_hex, fill_type="solid")

        if not is_light_color(random_color_hex):
            year_header_cell.font = Font(color="FFFFFF")
        else:
            year_header_cell.font = Font(color="000000")

        for m in range(1,13):
            year_header_cell.value = y
            ws.cell(MONTH_ROW, col).value = m
            month_col_map[(y,m)] = col
            col += 1
        year_end_col = col - 1

        ws.merge_cells(start_row=YEAR_ROW, start_column=year_start_col, end_row=YEAR_ROW, end_column=year_end_col)

    for c_border in range(START_COL, col):
        ws.cell(YEAR_ROW, c_border).border = thin_border
        ws.cell(MONTH_ROW, c_border).border = thin_border

    # ------------------------------------------------
    # Γραμμές & μπάρες
    # ------------------------------------------------
    row = START_ROW

    for period, am, months in data:
        ws.cell(row,2).value = period
        ws.cell(row,2).border = thin_border
        ws.cell(row,3).value = am
        ws.cell(row,3).border = thin_border

        first_month_of_period = True
        for (y,m) in months:
            if (y,m) in month_col_map:
                cell_to_fill = ws.cell(row, month_col_map[(y,m)])
                cell_to_fill.fill = yellow

                if first_month_of_period:
                    cell_to_fill.value = am
                    if am > len(months):
                        cell_to_fill.font = Font(color="FF0000", bold=True)
                    else:
                        cell_to_fill.font = Font(color="000000")
                    first_month_of_period = False

        for c_border in range(START_COL, col):
            ws.cell(row, c_border).border = thin_border

        row += 1

    for c_width in range(START_COL, col):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_width)].width = 2.5

    # ------------------------------------------------
    # Save & download for Streamlit
    # ------------------------------------------------
    output_buffer = BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0) # Rewind the buffer to the beginning

    st.download_button(
        label="Κατεβάστε το επεξεργασμένο αρχείο Excel",
        data=output_buffer,
        file_name="OUTPUT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Παρακαλώ ανεβάστε το αρχείο INPUT excel για να ξεκινήσει η επεξεργασία.")