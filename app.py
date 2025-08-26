# app.py
from datetime import datetime, date
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from PIL import Image

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

st.set_page_config(page_title="Enrollment Formatter", layout="centered")

# ----------------------------
# Header / UI
# ----------------------------
try:
    logo = Image.open("header_logo.png")
    st.image(logo, width=300)
except Exception:
    pass

st.title("HCHSP Enrollment Checklist (2025–2026)")
st.markdown("Upload your **Enrollment.xlsx** file to receive a formatted version.")

uploaded_file = st.file_uploader("Upload Enrollment.xlsx", type=["xlsx"])

# ----------------------------
# Helpers
# ----------------------------
def coerce_to_dt(v):
    """Try to convert mixed Excel/str/number inputs into datetime."""
    if pd.isna(v):
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return from_excel(v)  # Excel serial number
        except Exception:
            return None
    if isinstance(v, str):
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip(), fmt)
            except Exception:
                continue
    return None

def row_anchor_dt(row):
    """
    Build an 'anchor' timestamp for the row by scanning all values and
    taking the max datetime we can parse. If no dates, return datetime.min.
    This lets us pick the truly latest row per PID so the 'Class' text
    comes from that same latest row.
    """
    dates = []
    for _, val in row.items():
        dt = coerce_to_dt(val)
        if dt:
            dates.append(dt)
    return max(dates) if dates else datetime.min

if uploaded_file:
    # ----------------------------
    # 1) Find the header row
    # ----------------------------
    wb_src = load_workbook(uploaded_file, data_only=True)
    ws_src = wb_src.active

    header_row = None
    for row in ws_src.iter_rows(min_row=1, max_row=30):
        for cell in row:
            if isinstance(cell.value, str) and "ST: Participant PID" in cell.value:
                header_row = cell.row
                break
        if header_row:
            break

    if not header_row:
        st.error("Couldn't find 'ST: Participant PID' in the file.")
        st.stop()

    uploaded_file.seek(0)

    # ----------------------------
    # 2) Load table into pandas
    # ----------------------------
    df = pd.read_excel(uploaded_file, header=header_row - 1)

    # Normalize column names (drop the 'ST: ' prefix)
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]

    if "Participant PID" not in df.columns:
        st.error("The file is missing 'Participant PID'.")
        st.stop()

    # ====== KEY CHANGE #1: keep the single most-recent row per PID ======
    df = df.dropna(subset=["Participant PID"]).copy()
    df["__anchor"] = df.apply(row_anchor_dt, axis=1)
    df = (
        df.sort_values(["Participant PID", "__anchor"])
          .drop_duplicates(subset=["Participant PID"], keep="last")
          .drop(columns="__anchor")
    )
    # ====================================================================

    # ----------------------------
    # 3) Write temp workbook
    # ----------------------------
    title_text = "Enrollment Checklist 2025–2026"
    central_now = datetime.now(ZoneInfo("America/Chicago"))
    timestamp_text = central_now.strftime("Generated on %B %d, %Y at %I:%M %p %Z")

    temp_path = "Enrollment_Cleaned.xlsx"
    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        # We'll overwrite these cells during styling (merging/centering),
        # but writing placeholders keeps rows allocated.
        pd.DataFrame([[title_text]]).to_excel(writer, index=False, header=False, startrow=0)
        pd.DataFrame([[timestamp_text]]).to_excel(writer, index=False, header=False, startrow=1)
        df.to_excel(writer, index=False, startrow=3)  # header row initially at 4

    # ----------------------------
    # 4) Style with openpyxl
    # ----------------------------
    wb = load_workbook(temp_path)
    ws = wb.active

    # ====== KEY CHANGE #2: insert a Grand Total row BEFORE the table ======
    # Insert a blank row at row 3, pushing header to row 5 and data to row 6+
    totals_row = 3
    ws.insert_rows(totals_row, amount=1)
    # After insertion:
    filter_row = 5               # header row of the data table
    data_start = filter_row + 1  # first data row
    data_end = ws.max_row        # last data row (before we add anything else)
    max_col = ws.max_column
    # ======================================================================

    # Freeze panes so PID (col A) and header row (row 5) stay visible:
    ws.freeze_panes = "B6"

    # AutoFilter on the table
    ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(max_col)}{data_end}"

    # ==== Title + timestamp styling (MERGED & CENTERED) ====
    title_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    ts_fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")

    # Merge across all columns currently present
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)

    # Title cell
    tcell = ws.cell(row=1, column=1)
    tcell.value = title_text
    tcell.font = Font(size=14, bold=True)
    tcell.alignment = Alignment(horizontal="center", vertical="center")
    tcell.fill = title_fill

    # Timestamp cell
    scell = ws.cell(row=2, column=1)
    scell.value = timestamp_text
    scell.font = Font(size=10, italic=True, color="555555")
    scell.alignment = Alignment(horizontal="center", vertical="center")
    scell.fill = ts_fill

    # ==== Header styling (dark blue + white bold + wrap text) ====
    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[filter_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill

    # Borders & fonts
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    red_font = Font(color="FF0000", bold=True)

    # Identify key columns
    immun_col = None
    name_col_idx = None
    headers = [ws.cell(row=filter_row, column=c).value for c in range(1, max_col + 1)]
    for idx, h in enumerate(headers, start=1):
        if isinstance(h, str):
            low = h.lower()
            if immun_col is None and "immun" in low:
                immun_col = idx
            if name_col_idx is None and "name" in low:
                name_col_idx = idx
    if name_col_idx is None:
        name_col_idx = 2  # fallback if no "Name" header

    # Cutoffs
    cutoff_date = datetime(2025, 5, 11)
    immun_cutoff = datetime(2024, 5, 11)

    # Remove any stray "Filtered Total: ..." wording anywhere (just in case)
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "filtered total" in v.lower():
                ws.cell(row=r, column=c).value = None

    # Validate & format data cells
    for r in range(data_start, data_end + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            cell.border = thin_border

            if val in (None, "", "nan", "NaT"):
                cell.value = "X"
                cell.font = red_font
                continue

            dt = coerce_to_dt(val)
            if dt:
                # Immunization special rule: keep dates before 5/11/2024 in red
                if c == immun_col and dt < immun_cutoff:
                    cell.value = dt
                    cell.number_format = "m/d/yy"
                    cell.font = red_font
                # General rule: before cutoff -> X
                elif dt < cutoff_date:
                    cell.value = "X"
                    cell.font = red_font
                else:
                    cell.value = dt
                    cell.number_format = "m/d/yy"
                continue
            # non-date: leave as-is

    # Column widths
    width_map = {1: 16, 2: 22}
    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = width_map.get(c, 14)

    # ----------------------------
    # 5) Grand Total row (placed BEFORE the table, at row 3)
    # ----------------------------
    ws.cell(row=totals_row, column=1, value="Grand Total").font = Font(bold=True)
    ws.cell(row=totals_row, column=1).alignment = Alignment(horizontal="left", vertical="center")

    center = Alignment(horizontal="center", vertical="center")
    for c in range(1, max_col + 1):
        if c <= name_col_idx:
            continue

        valid_count = 0
        for r in range(data_start, data_end + 1):
            if ws.cell(row=r, column=c).value != "X":
                valid_count += 1

        cell = ws.cell(row=totals_row, column=c, value=valid_count)
        cell.alignment = center
        cell.font = Font(bold=True)
        # optional subtle divider above totals
        cell.border = Border(bottom=Side(style="thin"))

    # ----------------------------
    # 6) Save and download
    # ----------------------------
    final_output = "Formatted_Enrollment_Checklist.xlsx"
    wb.save(final_output)

    with open(final_output, "rb") as f:
        st.download_button("📥 Download Formatted Excel", f, file_name=final_output)



