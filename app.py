# app.py
from datetime import datetime, date
from zoneinfo import ZoneInfo
import re

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
st.markdown(
    "Upload **Enrollment.xlsx** (main export) and a **VF QuickReport** (e.g., file with 'GEHS_QuickReport'). "
    "This fixes repeating 'Classroom 30' and puts each campus Total at the top of its section."
)

enr_file = st.file_uploader("Upload Enrollment.xlsx", type=["xlsx"], key="enr")
vf_file  = st.file_uploader("Upload VF QuickReport (has ST: Participant PID / Center / Class)", type=["xlsx"], key="vf")

# ----------------------------
# Helpers
# ----------------------------
def coerce_to_dt(v):
    if pd.isna(v):
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return from_excel(v)
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
    dates = []
    for _, val in row.items():
        dt = coerce_to_dt(val)
        if dt:
            dates.append(dt)
    return max(dates) if dates else datetime.min

def find_header_row(ws, probe="ST: Participant PID", search_rows=80):
    for row in ws.iter_rows(min_row=1, max_row=search_rows):
        for cell in row:
            if isinstance(cell.value, str) and probe in cell.value:
                return cell.row
    return None

def norm_pid_series(s):
    # CHANGE A(1): Normalize PID dtype/string so merges line up (prevents 'classroom 30' sticking)
    return s.astype(str).map(lambda x: re.sub(r"\D+", "", x)).str.lstrip("0")

def find_class_columns(cols):
    # Try to catch Class Name / Classroom / Class etc., but avoid 'Classification' etc.
    out = []
    for c in cols:
        if not isinstance(c, str):
            continue
        low = c.lower().strip()
        if "classification" in low:
            continue
        if "class name" in low or "classroom" in low:
            out.append(c)
        elif low == "class" or low.startswith("class "):
            out.append(c)
    return out

def try_load_vf_pid_map(xlsx_file):
    # Look for any sheet with PID/Center/Class columns
    try:
        wb = load_workbook(xlsx_file, data_only=True)
        for s in wb.sheetnames:
            ws = wb[s]
            hdr = find_header_row(ws, probe="ST: Participant PID", search_rows=120)
            if not hdr:
                continue
            xlsx_file.seek(0)
            tmp = pd.read_excel(xlsx_file, sheet_name=s, header=hdr - 1)
            needed = {"ST: Participant PID", "ST: Center Name", "ST: Class Name"}
            if needed.issubset(tmp.columns):
                vf = (tmp[list(needed)]
                      .dropna(subset=["ST: Participant PID"])
                      .rename(columns={
                          "ST: Participant PID": "Participant PID",
                          "ST: Center Name":     "Center Name (VF)",
                          "ST: Class Name":      "Class Name (VF)",
                      }))
                # Normalize PID to string-digits
                vf["Participant PID"] = norm_pid_series(vf["Participant PID"])
                vf = vf.drop_duplicates(subset=["Participant PID"], keep="last")
                return vf
    except Exception:
        return None
    return None

# ----------------------------
# Main
# ----------------------------
if enr_file:
    # 1) Parse Enrollment.xlsx
    wb_src = load_workbook(enr_file, data_only=True)
    ws_src = wb_src.active
    header_row = find_header_row(ws_src)
    if not header_row:
        st.error("Couldn't find 'ST: Participant PID' in Enrollment.xlsx.")
        st.stop()

    enr_file.seek(0)
    df = pd.read_excel(enr_file, header=header_row - 1)
    # Drop 'ST: ' prefix for convenience
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]

    if "Participant PID" not in df.columns:
        st.error("Enrollment file is missing 'Participant PID'.")
        st.stop()

    # Normalize PID dtype
    df["Participant PID"] = norm_pid_series(df["Participant PID"])

    # Keep the most-recent row per PID (prevents per-column mixing like 'Classroom 30')
    df["_anchor"] = df.apply(row_anchor_dt, axis=1)
    df = (df.sort_values(["Participant PID", "_anchor"])
            .drop_duplicates(subset=["Participant PID"], keep="last")
            .drop(columns="_anchor"))

    # 2) Load VF map (optional but recommended) and FORCE override class values
    vf_map = None
    if vf_file:
        vf_map = try_load_vf_pid_map(vf_file)
        if vf_map is None:
            st.info("VF file did not include PID/Center/Class columns. Skipping VF overrides.")
        else:
            vf_file.seek(0)
            # Merge by normalized PID
            df = df.merge(vf_map, on="Participant PID", how="left")

            # CHANGE A(2): Force class override into *any* class-like columns present
            class_cols = find_class_columns(df.columns)
            if "Class Name (VF)" in df.columns:
                if class_cols:
                    for col in class_cols:
                        # Overwrite class in Enrollment wherever VF has a value
                        mask = df["Class Name (VF)"].notna() & (df["Class Name (VF)"].astype(str).str.strip() != "")
                        df.loc[mask, col] = df.loc[mask, "Class Name (VF)"]
                else:
                    # If no class column exists, create a standard one
                    df["Class Name"] = df["Class Name (VF)"]

            # Optionally override center name if you want (lighter touch; combine_first)
            if "Center Name (VF)" in df.columns:
                if "Center Name" in df.columns:
                    df["Center Name"] = df["Center Name (VF)"].combine_first(df["Center Name"])
                elif "Site" in df.columns:
                    df["Site"] = df["Center Name (VF)"].combine_first(df["Site"])
                else:
                    df["Center Name"] = df["Center Name (VF)"]

            # Drop helper columns
            drop_cols = [c for c in ["Center Name (VF)", "Class Name (VF)"] if c in df.columns]
            if drop_cols:
                df.drop(columns=drop_cols, inplace=True)

    # 3) Write workbook: title/timestamp + data
    title_text = "Enrollment Checklist 2025–2026"
    central_now = datetime.now(ZoneInfo("America/Chicago"))
    timestamp_text = central_now.strftime("Generated on %B %d, %Y at %I:%M %p %Z")

    temp_path = "Enrollment_Cleaned.xlsx"
    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        pd.DataFrame([[title_text]]).to_excel(writer, index=False, header=False, startrow=0, sheet_name="Enrollment")
        pd.DataFrame([[timestamp_text]]).to_excel(writer, index=False, header=False, startrow=1, sheet_name="Enrollment")
        df.to_excel(writer, index=False, startrow=3, sheet_name="Enrollment")  # header at row 4

    # 4) Style Enrollment (unchanged except for your two requests)
    wb = load_workbook(temp_path)
    ws = wb["Enrollment"]

    filter_row = 4
    data_start = filter_row + 1
    data_end = ws.max_row
    max_col = ws.max_column

    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(max_col)}{data_end}"

    title_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    ts_fill    = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)

    tcell = ws.cell(row=1, column=1); tcell.value = title_text
    tcell.font = Font(size=14, bold=True); tcell.alignment = Alignment(horizontal="center", vertical="center")
    tcell.fill = title_fill

    scell = ws.cell(row=2, column=1); scell.value = timestamp_text
    scell.font = Font(size=10, italic=True, color="555555"); scell.alignment = Alignment(horizontal="center", vertical="center")
    scell.fill = ts_fill

    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[filter_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill

    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"),  bottom=Side(style="thin"))
    red_font = Font(color="FF0000", bold=True)

    headers = [ws.cell(row=filter_row, column=c).value for c in range(1, max_col + 1)]
    immun_col = None
    name_col_idx = None
    for idx, h in enumerate(headers, start=1):
        if isinstance(h, str):
            low = h.lower()
            if immun_col is None and "immun" in low:
                immun_col = idx
            if name_col_idx is None and "name" in low:
                name_col_idx = idx
    if name_col_idx is None:
        name_col_idx = 2

    cutoff_date  = datetime(2025, 5, 11)
    immun_cutoff = datetime(2024, 5, 11)

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "filtered total" in v.lower():
                ws.cell(row=r, column=c).value = None

    for r in range(data_start, data_end + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val  = cell.value
            cell.border = thin_border

            if val in (None, "", "nan", "NaT"):
                cell.value = "X"; cell.font = red_font
                continue
            dt = coerce_to_dt(val)
            if dt:
                if c == immun_col and dt < immun_cutoff:
                    cell.value = dt; cell.number_format = "m/d/yy"; cell.font = red_font
                elif dt < cutoff_date:
                    cell.value = "X"; cell.font = red_font
                else:
                    cell.value = dt; cell.number_format = "m/d/yy"
                continue
            # non-date text stays as-is (including the forced class names)

    width_map = {1: 16, 2: 22}
    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = width_map.get(c, 14)

    # Overall Grand Total stays at the bottom (unchanged)
    total_row = ws.max_row + 2
    ws.cell(row=total_row, column=1, value="Grand Total").font = Font(bold=True)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    center = Alignment(horizontal="center", vertical="center")
    for c in range(1, max_col + 1):
        if c <= name_col_idx:
            continue
        valid_count = 0
        for r in range(data_start, data_end + 1):
            if ws.cell(row=r, column=c).value != "X":
                valid_count += 1
        cell = ws.cell(row=total_row, column=c, value=valid_count)
        cell.alignment = center; cell.font = Font(bold=True)
        cell.border = Border(top=Side(style="thin"))

    # 5) Center Summary sheet: Campus Total at TOP of each campus section
    base_for_counts = None
    if vf_file and vf_map is not None and {"Participant PID","Center Name (VF)","Class Name (VF)"}.issubset(vf_map.columns):
        base_for_counts = vf_map.rename(columns={"Center Name (VF)":"Center","Class Name (VF)":"Class"})[
            ["Participant PID","Center","Class"]
        ]
    else:
        # Fallback: use Enrollment's class/center columns if present
        cand_center = None
        for cn in ["Center Name", "Site", "Campus", "Center"]:
            if cn in df.columns:
                cand_center = cn; break
        cand_class = None
        for cc in find_class_columns(df.columns):
            cand_class = cc; break
        if cand_center and cand_class:
            base_for_counts = df.rename(columns={cand_center:"Center", cand_class:"Class"})[
                ["Participant PID","Center","Class"]
            ]

    if base_for_counts is not None:
        class_counts = (
            base_for_counts.dropna(subset=["Class"])
                           .groupby(["Center","Class"])["Participant PID"]
                           .nunique()
                           .reset_index(name="Students")
        )
        rows = []
        for center_name, grp in class_counts.sort_values(["Center","Class"]).groupby("Center", sort=False):
            rows.append({
                "Center":   f"{center_name} — Total",
                "Class":    f"{grp['Class'].nunique()} classes",
                "Students": int(grp["Students"].sum())
            })
            rows.extend(grp.to_dict("records"))
        summary = pd.DataFrame(rows, columns=["Center","Class","Students"])

        if "Center Summary" in wb.sheetnames:
            del wb["Center Summary"]
        ws_sum = wb.create_sheet("Center Summary")

        # Write headers
        for j, h in enumerate(["Center","Class","Students"], start=1):
            cell = ws_sum.cell(row=1, column=j, value=h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")

        # Write data
        for i, row in enumerate(summary.itertuples(index=False), start=2):
            ws_sum.cell(row=i, column=1, value=row.Center)
            ws_sum.cell(row=i, column=2, value=row.Class)
            ws_sum.cell(row=i_


