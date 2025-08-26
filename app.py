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

# =========================
# UI
# =========================
try:
    logo = Image.open("header_logo.png")
    st.image(logo, width=300)
except Exception:
    pass

st.title("HCHSP Enrollment Checklist (2025–2026)")
st.markdown(
    "Upload **Enrollment.xlsx** and the **VF QuickReport** (the 10422… file with a sheet that has "
    "`ST: Participant PID`, `ST: Center Name`, `ST: Class Name`).\n\n"
    "• Fixes repeating **Class 30** by forcing lettered class names from VF.\n"
    "• Keeps the **overall Grand Total at the bottom** of the Enrollment sheet.\n"
    "• On **Center Summary**: Center total appears **first**, and each class shows a **total at the beginning**."
)

enr_file = st.file_uploader("Upload Enrollment.xlsx", type=["xlsx"], key="enr")
vf_file  = st.file_uploader("Upload VF QuickReport (e.g., GEHS_QuickReport)", type=["xlsx"], key="vf")

# =========================
# Helpers
# =========================
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

def find_header_row(ws, probe="ST: Participant PID", search_rows=160):
    for row in ws.iter_rows(min_row=1, max_row=search_rows):
        for cell in row:
            if isinstance(cell.value, str) and probe in cell.value:
                return cell.row
    return None

def pid_norm(series: pd.Series) -> pd.Series:
    # Keep only digits; strip trailing .0; remove leading zeros
    s = series.astype(str).str.replace(r"\.0+$", "", regex=True)
    return s.map(lambda x: re.sub(r"\D+", "", x)).str.lstrip("0")

def find_class_columns(cols):
    # Detect class-like columns; avoid false positives like 'Classification'
    out = []
    for c in cols:
        if not isinstance(c, str):
            continue
        low = c.lower().strip()
        if any(bad in low for bad in ["classification", "class size", "capacity"]):
            continue
        if "class name" in low or "classroom" in low or low == "class" or low.startswith("class "):
            out.append(c)
    # de-dupe while preserving order
    seen, keep = set(), []
    for c in out:
        if c not in seen:
            seen.add(c); keep.append(c)
    return keep

def load_vf_pid_map(vf_file):
    """
    Scan the VF workbook for a sheet that contains the three columns:
    'ST: Participant PID', 'ST: Center Name', 'ST: Class Name'.
    Return a DF indexed by normalized PID with columns:
    ['VF_Class', 'VF_Center'].
    """
    try:
        wb = load_workbook(vf_file, data_only=True)
        for s in wb.sheetnames:
            ws = wb[s]
            hdr = find_header_row(ws, probe="ST: Participant PID", search_rows=160)
            if not hdr:
                continue
            vf_file.seek(0)
            tmp = pd.read_excel(vf_file, sheet_name=s, header=hdr - 1, dtype=str)
            need = {"ST: Participant PID", "ST: Center Name", "ST: Class Name"}
            if not need.issubset(tmp.columns):
                continue
            tmp["PID_norm"] = pid_norm(tmp["ST: Participant PID"])
            tmp = tmp.drop_duplicates(subset=["PID_norm"], keep="last")
            return (tmp.set_index("PID_norm")[["ST: Class Name", "ST: Center Name"]]
                      .rename(columns={"ST: Class Name":"VF_Class", "ST: Center Name":"VF_Center"}))
    except Exception:
        return None
    return None

# =========================
# Main
# =========================
if enr_file:
    # ---- Enrollment load ----
    wb_enr = load_workbook(enr_file, data_only=True)
    ws_enr = wb_enr.active
    enr_hdr = find_header_row(ws_enr)
    if not enr_hdr:
        st.error("Couldn't find 'ST: Participant PID' in Enrollment.xlsx.")
        st.stop()

    enr_file.seek(0)
    df = pd.read_excel(enr_file, header=enr_hdr - 1)

    # normalize column names
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]
    if "Participant PID" not in df.columns:
        st.error("Enrollment file is missing 'Participant PID'.")
        st.stop()

    # keep only the most recent row per PID so columns stay in-sync
    df["PID_norm"] = pid_norm(df["Participant PID"])
    df["_anchor"]  = df.apply(row_anchor_dt, axis=1)
    df = (df.sort_values(["PID_norm", "_anchor"])
            .drop_duplicates(subset=["PID_norm"], keep="last")
            .drop(columns="_anchor"))

    # ---- VF override (force lettered class names) ----
    vf_map = None
    if vf_file:
        vf_map = load_vf_pid_map(vf_file)
        if vf_map is None:
            st.warning("VF file did not include the required PID/Center/Class columns. Skipping VF overrides.")
        else:
            vf_file.seek(0)
            # join by normalized PID index
            df = df.merge(vf_map, left_on="PID_norm", right_index=True, how="left")

            # overwrite all class-like columns with VF class (letters preserved)
            class_cols = find_class_columns(df.columns)
            if "VF_Class" in df.columns:
                has_vf = df["VF_Class"].notna() & (df["VF_Class"].astype(str).str.strip() != "")
                if class_cols:
                    for col in class_cols:
                        df.loc[has_vf, col] = df.loc[has_vf, "VF_Class"]
                else:
                    # create a standard column if none exists
                    df["Class Name"] = df["VF_Class"].where(has_vf, df.get("Class Name"))

            # Prefer VF center text where present
            if "VF_Center" in df.columns:
                for center_col in ["Center Name", "Site", "Campus", "Center"]:
                    if center_col in df.columns:
                        df[center_col] = df["VF_Center"].where(
                            df["VF_Center"].notna() & (df["VF_Center"].astype(str).str.strip() != ""),
                            df[center_col]
                        )
                        break

    # ---- Write Enrollment sheet (title/timestamp + table) ----
    title_text = "Enrollment Checklist 2025–2026"
    central_now = datetime.now(ZoneInfo("America/Chicago"))
    timestamp_text = central_now.strftime("Generated on %B %d, %Y at %I:%M %p %Z")

    out_path = "Enrollment_Cleaned.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame([[title_text]]).to_excel(writer, index=False, header=False, startrow=0, sheet_name="Enrollment")
        pd.DataFrame([[timestamp_text]]).to_excel(writer, index=False, header=False, startrow=1, sheet_name="Enrollment")
        df.to_excel(writer, index=False, startrow=3, sheet_name="Enrollment")  # header row = 4

    # ---- Style Enrollment (overall Grand Total at BOTTOM, unchanged) ----
    wb = load_workbook(out_path)
    ws = wb["Enrollment"]

    filter_row = 4
    data_start = filter_row + 1
    data_end   = ws.max_row
    max_col    = ws.max_column

    # panes and filter
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(max_col)}{data_end}"

    # title/timestamp
    title_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    ts_fill    = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    tcell = ws.cell(row=1, column=1); tcell.value = title_text
    tcell.font = Font(size=14, bold=True); tcell.alignment = Alignment(horizontal="center", vertical="center"); tcell.fill = title_fill
    scell = ws.cell(row=2, column=1); scell.value = timestamp_text
    scell.font = Font(size=10, italic=True, color="555555"); scell.alignment = Alignment(horizontal="center", vertical="center"); scell.fill = ts_fill

    # header styling
    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[filter_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill

    # borders/colors
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"),  bottom=Side(style="thin"))
    red_font    = Font(color="FF0000", bold=True)

    # detect immun & first-name col for totals logic
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

    # cutoffs
    cutoff_date  = datetime(2025, 5, 11)
    immun_cutoff = datetime(2024, 5, 11)

    # cell validation/formatting
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

    # column widths
    width_map = {1: 16, 2: 22}
    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = width_map.get(c, 14)

    # overall Grand Total at bottom (unchanged)
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

    # =========================
    # Center Summary sheet
    #   • Center Total FIRST
    #   • Each Class shows "Class <name> — Total N" FIRST for that class
    # =========================
    # Prefer VF for counts; fallback to Enrollment if needed.
    base = None
    if vf_map is not None:
        base = vf_map.reset_index().rename(columns={"index":"PID_norm"})  # columns: PID_norm, VF_Class, VF_Center
        base = base[["VF_Center", "VF_Class", "PID_norm"]].rename(columns={"VF_Center":"Center", "VF_Class":"Class"})
    else:
        # fallback: try to find center/class columns in Enrollment
        center_col = next((c for c in ["Center Name","Site","Campus","Center"] if c in df.columns), None)
        class_col  = None
        for cc in find_class_columns(df.columns):
            class_col = cc; break
        if center_col and class_col:
            base = df.rename(columns={center_col:"Center", class_col:"Class"})[
                ["Center", "Class", "PID_norm"]
            ]

    if base is not None:
        counts = (
            base.dropna(subset=["Class"])
                .groupby(["Center","Class"])["PID_norm"]
                .nunique()
                .reset_index(name="Students")
        )

        # Build rows with Center-intro and Class-intro at the beginning of each section
        rows = []
        for center_name, grp_c in counts.sort_values(["Center","Class"]).groupby("Center", sort=False):
            # Center intro total
            rows.append({"Center": f"{center_name} — Total", "Class": f"{grp_c['Class'].nunique()} classes", "Students": int(grp_c["Students"].sum())})
            # For each class, put its total FIRST (intro for the class)
            for _, row in grp_c.iterrows():
                rows.append({"Center": center_name, "Class": f"Class {row['Class']} — Total", "Students": int(row["Students"])})
                # (No trailing totals; if you later want to list details per class, insert them after this intro line)
        summary = pd.DataFrame(rows, columns=["Center","Class","Students"])

        # Write/replace sheet
        if "Center Summary" in wb.sheetnames:
            del wb["Center Summary"]
        ws_sum = wb.create_sheet("Center Summary")

        # Header styling
        for j, h in enumerate(["Center","Class","Students"], start=1):
            cell = ws_sum.cell(row=1, column=j, value=h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")

        # Data
        for i, r in enumerate(summary.itertuples(index=False), start=2):
            ws_sum.cell(row=i, column=1, value=r.Center)
            ws_sum.cell(row=i, column=2, value=r.Class)
            ws_sum.cell(row=i, column=3, value=r.Students)

        ws_sum.column_dimensions["A"].width = 42
        ws_sum.column_dimensions["B"].width = 28
        ws_sum.column_dimensions["C"].width = 12

    # ---- Save + Download ----
    final_output = "Formatted_Enrollment_Checklist.xlsx"
    wb.save(final_output)
    with open(final_output, "rb") as f:
        st.download_button("📥 Download Formatted Excel", f, file_name=final_output)

