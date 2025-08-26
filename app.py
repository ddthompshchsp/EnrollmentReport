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
    "Upload **Enrollment.xlsx** and your **VF QuickReport** (the 10422… file with the `GEHS_QuickReport` sheet). "
    "This will fix repeating 'Classroom 30' by forcing lettered class names from VF, and put each campus Total at the TOP "
    "of the Center Summary tab (overall Grand Total stays at the bottom of the main sheet)."
)

enr_file = st.file_uploader("Upload Enrollment.xlsx", type=["xlsx"], key="enr")
vf_file  = st.file_uploader("Upload VF QuickReport (must include ST: Participant PID / Center / Class)", type=["xlsx"], key="vf")

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

def find_header_row(ws, probe="ST: Participant PID", search_rows=120):
    for row in ws.iter_rows(min_row=1, max_row=search_rows):
        for cell in row:
            if isinstance(cell.value, str) and probe in cell.value:
                return cell.row
    return None

def norm_pid(s):
    # digit-only, strip leading zeros; also return raw-as-string for double-key matching
    s_str = s.astype(str)
    raw = s_str.str.replace(r"\.0+$", "", regex=True)
    norm = s_str.map(lambda x: re.sub(r"\D+", "", x)).str.lstrip("0")
    return raw, norm

def find_class_columns(cols):
    # Capture Class Name / Classroom / Class, but NOT 'Classification'
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
    """
    Returns a DataFrame with:
      ['pid_raw','pid_norm','Center Name (VF)','Class Name (VF)']
    by scanning sheets for the GEHS_QuickReport-style header row.
    """
    try:
        wb = load_workbook(xlsx_file, data_only=True)
        for s in wb.sheetnames:
            ws = wb[s]
            hdr = find_header_row(ws, probe="ST: Participant PID", search_rows=160)
            if not hdr:
                continue
            xlsx_file.seek(0)
            tmp = pd.read_excel(xlsx_file, sheet_name=s, header=hdr - 1)
            need = {"ST: Participant PID", "ST: Center Name", "ST: Class Name"}
            if not need.issubset(tmp.columns):
                continue

            # Keep exact text for Class Name (letters like E117, A01)
            pid_raw, pid_norm = norm_pid(tmp["ST: Participant PID"])
            vf = pd.DataFrame({
                "pid_raw":  pid_raw,
                "pid_norm": pid_norm,
                "Center Name (VF)": tmp["ST: Center Name"].astype(str),
                "Class Name (VF)":  tmp["ST: Class Name"].astype(str),
            })
            # last occurrence wins
            vf = vf.drop_duplicates(subset=["pid_norm"], keep="last")
            vf = vf.drop_duplicates(subset=["pid_raw"],  keep="last")
            return vf
    except Exception:
        return None
    return None

# ----------------------------
# Main
# ----------------------------
if enr_file:
    # ---- Load Enrollment.xlsx ----
    wb_enr = load_workbook(enr_file, data_only=True)
    ws_enr = wb_enr.active
    enr_hdr = find_header_row(ws_enr)
    if not enr_hdr:
        st.error("Couldn't find 'ST: Participant PID' in Enrollment.xlsx.")
        st.stop()

    enr_file.seek(0)
    df = pd.read_excel(enr_file, header=enr_hdr - 1)
    # Normalize header names (drop 'ST: ')
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]

    if "Participant PID" not in df.columns:
        st.error("Enrollment file is missing 'Participant PID'.")
        st.stop()

    # Keep single most-recent row per PID (prevents per-column mixing)
    pid_raw_enr, pid_norm_enr = norm_pid(df["Participant PID"])
    df["pid_raw"]  = pid_raw_enr
    df["pid_norm"] = pid_norm_enr
    df["_anchor"]  = df.apply(row_anchor_dt, axis=1)
    df = (
        df.sort_values(["pid_norm", "_anchor"])
          .drop_duplicates(subset=["pid_norm"], keep="last")
          .drop(columns="_anchor")
    )

    # ---- Load VF QuickReport & FORCE override classes (letters preserved) ----
    vf_map = None
    if vf_file:
        vf_map = try_load_vf_pid_map(vf_file)
        if vf_map is None:
            st.warning("VF file didn’t contain the PID/Center/Class columns. Skipping VF overrides.")
        else:
            vf_file.seek(0)

            # Merge on BOTH keys (pid_norm first, then use pid_raw as fallback)
            df = df.merge(vf_map[["pid_norm","Class Name (VF)","Center Name (VF)"]],
                          on="pid_norm", how="left")
            df = df.merge(vf_map[["pid_raw","Class Name (VF)","Center Name (VF)"]]
                          .rename(columns={
                              "Class Name (VF)": "Class Name (VF)_raw",
                              "Center Name (VF)": "Center Name (VF)_raw"
                          }),
                          on="pid_raw", how="left")

            # Resolve VF class/center from either match
            df["__class_vf"] = df["Class Name (VF)"].fillna(df["Class Name (VF)_raw"])
            df["__center_vf"] = df["Center Name (VF)"].fillna(df["Center Name (VF)_raw"])

            # Find class-like columns to overwrite
            class_cols = find_class_columns(df.columns)

            # Always override any 'Class 30' / 'Classroom 30' or any stale class where VF has a value
            if "__class_vf" in df.columns:
                mask_vf = df["__class_vf"].notna() & (df["__class_vf"].astype(str).str.strip() != "")
                if class_cols:
                    for col in class_cols:
                        # Hard override when VF has a value
                        df.loc[mask_vf, col] = df.loc[mask_vf, "__class_vf"].astype(str)
                else:
                    df["Class Name"] = df["__class_vf"].where(mask_vf, df.get("Class Name"))

            # Prefer VF center where present, but don't wipe existing non-null values
            if "__center_vf" in df.columns:
                if "Center Name" in df.columns:
                    df["Center Name"] = df["__center_vf"].where(
                        df["__center_vf"].notna() & (df["__center_vf"].astype(str).str.strip() != ""),
                        df["Center Name"]
                    )
                elif "Site" in df.columns:
                    df["Site"] = df["__center_vf"].where(
                        df["__center_vf"].notna() & (df["__center_vf"].astype(str).str.strip() != ""),
                        df["Site"]
                    )
                else:
                    df["Center Name"] = df["__center_vf"]

            # Tidy helper cols
            drop_cols = [c for c in ["Class Name (VF)","Center Name (VF)","Class Name (VF)_raw","Center Name (VF)_raw","__class_vf","__center_vf"] if c in df.columns]
            if drop_cols:
                df.drop(columns=drop_cols, inplace=True)

    # ----------------------------
    # Write workbook (title/timestamp + data)
    # ----------------------------
    title_text = "Enrollment Checklist 2025–2026"
    central_now = datetime.now(ZoneInfo("America/Chicago"))
    timestamp_text = central_now.strftime("Generated on %B %d, %Y at %I:%M %p %Z")

    out_path = "Enrollment_Cleaned.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame([[title_text]]).to_excel(writer, index=False, header=False, startrow=0, sheet_name="Enrollment")
        pd.DataFrame([[timestamp_text]]).to_excel(writer, index=False, header=False, startrow=1, sheet_name="Enrollment")
        df.to_excel(writer, index=False, startrow=3, sheet_name="Enrollment")  # header row = 4

    # ----------------------------
    # Style Enrollment sheet (overall Grand Total stays at BOTTOM)
    # ----------------------------
    wb = load_workbook(out_path)
    ws = wb["Enrollment"]

    filter_row = 4
    data_start = filter_row + 1
    data_end = ws.max_row
    max_col = ws.max_column

    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(max_col)}{data_end}"

    # Title / Timestamp
    title_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    ts_fill    = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    tcell = ws.cell(row=1, column=1); tcell.value = title_text
    tcell.font = Font(size=14, bold=True); tcell.alignment = Alignment(horizontal="center", vertical="center"); tcell.fill = title_fill
    scell = ws.cell(row=2, column=1); scell.value = timestamp_text
    scell.font = Font(size=10, italic=True, color="555555"); scell.alignment = Alignment(horizontal="center", vertical="center"); scell.fill = ts_fill

    # Header
    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[filter_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill

    # Borders / colors
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

    # Column widths
    width_map = {1: 16, 2: 22}
    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = width_map.get(c, 14)

    # Overall Grand Total (BOTTOM — unchanged)
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

    # ----------------------------
    # Center Summary sheet (campus TOTAL at TOP, then classes)
    # ----------------------------
    base = None
    if vf_map is not None:
        # Use VF for counts
        base = vf_map.rename(columns={"Center Name (VF)":"Center","Class Name (VF)":"Class"})[
            ["pid_norm","Center","Class"]
        ]
    else:
        # Fallback to whatever Enrollment calls center/class
        center_col = next((c for c in ["Center Name","Site","Campus","Center"] if c in df.columns), None)
        class_col  = None
        for cc in find_class_columns(df.columns):
            class_col = cc; break
        if center_col and class_col:
            base = df.rename(columns={center_col:"Center", class_col:"Class"})[
                ["pid_norm","Center","Class"]
            ]

    if base is not None:
        class_counts = (
            base.dropna(subset=["Class"])
                .groupby(["Center","Class"])["pid_norm"]
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

        # Header
        for j, h in enumerate(["Center","Class","Students"], start=1):
            cell = ws_sum.cell(row=1, column=j, value=h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")

        # Data
        for i, row in enumerate(summary.itertuples(index=False), start=2):
            ws_sum.cell(row=i, column=1, value=row.Center)
            ws_sum.cell(row=i, column=2, value=row.Class)
            ws_sum.cell(row=i, column=3, value=row.Students)

        ws_sum.column_dimensions["A"].width = 40
        ws_sum.column_dimensions["B"].width = 22
        ws_sum.column_dimensions["C"].width = 12

    # Save + download
    final_output = "Formatted_Enrollment_Checklist.xlsx"
    wb.save(final_output)
    with open(final_output, "rb") as f:
        st.download_button("📥 Download Formatted Excel", f, file_name=final_output)
