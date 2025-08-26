# app.py  — Campus Classroom Enrollment fixer
from datetime import datetime
from zoneinfo import ZoneInfo
import re
from typing import Dict, List

import pandas as pd
import streamlit as st
from PIL import Image

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Campus Classroom Enrollment Fixer", layout="centered")

# ----------------------------
# Header
# ----------------------------
try:
    logo = Image.open("header_logo.png")
    st.image(logo, width=240)
except Exception:
    pass

st.title("Head Start — Campus Classroom Enrollment (Fix Class Names + Totals Order)")
st.markdown(
    "- Upload your **Campus Classroom Enrollment** Excel (has columns like *Center, Class, Funded, Enrolled*).\n"
    "- Upload your **VF QuickReport** (the 10422… file that contains `ST: Participant PID / Center / Class`).\n\n"
    "This will:\n"
    "1) Replace placeholder class labels (e.g., **Class 30**) with **lettered class names** from VF (e.g., **E117, A01, B114**), per center.\n"
    "2) Move every **Center Total** row to the **top** of its center section (intro)."
)

main_file = st.file_uploader("Upload Campus Classroom Enrollment Excel", type=["xlsx"], key="main")
vf_file   = st.file_uploader("Upload VF QuickReport (the 10422… file)", type=["xlsx"], key="vf")

# ----------------------------
# Helpers
# ----------------------------
PLACEHOLDER_RE = re.compile(r"^\s*class(room)?\s*30\s*$", re.IGNORECASE)

def find_header_row_by_columns(ws, required_cols: List[str], search_rows: int = 60) -> int | None:
    """
    Find the first row whose values contain all `required_cols` (case-insensitive).
    """
    req = [c.lower() for c in required_cols]
    for r in ws.iter_rows(min_row=1, max_row=search_rows):
        vals = [str(c.value).strip().lower() if c.value is not None else "" for c in r]
        if all(any(reqcol == v for v in vals) for reqcol in req):
            return r[0].row
    return None

def read_main_grid(file) -> pd.DataFrame:
    """
    Read the main Campus Classroom Enrollment grid by discovering a header row that
    contains at least 'Center' and 'Class'.
    """
    wb = load_workbook(file, data_only=True)
    # choose the first sheet that has our columns
    header_row = None
    chosen = None
    for name in wb.sheetnames:
        ws = wb[name]
        header_row = find_header_row_by_columns(ws, ["Center", "Class"])
        if header_row:
            chosen = name
            break
    if not chosen:
        st.error("Couldn’t find a sheet with 'Center' and 'Class' headers in the main workbook.")
        st.stop()
    file.seek(0)
    df = pd.read_excel(file, sheet_name=chosen, header=header_row - 1, dtype=str)
    # keep original dtypes for numeric columns later by re-casting
    return df

def load_vf_center_to_classes(file) -> Dict[str, List[str]]:
    """
    From the VF QuickReport, build: center -> sorted unique list of class names (lettered).
    """
    try:
        wb = load_workbook(file, data_only=True)
        target_sheet = None
        hdr_row = None
        for s in wb.sheetnames:
            ws = wb[s]
            # look for row containing "ST: Participant PID" & "ST: Class Name"
            has_pid = False
            has_class = False
            for r in ws.iter_rows(min_row=1, max_row=160):
                vals = [str(c.value).strip().lower() if c.value is not None else "" for c in r]
                if any("st: participant pid" == v for v in vals): has_pid = True
                if any("st: class name" == v for v in vals): has_class = True
                if has_pid and has_class:
                    hdr_row = r[0].row
                    target_sheet = s
                    break
            if target_sheet:
                break
        if not target_sheet or not hdr_row:
            st.error("VF QuickReport: couldn’t find the required headers (`ST: Participant PID`, `ST: Class Name`).")
            st.stop()
        file.seek(0)
        vf = pd.read_excel(file, sheet_name=target_sheet, header=hdr_row - 1, dtype=str)
        need = {"ST: Center Name", "ST: Class Name"}
        if not need.issubset(vf.columns):
            st.error("VF QuickReport is missing 'ST: Center Name' or 'ST: Class Name'.")
            st.stop()

        vf = vf[list(need)].dropna(subset=["ST: Class Name"]).copy()
        vf["ST: Center Name"] = vf["ST: Center Name"].astype(str).str.strip()
        vf["ST: Class Name"]  = vf["ST: Class Name"].astype(str).str.strip()

        # Build mapping: exact center text -> sorted unique class list
        mapping: Dict[str, List[str]] = {}
        for center, grp in vf.groupby("ST: Center Name"):
            classes = sorted(grp["ST: Class Name"].dropna().unique(), key=lambda x: (re.sub(r"[^A-Za-z0-9]", "", x)))
            mapping[center] = list(classes)
        return mapping
    except Exception as e:
        st.error(f"Couldn’t read VF QuickReport: {e}")
        st.stop()

def replace_placeholders_from_vf(main_df: pd.DataFrame, center_col="Center", class_col="Class",
                                 center_to_classes: Dict[str, List[str]]) -> pd.DataFrame:
    df = main_df.copy()
    # Normalize text
    df[center_col] = df[center_col].astype(str).str.strip()
    df[class_col]  = df[class_col].astype(str).str.strip()

    # For each center we know about, replace placeholder rows in appearance order
    for vf_center, class_list in center_to_classes.items():
        # match rows where center cell contains the vf_center text (case-insensitive)
        mask_center = df[center_col].str.contains(re.escape(vf_center), case=False, na=False)
        # rows with placeholder in Class column
        idxs = df.index[mask_center & df[class_col].str.match(PLACEHOLDER_RE, na=False)].tolist()
        if not idxs:
            continue
        # Determine which class names are not already present in those rows
        assigned = 0
        for i, ridx in enumerate(idxs):
            if assigned >= len(class_list):
                break
            df.at[ridx, class_col] = class_list[assigned]
            assigned += 1
        # If fewer placeholders than class names, that’s fine; if more, remaining stay as-is.
    return df

def move_center_totals_to_top(df: pd.DataFrame, center_col="Center", class_col="Class") -> pd.DataFrame:
    """
    If a row with '<Center> Total' exists, move it to the first row for that center.
    Otherwise, compute a total row and insert it first.
    """
    # Identify numeric columns for summing
    numeric_cols = []
    for c in df.columns:
        if c in [center_col, class_col]:
            continue
        # consider columns that are purely numeric once you drop NaNs/blanks
        try:
            pd.to_numeric(df[c].dropna().replace("", pd.NA))
            numeric_cols.append(c)
        except Exception:
            pass

    out_rows = []
    seen_centers = []
    # Build ordered list of unique centers ignoring rows that look like totals
    base_centers = []
    for val in df[center_col].astype(str):
        v = val.strip()
        if not v:
            continue
        if v.lower().endswith(" total"):
            continue
        if v not in base_centers:
            base_centers.append(v)

    for center in base_centers:
        grp = df[df[center_col] == center].copy()

        # look for existing total row text variations
        total_mask = df[center_col].astype(str).str.strip().str.lower() == f"{center.lower()} total"
        existing_total = df[total_mask].copy()

        if not existing_total.empty:
            total_row = existing_total.iloc[0].copy()
        else:
            # compute a fresh total row
            total_row = pd.Series({col: "" for col in df.columns})
            total_row[center_col] = f"{center} Total"
            for c in numeric_cols:
                # sum numeric with coercion
                total_row[c] = pd.to_numeric(grp[c], errors="coerce").sum(min_count=1)

        # append total row FIRST, then the class records for that center
        out_rows.append(total_row.to_dict())
        out_rows.extend(grp.to_dict(orient="records"))

        # remove the original total row from further processing
        df = df[~total_mask]

    # append any remaining rows that didn’t match pattern (headers, misc)
    remaining = df[~df[center_col].isin(base_centers)]
    out_rows.extend(remaining.to_dict(orient="records"))

    return pd.DataFrame(out_rows, columns=main_df.columns)

def style_basic(ws):
    # Dark blue header, centered
    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill
    # Widths
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 16 if col_idx > 2 else 28

# ----------------------------
# Run
# ----------------------------
if main_file and vf_file:
    # 1) Read main grid
    main_df = read_main_grid(main_file)

    # Make sure the essential columns exist
    if not {"Center", "Class"}.issubset(main_df.columns):
        st.error("Main file must contain 'Center' and 'Class' columns.")
        st.stop()

    # 2) Get (center -> list of classes) from VF
    center_to_classes = load_vf_center_to_classes(vf_file)

    # 3) Replace placeholder class names per center using VF class lists
    main_df = replace_placeholders_from_vf(main_df, center_col="Center", class_col="Class",
                                           center_to_classes=center_to_classes)

    # 4) Move each Center Total row to the TOP (intro) of its center section
    fixed_df = move_center_totals_to_top(main_df, center_col="Center", class_col="Class")

    # 5) Write a clean workbook with simple styling
    out_path = "Campus_Classroom_Enrollment_FIXED.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        fixed_df.to_excel(writer, index=False, sheet_name="Enrollment")
        ws = writer.book["Enrollment"]
        style_basic(ws)

        # Title & timestamp sheet (optional)
        ts = datetime.now(ZoneInfo("America/Chicago")).strftime("%m.%d.%y, %I:%M %p %Z")
        title_ws = writer.book.create_sheet("Info")
        title_ws["A1"] = "Head Start – 2025–2026 Campus Classroom Enrollment"
        title_ws["A2"] = f"As of {ts}"

    with open(out_path, "rb") as f:
        st.download_button("📥 Download: Campus_Classroom_Enrollment_FIXED.xlsx", f, file_name=out_path)

    st.success("Done! Placeholders replaced with VF class names and center totals moved to the top.")


