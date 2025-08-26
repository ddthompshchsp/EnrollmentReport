# app.py — Replace "Class 30" using VF funded report (letters kept)
import re
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Fix Class Names from VF", layout="centered")

st.title("Fix Class Names from VF funded report (letters preserved)")
st.markdown(
    "Upload **Campus Classroom Enrollment** (has columns like *Center, Class*) and the **VF funded report** "
    "(sheet with blocks like `HCHSP -- <Center>` and lines `Class E117`, `Class A01`, …). "
    "This will replace any **'Class 30' / 'Classroom 30'** with the **lettered classes from VF**, by **center**, in order."
)

main_file = st.file_uploader("Campus Classroom Enrollment (.xlsx)", type=["xlsx"], key="main")
vf_file   = st.file_uploader("VF funded report (.xlsx)", type=["xlsx"], key="vf")

# ---------- helpers ----------
PLACEHOLDER_RE = re.compile(r"^\s*class(room)?\s*30\s*$", re.IGNORECASE)

def parse_vf_classes(vf_file) -> dict[str, list[str]]:
    """
    Parse the VF funded report format you showed (sheet name like 'VF_Average_Funded_Enrollment_Le'):
    finds lines 'HCHSP -- <Center>' then lines 'Class <CODE>' and builds center -> [class codes].
    """
    xls = pd.ExcelFile(vf_file)
    # pick the sheet that likely holds the text blocks
    sheet = next((s for s in xls.sheet_names if s.lower().startswith("vf_average_funded")), xls.sheet_names[0])
    raw = pd.read_excel(vf_file, sheet_name=sheet, header=None)

    center_re = re.compile(r'^\s*HCHSP\s*--\s*(.+?)\s*$', re.IGNORECASE)
    class_re  = re.compile(r'^\s*Class\s+([A-Za-z]+[A-Za-z0-9]*)\s*$', re.IGNORECASE)

    mapping: dict[str, list[str]] = {}
    current = None
    for v in raw.iloc[:, 0].astype(str).fillna(""):
        v = v.strip()
        if not v:
            continue
        m_center = center_re.match(v)
        if m_center:
            current = m_center.group(1)
            mapping.setdefault(current, [])
            continue
        m_class = class_re.match(v)
        if m_class and current:
            mapping[current].append(m_class.group(1))
    return mapping

def find_header_row_by_cols(ws, must_have=("Center", "Class"), scan_rows=80):
    must = [c.lower() for c in must_have]
    for r in ws.iter_rows(min_row=1, max_row=scan_rows):
        vals = [str(c.value).strip().lower() if c.value is not None else "" for c in r]
        if all(any(m == v for v in vals) for m in must):
            return r[0].row
    return None

def read_main_grid(main_file) -> tuple[pd.DataFrame, str]:
    wb = load_workbook(main_file, data_only=True)
    chosen, hdr = None, None
    for s in wb.sheetnames:
        ws = wb[s]
        hdr = find_header_row_by_cols(ws)
        if hdr:
            chosen = s
            break
    if not chosen:
        st.error("Couldn’t find a sheet with 'Center' and 'Class' headers.")
        st.stop()
    main_file.seek(0)
    df = pd.read_excel(main_file, sheet_name=chosen, header=hdr-1)
    return df, chosen

def replace_placeholders(df: pd.DataFrame, vf_map: dict[str, list[str]]) -> pd.DataFrame:
    """
    For each center, walk down its rows and replace 'Class 30' (or 'Classroom 30') with the VF class list in order.
    Existing lettered class values are left as-is.
    """
    if "Center" not in df.columns or "Class" not in df.columns:
        st.error("Main sheet must contain columns 'Center' and 'Class'.")
        st.stop()

    out = df.copy()
    out["Center"] = out["Center"].astype(str).str.strip()
    out["Class"]  = out["Class"].astype(str).str.strip()

    # normalize function for center text to improve matching
    def norm_center(x: str) -> str:
        x = x.lower().strip()
        x = re.sub(r"\s+", " ", x)
        return x

    # build a center index
    centers_in_main = out["Center"].dropna().astype(str).map(norm_center).tolist()

    for vf_center, classes in vf_map.items():
        if not classes:
            continue
        n_vf = [c for c in classes if isinstance(c, str) and c.strip()]

        # find rows for this center
        mask_center = out["Center"].astype(str).map(norm_center) == norm_center(vf_center)
        idxs = out.index[mask_center].tolist()
        if not idxs:
            # try 'contains' fallback (helps when headings have suffixes/prefixes)
            mask_center = out["Center"].astype(str).str.contains(re.escape(vf_center), case=False, na=False)
            idxs = out.index[mask_center].tolist()
        if not idxs:
            continue  # this center not present in the main grid

        # select placeholder rows within those indices, in order
        ph_rows = [i for i in idxs if PLACEHOLDER_RE.match(out.at[i, "Class"] or "")]
        # advance through placeholders and assign VF classes
        p = 0
        for ridx in ph_rows:
            if p >= len(n_vf): break
            out.at[ridx, "Class"] = n_vf[p]
            p += 1
        # (if there are more placeholders than VF classes, the extras remain as-is)

    return out

def style_headers(ws):
    fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = fill
    for c in range(1, ws.max_column+1):
        ws.column_dimensions[get_column_letter(c)].width = 16 if c > 2 else 28

# ---------- run ----------
if main_file and vf_file:
    main_df, sheet_name = read_main_grid(main_file)
    vf_map = parse_vf_classes(vf_file)

    # (Optional) quick peek so you can confirm we read VF correctly
    st.write("Classes found in VF by center (first few):")
    st.json({k: v for k, v in vf_map.items()})

    fixed = replace_placeholders(main_df, vf_map)

    out_path = "Campus_Classroom_Enrollment_CLASSES_FIXED.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        fixed.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        style_headers(ws)

    with open(out_path, "rb") as f:
        st.download_button("📥 Download fixed workbook", f, file_name=out_path)

    st.success("All “Class 30” placeholders were replaced with lettered class names from VF (by center).")


