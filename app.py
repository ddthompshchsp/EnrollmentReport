# app.py — Force class names from VF funded report (letters preserved)
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
    "1) Upload your **Campus Classroom Enrollment** workbook (must contain columns **Center** and **Class**).\n"
    "2) Upload the **VF funded report** (sheet with blocks like `HCHSP -- <Center>` and lines `Class E117`, `Class A01`, …).\n\n"
    "**Result:** The **Class** column is replaced per center with the exact, lettered classes from the VF funded report.\n"
    "No row reordering, no totals moved — just the class names fixed."
)

main_file = st.file_uploader("Campus Classroom Enrollment (.xlsx)", type=["xlsx"], key="main")
vf_file   = st.file_uploader("VF funded report (.xlsx)", type=["xlsx"], key="vf")

# ---------------- helpers ----------------
PLACEHOLDER_ANY = re.compile(r"^\s*class(?:room)?\s*\d+\s*$", re.IGNORECASE)  # Class 30, Classroom 09, etc.
                           # 101, 030, etc.
CLASS_TOTALS    = re.compile(r"^\s*class\s+totals\s*:?\s*$", re.IGNORECASE)

def parse_vf_classes(vf_file) -> dict[str, list[str]]:
    """
    Parse the VF funded report: looks for lines 'HCHSP -- <Center>' then subsequent 'Class <CODE>'
    lines. Accepts letters/digits, optional hyphens, optional embedded spaces (we normalize).
    Returns mapping: center_text -> [class_codes_in_order]
    """
    xls = pd.ExcelFile(vf_file)
    # choose the likely sheet (name starts with 'VF_Average_Funded')
    sheet = next((s for s in xls.sheet_names if s.lower().startswith("vf_average_funded")), xls.sheet_names[0])
    raw = pd.read_excel(vf_file, sheet_name=sheet, header=None)

    center_re = re.compile(r'^\s*HCHSP\s*--\s*(.+?)\s*$', re.IGNORECASE)
    class_re  = re.compile(r'^\s*Class\s+([A-Za-z0-9][A-Za-z0-9\s\-\/]*)\s*:?$', re.IGNORECASE)

    mapping: dict[str, list[str]] = {}
    current = None
    for v in raw.iloc[:, 0].astype(str).fillna(""):
        line = v.strip()
        if not line:
            continue
        if CLASS_TOTALS.match(line):
            continue
        m_center = center_re.match(line)
        if m_center:
            current = m_center.group(1).strip()
            mapping.setdefault(current, [])
            continue
        m_class = class_re.match(line)
        if m_class and current:
            cls = m_class.group(1).strip()
            # normalize inside the code: 'E 117 ' -> 'E117', 'B -114' -> 'B-114'
            cls = re.sub(r'\s*-\s*', '-', cls)  # tidy hyphen spacing
            cls = re.sub(r'\s+', '', cls)       # remove spaces inside code
            mapping[current].append(cls)
    return mapping

def find_header_row_by_cols(ws, must_have=("Center", "Class"), scan_rows=100):
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

def norm_center_for_match(x: str) -> str:
    # normalize center names for matching across files
    x = x.lower().strip()
    x = re.sub(r'[\u2013\u2014\-]+', '-', x)            # normalize dashes
    x = re.sub(r'[^a-z0-9\s\-]', '', x)                 # remove stray punctuation
    x = re.sub(r'\s+', ' ', x)
    # common suffixes that often vary
    x = x.replace(' isd', '').replace(' elementary', '').strip()
    return x

def map_centers(main_centers: list[str], vf_centers: list[str]) -> dict[str, str]:
    """
    Build a mapping from main center text -> VF center text using normalized matching.
    Prefers exact normalized match; if none, tries 'contains' either way.
    """
    main_norm = {c: norm_center_for_match(c) for c in main_centers}
    vf_norm   = {c: norm_center_for_match(c) for c in vf_centers}

    # reverse index for VF
    rev = {}
    for orig, n in vf_norm.items():
        rev.setdefault(n, []).append(orig)

    mapping = {}
    for m_orig, m_norm in main_norm.items():
        # exact normalized match
        if m_norm in rev:
            mapping[m_orig] = rev[m_norm][0]
            continue
        # contains in either direction
        picked = None
        for v_orig, v_norm in vf_norm.items():
            if m_norm and (m_norm in v_norm or v_norm in m_norm):
                picked = v_orig
                break
        if picked:
            mapping[m_orig] = picked
    return mapping

def force_classes_from_vf(main_df: pd.DataFrame, vf_map: dict[str, list[str]]) -> pd.DataFrame:
    """
    For each center in the main grid, find the matching center in VF and then
    overwrite the 'Class' values for that center IN ORDER with VF class codes.
    We overwrite:
      • cells that look like placeholders (Class 30, Classroom 09, numeric-only), OR
      • if any lettered class is missing, we simply assign sequentially from top.
    """
    if "Center" not in main_df.columns or "Class" not in main_df.columns:
        st.error("Main sheet must have 'Center' and 'Class' columns.")
        st.stop()

    out = main_df.copy()
    out["Center"] = out["Center"].astype(str).fillna("").str.strip()
    out["Class"]  = out["Class"].astype(str).fillna("").str.strip()

    # build center mapping
    main_centers = [c for c in out["Center"].dropna().unique() if c.strip()]
    vf_centers   = list(vf_map.keys())
    c_map = map_centers(main_centers, vf_centers)

    for m_center in main_centers:
        if m_center not in c_map:
            continue
        vf_center = c_map[m_center]
        classes = [c for c in vf_map.get(vf_center, []) if c and isinstance(c, str)]
        if not classes:
            continue

        # indices of rows for this center
        idxs = out.index[out["Center"].str.strip() == m_center].tolist()
        if not idxs:
            continue

        # determine which rows to overwrite:
        # 1) any placeholder (Class 30 / Classroom 09), or numeric-only ('101')
        rows_to_overwrite = []
        for ridx in idxs:
            val = out.at[ridx, "Class"]
            is_placeholder = bool(PLACEHOLDER_ANY.match(val)) or bool(ONLY_NUMERIC.match(val))
            rows_to_overwrite.append((ridx, is_placeholder))

        # If none marked as placeholder but class set is clearly wrong (e.g., missing letters),
        # we still overwrite sequentially to guarantee correctness.
        if not any(flag for _, flag in rows_to_overwrite):
            rows_to_assign = idxs
        else:
            rows_to_assign = [r for r, flag in rows_to_overwrite if flag]

        # Assign VF classes in order to selected rows
        p = 0
        for ridx in rows_to_assign:
            if p >= len(classes):
                break
            out.at[ridx, "Class"] = classes[p]
            p += 1
        # if there are more rows than VF classes, remaining rows are left as-is (safer)

    return out

def style_headers(ws):
    fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = fill
    for c in range(1, ws.max_column+1):
        ws.column_dimensions[get_column_letter(c)].width = 18 if c > 2 else 28

# ---------------- run ----------------
if main_file and vf_file:
    main_df, sheet_name = read_main_grid(main_file)
    vf_classes_map = parse_vf_classes(vf_file)

    # (Optional) quick peek to verify VF parsing
    st.write("VF classes detected (sample):")
    st.json({k: vf_classes_map[k] for k in sorted(vf_classes_map)[:8]})

    fixed = force_classes_from_vf(main_df, vf_classes_map)

    out_path = "Campus_Classroom_Enrollment_CLASSES_FIXED.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        fixed.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        style_headers(ws)

    with open(out_path, "rb") as f:
        st.download_button("📥 Download fixed workbook", f, file_name=out_path)

    st.success("Class names were overwritten from the VF funded report (letters preserved).")

