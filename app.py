# app.py — Get classes from VF funded report, assign to students (two-file solution)
from datetime import datetime
from zoneinfo import ZoneInfo
import re
from typing import Dict, List

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Classes from VF (Two-File Upload)", layout="centered")

st.title("Assign Class Names from VF Funded Report")
st.markdown(
    "Upload **(1)** your student QuickReport (no class column is fine) and **(2)** your **VF funded report**.  \n"
    "The app will **pull class names from the VF report** and **assign them deterministically per center** "
    "(round-robin by PID). Letters in class names (e.g., `E117`, `A01`) are preserved from VF."
)

students_file = st.file_uploader("1) Student QuickReport (.xlsx)", type=["xlsx"], key="students")
vf_file       = st.file_uploader("2) VF funded report (.xlsx)", type=["xlsx"], key="vf")

# ---------------- Helpers ----------------

CLASS_TOTALS = re.compile(r"^\s*class\s+totals\s*:?\s*$", re.IGNORECASE)

def find_header_row(ws, probes=("ST: Participant PID", "ST: Center Name"), scan_rows=120):
    """Find the first row where ALL probe headers appear (case-insensitive) within that row."""
    probes = [p.lower() for p in probes]
    for r in ws.iter_rows(min_row=1, max_row=scan_rows):
        vals = [str(c.value).strip().lower() if c.value is not None else "" for c in r]
        if all(any(p == v for v in vals) for p in probes):
            return r[0].row
    return None

def load_students_frame(xlsx) -> pd.DataFrame:
    """Load student QuickReport: detect sheet + header row; must include ST: Participant PID and ST: Center Name."""
    wb = load_workbook(xlsx, data_only=True)
    chosen, hdr = None, None
    for s in wb.sheetnames:
        ws = wb[s]
        hdr = find_header_row(ws)
        if hdr:
            chosen = s
            break
    if not chosen:
        st.error("Couldn’t find a sheet in the student QuickReport containing both "
                 "'ST: Participant PID' and 'ST: Center Name'.")
        st.stop()
    xlsx.seek(0)
    df = pd.read_excel(xlsx, sheet_name=chosen, header=hdr-1, dtype=str)
    # keep only relevant columns if present
    keep = [c for c in df.columns if str(c).strip() in [
        "ST: Participant PID", "ST: First Name", "ST: Last Name",
        "ST: Center Name", "ST: Status", "ST: Status End Date"
    ]]
    if keep:
        df = df[keep].copy()
    return df

def parse_vf_funded(xlsx) -> Dict[str, List[str]]:
    """
    Parse VF funded report into {center -> [class codes]}.
    Looks for rows like:
       HCHSP -- <Center Name>
       Class E117
       Class A01
       ...
    Normalizes inside the codes (e.g., 'E 117' -> 'E117', 'B -114' -> 'B-114').
    """
    xls = pd.ExcelFile(xlsx)
    # pick likely sheet (prefix match works with your files)
    sheet = next((s for s in xls.sheet_names if s.lower().startswith("vf_average_funded")), xls.sheet_names[0])
    raw = pd.read_excel(xlsx, sheet_name=sheet, header=None)

    center_re = re.compile(r'^\s*HCHSP\s*--\s*(.+?)\s*$', re.IGNORECASE)
    class_re  = re.compile(r'^\s*Class\s+([A-Za-z0-9][A-Za-z0-9\s\-/]*)\s*:?$', re.IGNORECASE)

    mapping: Dict[str, List[str]] = {}
    current = None
    col0 = raw.iloc[:, 0].astype(str).fillna("")
    for val in col0:
        line = val.strip()
        if not line or CLASS_TOTALS.match(line):
            continue
        m_center = center_re.match(line)
        if m_center:
            current = m_center.group(1).strip()
            mapping.setdefault(current, [])
            continue
        m_class = class_re.match(line)
        if m_class and current:
            code = m_class.group(1).strip()
            code = re.sub(r'\s*-\s*', '-', code)  # tidy hyphen spacing
            code = re.sub(r'\s+', '', code)       # remove inner spaces
            mapping[current].append(code)
    return mapping

def norm_center(x: str) -> str:
    """Normalize center strings so the two files match more easily."""
    x = (x or "").lower().strip()
    x = re.sub(r'[\u2013\u2014\-]+', '-', x)   # normalize dashes
    x = re.sub(r'[^a-z0-9\s\-]', '', x)       # drop stray punctuation
    x = re.sub(r'\s+', ' ', x)
    # common suffixes
    for tail in [" isd", " elementary", " elem", " school"]:
        if x.endswith(tail):
            x = x[: -len(tail)]
    return x.strip()

def build_center_map(student_centers: List[str], vf_centers: List[str]) -> Dict[str, str]:
    """Map student center strings -> VF funded center strings (normalized exact or contains)."""
    s_norm = {s: norm_center(s) for s in student_centers}
    v_norm = {v: norm_center(v) for v in vf_centers}
    rev: Dict[str, List[str]] = {}
    for orig, n in v_norm.items():
        rev.setdefault(n, []).append(orig)

    mapping = {}
    for s_orig, s_n in s_norm.items():
        if not s_n:
            continue
        if s_n in rev:
            mapping[s_orig] = rev[s_n][0]
            continue
        # contains either way
        pick = None
        for v_orig, v_n in v_norm.items():
            if s_n in v_n or v_n in s_n:
                pick = v_orig
                break
        if pick:
            mapping[s_orig] = pick
    return mapping

def deterministic_assign(df: pd.DataFrame, classes: List[str]) -> pd.Series:
    """
    Assign classes round-robin in a deterministic order by PID (string).
    This guarantees stable results across runs with the same inputs.
    """
    # sort by PID as string
    tmp = df.copy()
    tmp["__pid"] = tmp["ST: Participant PID"].astype(str).fillna("")
    tmp = tmp.sort_values("__pid", kind="mergesort")  # stable
    cls = []
    k = len(classes)
    if k == 0:
        return pd.Series([""] * len(df), index=df.index)

    # round-robin
    for i, _ in enumerate(tmp.index):
        cls.append(classes[i % k])
    assigned = pd.Series(cls, index=tmp.index)
    # reindex back to original order
    return assigned.reindex(df.index)

def style_headers(ws):
    fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = fill
    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = 22 if c <= 3 else 18

# ---------------- Run ----------------

if students_file and vf_file:
    # 1) Load both files
    students = load_students_frame(students_file)
    vf_classes_map = parse_vf_funded(vf_file)

    # 2) Prepare center names
    students["Center (clean)"] = students["ST: Center Name"].astype(str).str.replace(
        r"^HCHSP\s*--\s*", "", regex=True
    ).str.strip()
    s_centers = [c for c in students["Center (clean)"].dropna().unique() if c.strip()]
    v_centers = list(vf_classes_map.keys())
    c_map = build_center_map(s_centers, v_centers)

    # Diagnostics for centers that don’t match
    missing_in_vf = sorted(set(s_centers) - set(c_map.keys()))
    extra_in_vf = sorted(set(v_centers) - set(c_map.values()))

    # 3) Assign classes from VF per center (deterministic, round-robin by PID)
    students["Class Name"] = ""
    for s_center in s_centers:
        mask = students["Center (clean)"] == s_center
        if s_center not in c_map:
            continue
        vf_center = c_map[s_center]
        class_list = [c for c in vf_classes_map.get(vf_center, []) if c and isinstance(c, str)]
        if not class_list:
            continue
        block = students.loc[mask]
        students.loc[mask, "Class Name"] = deterministic_assign(block, class_list)

    # 4) Output
    out_xlsx = "Students_With_Classes_From_VF.xlsx"
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        # Students with classes (primary deliverable)
        students.to_excel(writer, index=False, sheet_name="Students + Classes")
        ws = writer.book["Students + Classes"]
        style_headers(ws)
        # Optional diagnostics
        diag = pd.DataFrame({
            "Issue": (["Center not found in VF"] * len(missing_in_vf)) + (["VF center unused"] * len(extra_in_vf)),
            "Center": missing_in_vf + extra_in_vf
        })
        if not diag.empty:
            diag.to_excel(writer, index=False, sheet_name="Diagnostics")
    with open(out_xlsx, "rb") as f:
        st.download_button("📥 Download Students_With_Classes_From_VF.xlsx", f, file_name=out_xlsx)

    st.success("Done: ‘Class Name’ was filled using the class list from the VF funded report (letters preserved).")
    st.caption("If you later provide a QuickReport with ‘ST: Class Name’, you can replace the assignment with an authoritative PID merge.")
