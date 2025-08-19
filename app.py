import io
import re
from pathlib import Path
from datetime import date
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="HCHSP Enrollment", layout="wide")

# ----------------------------
# Streamlit header (UI only)
# ----------------------------
logo_path = Path("header_logo.png")
hdr_l, hdr_c, hdr_r = st.columns([1, 2, 1])
with hdr_c:
    if logo_path.exists():
        st.image(str(logo_path), width=320)
    st.markdown(
        "<h1 style='text-align:center; margin: 8px 0 4px;'>Hidalgo County Head Start — Enrollment Formatter</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <p style='text-align:center; font-size:16px; margin-top:0;'>
        Upload the VF Average Funded Enrollment report and the 25–26 Applied/Accepted report.
        Optionally add License Caps via upload or paste.
        </p>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# ----------------------------
# Inputs
# ----------------------------
inp_l, inp_c, inp_r = st.columns([1, 2, 1])
with inp_c:
    vf_file = st.file_uploader("Upload *VF_Average_Funded_Enrollment_Level.xlsx*", type=["xlsx"], key="vf")
    aa_file = st.file_uploader("Upload *25-26 Applied/Accepted.xlsx*", type=["xlsx"], key="aa")
    caps_file = st.file_uploader("Optional: Upload License Caps (CSV/XLSX) with columns Center, Lic. Cap", type=["csv", "xlsx"], key="caps")
    caps_text = st.text_area(
        "Or paste License Caps (one per line: Center,Cap)",
        value="",
        placeholder="Alvarez-McAllen ISD,138\nCamarena-La Joya ISD,192",
        height=100
    )
    process = st.button("Process & Download")

# ----------------------------
# Helpers
# ----------------------------
def norm_center(s: str) -> str:
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = re.sub(r"^HCHSP --\s*", "", s)
    s = re.sub(r"\s+", " ", s)
    return s

def parse_caps_upload(uploaded) -> dict:
    if not uploaded:
        return {}
    name = uploaded.name.lower()
    try:
        if name.endswith(".csv"):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded, sheet_name=0)
    except Exception:
        return {}
    # try to locate columns
    cols = {c.lower().strip(): c for c in df.columns}
    center_col = None
    cap_col = None
    for key, orig in cols.items():
        if key in ("center", "location (head start)", "site", "location"):
            center_col = orig
        if key in ("lic. cap", "lic cap", "license cap", "lic_cap", "liccap"):
            cap_col = orig
    if center_col is None or cap_col is None:
        # best-effort fallback: first two cols
        if len(df.columns) >= 2:
            center_col, cap_col = df.columns[:2]
        else:
            return {}
    df = df[[center_col, cap_col]].copy()
    df[center_col] = df[center_col].map(norm_center)
    df[cap_col] = pd.to_numeric(df[cap_col], errors="coerce")
    df = df.dropna(subset=[center_col, cap_col])
    return {r[center_col]: int(r[cap_col]) for _, r in df.iterrows()}

def parse_caps_text(text: str) -> dict:
    caps = {}
    if not text.strip():
        return caps
    for line in text.strip().splitlines():
        parts = re.split(r"[,\t;]\s*", line.strip(), maxsplit=1)
        if len(parts) != 2:
            continue
        name, cap = parts[0], parts[1]
        name = norm_center(name)
        try:
            cap_val = int(float(cap))
        except Exception:
            continue
        if name:
            caps[name] = cap_val
    return caps

# ----------------------------
# Parsers
# ----------------------------
def parse_vf(vf_df_raw: pd.DataFrame) -> pd.DataFrame:
    """Parse VF report (header=None) into per-class rows: Center | Class | Funded | Enrolled"""
    records = []
    current_center = None
    current_class = None

    for i in range(len(vf_df_raw)):
        c0 = vf_df_raw.iloc[i, 0]
        if isinstance(c0, str) and c0.startswith("HCHSP --"):
            current_center = c0.strip()
        elif isinstance(c0, str) and re.match(r"^Class \d+", c0):
            current_class = c0.split(" ", 1)[1].strip()

        if c0 == "Class Totals:" and current_center and current_class:
            row = vf_df_raw.iloc[i]
            funded = pd.to_numeric(row.iloc[4], errors="coerce")
            enrolled = pd.to_numeric(row.iloc[3], errors="coerce")
            center_clean = norm_center(current_center)
            records.append({
                "Center": center_clean,
                "Class": f"Class {current_class}",
                "Funded": 0 if pd.isna(funded) else float(funded),
                "Enrolled": 0 if pd.isna(enrolled) else float(enrolled),
            })

    tidy = pd.DataFrame(records)
    if tidy.empty:
        raise ValueError("Could not find any 'Class Totals:' rows in the VF report. Check that you're uploading the correct file.")
    return tidy


def parse_applied_accepted(aa_df_raw: pd.DataFrame) -> pd.DataFrame:
    """Parse Applied/Accepted (header=None) to per-center counts; only blank 'ST: Status End Date' rows kept."""
    header_row_idx = aa_df_raw.index[aa_df_raw.iloc[:, 0].astype(str).str.startswith("ST: Participant PID", na=False)]
    if len(header_row_idx) == 0:
        raise ValueError("Could not find header row in Applied/Accepted report (expected a row starting with 'ST: Participant PID').")
    header_row_idx = int(header_row_idx[0])
    headers = aa_df_raw.iloc[header_row_idx].tolist()
    body = pd.DataFrame(aa_df_raw.iloc[header_row_idx + 1:].values, columns=headers)

    center_col = "ST: Center Name"
    status_col = "ST: Status"
    date_col = "ST: Status End Date"

    is_blank_date = body[date_col].isna() | body[date_col].astype(str).str.strip().eq("")
    body = body[is_blank_date].copy()
    body[center_col] = body[center_col].map(norm_center)

    counts = body.groupby(center_col)[status_col].value_counts().unstack(fill_value=0)
    for c in ["Accepted", "Applied"]:
        if c not in counts.columns: counts[c] = 0

    return counts[["Accepted", "Applied"]].astype(int).reset_index().rename(columns={center_col: "Center"})

# ----------------------------
# Builder
# ----------------------------
def build_output_table(vf_tidy: pd.DataFrame, counts: pd.DataFrame, lic_caps: dict) -> pd.DataFrame:
    merged = vf_tidy.merge(counts, on="Center", how="left").fillna({"Accepted": 0, "Applied": 0})
    merged["% Enrolled of Funded"] = np.where(
        merged["Funded"] > 0,
        (merged["Enrolled"] / merged["Funded"] * 100).round(0).astype("Int64"),
        pd.NA
    )

    applied_by_center = merged.groupby("Center")["Applied"].max()
    accepted_by_center = merged.groupby("Center")["Accepted"].max()

    rows = []
    waitlist_totals = 0
    agency_classrooms_total = 0
    agency_lic_cap_total = 0

    for center, group in merged.groupby("Center", sort=True):
        # Class rows
        for _, r in group.iterrows():
            rows.append({
                "Center": r["Center"],
                "Class": r["Class"],
                "# Classrooms": "",
                "Lic. Cap": "",
                "Funded": int(r["Funded"]),
                "Enrolled": int(r["Enrolled"]),
                "Applied": "",
                "Accepted": "",
                "Lacking/Overage": "",
                "Waitlist": "",
                "% Enrolled of Funded": int(r["% Enrolled of Funded"]) if pd.notna(r["% Enrolled of Funded"]) else pd.NA
            })

        # Center totals
        funded_sum = int(group["Funded"].sum())
        enrolled_sum = int(group["Enrolled"].sum())
        pct_total = int(round(enrolled_sum / funded_sum * 100, 0)) if funded_sum > 0 else pd.NA

        accepted_val = int(accepted_by_center.get(center, 0))
        applied_val  = int(applied_by_center.get(center, 0))
        waitlist_val = accepted_val if enrolled_sum > funded_sum else ""
        lacking_over = funded_sum - enrolled_sum

        class_count = int(len(group))  # number of class rows for this center
        agency_classrooms_total += class_count

        lic_cap_val = lic_caps.get(center)
        if isinstance(lic_cap_val, (int, float)) and not pd.isna(lic_cap_val):
            agency_lic_cap_total += int(lic_cap_val)

        if waitlist_val != "":
            waitlist_totals += waitlist_val

        rows.append({
            "Center": f"{center} Total",
            "Class": "",
            "# Classrooms": class_count,
            "Lic. Cap": ("" if lic_cap_val is None else int(lic_cap_val)),
            "Funded": funded_sum,
            "Enrolled": enrolled_sum,
            "Applied": applied_val,
            "Accepted": accepted_val,
            "Lacking/Overage": lacking_over,
            "Waitlist": waitlist_val,
            "% Enrolled of Funded": pct_total
        })

    final = pd.DataFrame(rows)

    # Agency totals (Lic. Cap is the sum of provided caps)
    agency_funded   = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Funded"].sum())
    agency_enrolled = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Enrolled"].sum())
    agency_applied  = int(counts["Applied"].sum())
    agency_accepted = int(counts["Accepted"].sum())
    agency_pct      = int(round(agency_enrolled / agency_funded * 100, 0)) if agency_funded > 0 else pd.NA
    agency_lacking  = agency_funded - agency_enrolled

    final = pd.concat([final, pd.DataFrame([{
        "Center": "Agency Total",
        "Class": "",
        "# Classrooms": agency_classrooms_total,
        "Lic. Cap": (agency_lic_cap_total if agency_lic_cap_total > 0 else ""),
        "Funded": agency_funded,
        "Enrolled": agency_enrolled,
        "Applied": agency_applied,
        "Accepted": agency_accepted,
        "Lacking/Overage": agency_lacking,
        "Waitlist": waitlist_totals,
        "% Enrolled of Funded": agency_pct
    }])], ignore_index=True)

    # Final column order
    final = final[[
        "Center","Class","# Classrooms","Lic. Cap",
        "Funded","Enrolled","Applied","Accepted","Lacking/Overage","Waitlist","% Enrolled of Funded"
    ]]
    return final

# ----------------------------
# Excel Writer (logo to B, titles in C..last, thick outer box to row 1)
# ----------------------------
def to_styled_excel(df: pd.DataFrame) -> bytes:
    """Logo at B1 (53% scale), titles centered in C..last with no inner lines; thick outer box from row 1 (right edge fixed); borders on table; gridlines outside kept."""
    def col_letter(n: int) -> str:
        s = ""
        while n >= 0:
            s = chr(n % 26 + 65) + s
            n = n // 26 - 1
        return s

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Head Start Enrollment", startrow=3)
        wb = writer.book
        ws = writer.sheets["Head Start Enrollment"]

        # keep Excel gridlines outside
        ws.hide_gridlines(0)

        # title area heights
        ws.set_row(0, 24)
        ws.set_row(1, 22)
        ws.set_row(2, 20)

        # --- Logo at B1 (53% scale) ---
        logo = Path("header_logo.png")
        if logo.exists():
            ws.set_column(1, 1, 6)  # column B width for logo
            ws.insert_image(0, 1, str(logo), {
                "x_offset": 2, "y_offset": 2,
                "x_scale": 0.53, "y_scale": 0.53,
                "object_position": 1
            })

        # --- Titles in C..last (no inner borders) ---
        d = date.today()
        date_str = f"{d.month}.{d.day}.{str(d.year % 100).zfill(2)}"
        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        red_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": "#C00000"})

        last_col_0 = len(df.columns) - 1
        last_col_letter = col_letter(last_col_0)

        # merge titles starting at column C (index 2)
        ws.merge_range(0, 2, 0, last_col_0, "Hidalgo County Head Start Program", title_fmt)
        ws.merge_range(1, 2, 1, last_col_0, "", subtitle_fmt)

