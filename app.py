# Overwrite app.py with a centered header (no logo uploader), fix missing imports, and wire header_logo.png into Excel export.
from pathlib import Path
from textwrap import dedent

new_app = dedent(r'''
import io
import re
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image

st.set_page_config(page_title="HCHSP Enrollment Report (2025–2026)", layout="wide")

# ===== Header (centered logo + title, no logo uploader) =====
logo_path = Path("header_logo.png")
logo_bytes = None
if logo_path.exists():
    try:
        logo_bytes = logo_path.read_bytes()
    except Exception:
        logo_bytes = None

col_l, col_c, col_r = st.columns([1, 2, 1])
with col_c:
    if logo_path.exists():
        st.image(str(logo_path), width=300)
    st.markdown(
        "<h1 style='text-align:center; margin-top:0; margin-bottom:6px;'>HCHSP Enrollment Report (2025–2026)</h1>",
        unsafe_allow_html=True
    )
    st.markdown(
        "<p style='text-align:center; font-size:16px; margin-top:0;'>Upload the VF Average Funded Enrollment report and the 25–26 Applied/Accepted report, then click Process & Download.</p>",
        unsafe_allow_html=True
    )

st.divider()

# ---- Inputs ----
vf_file = st.file_uploader("Upload *VF_Average_Funded_Enrollment_Level.xlsx*", type=["xlsx"], key="vf")
aa_file = st.file_uploader("Upload *25-26 Applied/Accepted.xlsx*", type=["xlsx"], key="aa")

# ----------------------------
# Utilities
# ----------------------------

def parse_vf(vf_df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Parse VF report (header=None) into per-class rows:
    Center | Class | Funded | Enrolled
    """
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
            funded = pd.to_numeric(row.iloc[4], errors="coerce")  # Number of Federal Slots Available
            enrolled = pd.to_numeric(row.iloc[3], errors="coerce")  # Number of Children Enrolled
            center_clean = re.sub(r"^HCHSP --\s*", "", current_center)
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
    """
    Parse Applied/Accepted report (header=None) to per-center counts.
    Excludes rows with a Status End Date.
    """
    header_row_idx = aa_df_raw.index[aa_df_raw.iloc[:, 0].astype(str).str.startswith("ST: Participant PID", na=False)]
    if len(header_row_idx) == 0:
        raise ValueError("Could not find header row in Applied/Accepted report (expected a row starting with 'ST: Participant PID').")
    header_row_idx = int(header_row_idx[0])
    headers = aa_df_raw.iloc[header_row_idx].tolist()
    body = pd.DataFrame(aa_df_raw.iloc[header_row_idx + 1:].values, columns=headers)

    center_col = "ST: Center Name"
    status_col = "ST: Status"
    date_col = "ST: Status End Date"

    body = body[body[date_col].isna()].copy()
    body[center_col] = body[center_col].astype(str).str.replace(r"^HCHSP --\s*", "", regex=True)

    counts = body.groupby(center_col)[status_col].value_counts().unstack(fill_value=0)
    for c in ["Accepted", "Applied"]:
        if c not in counts.columns:
            counts[c] = 0
    counts = counts[["Accepted", "Applied"]].astype(int).reset_index().rename(columns={center_col:"Center"})
    return counts


def build_output_table(vf_tidy: pd.DataFrame, counts: pd.DataFrame) -> pd.DataFrame:
    """
    Merge per-class VF with center-level counts; add center totals and agency total.
    Class rows: blank Applied/Accepted. Totals: include Applied/Accepted.
    % column is numeric whole number (e.g., 87).
    """
    merged = vf_tidy.merge(counts, on="Center", how="left").fillna({"Accepted": 0, "Applied": 0})
    merged["% Enrolled of Funded"] = np.where(
        merged["Funded"] > 0,
        (merged["Enrolled"] / merged["Funded"] * 100).round(0).astype("Int64"),
        pd.NA
    )

    applied_by_center = merged.groupby("Center")["Applied"].max()
    accepted_by_center = merged.groupby("Center")["Accepted"].max()

    rows = []
    for center, group in merged.groupby("Center", sort=True):
        for _, r in group.iterrows():
            rows.append({
                "Center": r["Center"],
                "Class": r["Class"],
                "Funded": int(r["Funded"]),
                "Enrolled": int(r["Enrolled"]),
                "Applied": "",    # blank on class rows
                "Accepted": "",   # blank on class rows
                "% Enrolled of Funded": int(r["% Enrolled of Funded"]) if pd.notna(r["% Enrolled of Funded"]) else pd.NA
            })
        funded_sum = int(group["Funded"].sum())
        enrolled_sum = int(group["Enrolled"].sum())
        pct_total = int(round(enrolled_sum / funded_sum * 100, 0)) if funded_sum > 0 else pd.NA

        rows.append({
            "Center": f"{center} Total",
            "Class": "",
            "Funded": funded_sum,
            "Enrolled": enrolled_sum,
            "Applied": int(applied_by_center.get(center, 0)),
            "Accepted": int(accepted_by_center.get(center, 0)),
            "% Enrolled of Funded": pct_total
        })

    final = pd.DataFrame(rows)

    # Agency totals computed from center totals
    agency_funded = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Funded"].sum())
    agency_enrolled = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Enrolled"].sum())
    agency_applied  = int(applied_by_center.sum())
    agency_accepted = int(accepted_by_center.sum())
    agency_pct = int(round(agency_enrolled / agency_funded * 100, 0)) if agency_funded > 0 else pd.NA

    final = pd.concat([
        final,
        pd.DataFrame([{
            "Center": "Agency Total",
            "Class": "",
            "Funded": agency_funded,
            "Enrolled": agency_enrolled,
            "Applied": agency_applied,
            "Accepted": agency_accepted,
            "% Enrolled of Funded": agency_pct
        }])
    ], ignore_index=True)

    final = final[["Center","Class","Funded","Enrolled","Applied","Accepted","% Enrolled of Funded"]]
    return final


def to_styled_excel(df: pd.DataFrame, logo_bytes: bytes | None) -> bytes:
    """
    Create the Excel with:
    - Title + subtitle
    - Optional logo (top-right)
    - Bold headers + filters + freeze panes
    - % column numeric with % display
    - Red font for % < 100
    - Bold rows for center totals and agency total
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Formatted", startrow=3)
        wb = writer.book
        ws = writer.sheets["Formatted"]

        # Titles
        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        ws.merge_range(0, 0, 0, len(df.columns)-1, "Hidalgo County Head Start Program", title_fmt)
        ws.merge_range(1, 0, 1, len(df.columns)-1, "2025-2026 Campus Classroom Enrollment", subtitle_fmt)

        # Fixed logo (top-right)
        if logo_bytes is not None:
            end_col = len(df.columns) - 1
            image_opts = {"x_scale": 0.6, "y_scale": 0.6, "x_offset": 10, "y_offset": 2}
            try:
                ws.insert_image(0, end_col, "header_logo.png", {"image_data": io.BytesIO(logo_bytes), **image_opts})
            except Exception:
                pass

        # Bold headers
        header_fmt = wb.add_format({"bold": True})
        for c, col in enumerate(df.columns):
            ws.write(3, c, col, header_fmt)

        # Filters + freeze header
        last_row = len(df) + 3
        last_col = len(df.columns) - 1
        ws.autofilter(3, 0, last_row, last_col)
        ws.freeze_panes(4, 0)

        # % column numeric with percent sign
        percent_col_idx = df.columns.get_loc("% Enrolled of Funded")
        percent_fmt = wb.add_format({"num_format": '0"%"'})
        ws.set_column(percent_col_idx, percent_col_idx, 16, percent_fmt)

        # Conditional red for % < 100
        def colnum_string(n: int) -> str:
            s = ""
            while n >= 0:
                s = chr(n % 26 + 65) + s
                n = n // 26 - 1
            return s
        percent_letter = colnum_string(percent_col_idx)
        percent_range = f"{percent_letter}5:{percent_letter}{last_row+1}"
        ws.conditional_format(percent_range, {
            "type": "cell",
            "criteria": "<",
            "value": 100,
            "format": wb.add_format({"font_color": "red"})
        })

        # Bold totals (center rows ending with " Total" or Agency Total)
        bold_fmt = wb.add_format({"bold": True})
        for ridx, val in enumerate(df["Center"].tolist()):
            if (isinstance(val, str) and val.endswith(" Total")) or (val == "Agency Total"):
                ws.set_row(ridx + 4, None, bold_fmt)

    return output.getvalue()


# ----------------------------
# Main
# ----------------------------
if st.button("Process & Download") and vf_file and aa_file:
    try:
        vf_raw = pd.read_excel(vf_file, sheet_name=0, header=None)
        aa_raw = pd.read_excel(aa_file, sheet_name=0, header=None)

        vf_tidy = parse_vf(vf_raw)
        aa_counts = parse_applied_accepted(aa_raw)
        final_df = build_output_table(vf_tidy, aa_counts)

        st.success("Preview below. Use the download button to get the Excel file.")
        st.dataframe(final_df, use_container_width=True)

        xlsx_bytes = to_styled_excel(final_df, logo_bytes)
        st.download_button(
            "Process & Download",
            data=xlsx_bytes,
            file_name="HCHSP_Enrollment_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Processing error: {e}")
''')

Path("/mnt/data/app.py").write_text(new_app, encoding="utf-8")
"/mnt/data/app.py"

