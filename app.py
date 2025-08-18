# app.py
import io
import re
import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image

st.set_page_config(page_title="HCHSP Enrollment Report Formatter (2025–2026)", layout="wide")

# ---- App Header (logo only in the UI; not embedded in Excel) ----
try:
    logo = Image.open("header_logo.png")  # keep this file next to app.py
    st.image(logo, width=300)
except Exception:
    pass

st.title("HCHSP Enrollment Report Formatter (2025–2026)")
st.divider()

# ---- File Uploads (stacked) ----
vf_file = st.file_uploader("Upload *VF_Average_Funded_Enrollment_Level.xlsx*", type=["xlsx"], key="vf")
aa_file = st.file_uploader("Upload *25-26 Applied/Accepted.xlsx*", type=["xlsx"], key="aa")

# ----------------------------
# Utilities
# ----------------------------
def parse_vf(vf_df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Parse the VF export (header=None) by scanning Column A markers:
      - Center lines start with "HCHSP --"
      - Class lines like "Class 09", "Class 10", ...
      - A 'Class Total...' line contains the numeric totals for that class
        (by default: Enrolled at col index 3, Funded at col index 4)
    Returns rows with exact class label from the VF (e.g., 'Class 09').
    """
    records = []
    current_center = None
    current_class_label = None

    for i in range(len(vf_df_raw)):
        c0 = vf_df_raw.iloc[i, 0]

        # Detect center marker
        if isinstance(c0, str) and c0.strip().startswith("HCHSP --"):
            current_center = c0.strip()

        # Detect class marker ("Class NN")
        elif isinstance(c0, str) and re.match(r"^\s*Class\s+\d+\b", c0, flags=re.I):
            # Keep EXACT label (e.g., "Class 09") so it transfers as-is
            current_class_label = c0.strip()

        # When we hit the totals row for that class, read Enrolled/Funded
        if (
            isinstance(c0, str)
            and c0.strip().lower().startswith("class total")  # handles "Class Total", "Class Totals:", etc.
            and current_center
            and current_class_label
        ):
            row = vf_df_raw.iloc[i]
            # Default indices: Enrolled @ 3, Funded @ 4 (adjust here if your export changes)
            enrolled = pd.to_numeric(row.iloc[3], errors="coerce")
            funded   = pd.to_numeric(row.iloc[4], errors="coerce")

            center_clean = re.sub(r"^HCHSP --\s*", "", current_center).strip()
            records.append({
                "Center": center_clean,
                "Class": current_class_label,                        # <-- exact label from VF ("Class 09")
                "Funded": 0 if pd.isna(funded) else int(funded),
                "Enrolled": 0 if pd.isna(enrolled) else int(enrolled),
            })

    tidy = pd.DataFrame(records)
    if tidy.empty:
        raise ValueError("Could not find any class totals rows in the VF file. Check that Column A contains Class/Center markers and a 'Class Total' line per class.")
    return tidy


def parse_applied_accepted(aa_df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Parse Applied/Accepted export (header=None) to per-center counts.
    Rows with a Status End Date are excluded.
    """
    idx = aa_df_raw.index[aa_df_raw.iloc[:, 0].astype(str).str.startswith("ST: Participant PID", na=False)]
    if len(idx) == 0:
        raise ValueError("Could not find header row in Applied/Accepted report (expected a row starting with 'ST: Participant PID').")
    header_row_idx = int(idx[0])

    headers = aa_df_raw.iloc[header_row_idx].tolist()
    body = pd.DataFrame(aa_df_raw.iloc[header_row_idx + 1:].values, columns=headers)

    center_col = "ST: Center Name"
    status_col = "ST: Status"
    date_col   = "ST: Status End Date"

    body = body[body[date_col].isna()].copy()
    body[center_col] = body[center_col].astype(str).str.replace(r"^HCHSP --\s*", "", regex=True)

    counts = body.groupby(center_col)[status_col].value_counts().unstack(fill_value=0)
    for c in ["Accepted", "Applied"]:
        if c not in counts.columns:
            counts[c] = 0
    counts = counts[["Accepted", "Applied"]].astype(int).reset_index().rename(columns={center_col: "Center"})
    return counts


def calc_waitlist_lacking(funded: int, enrolled: int) -> tuple[int, int]:
    """Waitlist = max(0, Enrolled - Funded); Lacking = max(0, Funded - Enrolled)."""
    if enrolled > funded:
        return enrolled - funded, 0
    elif funded > enrolled:
        return 0, funded - enrolled
    else:
        return 0, 0


def build_output_table(vf_tidy: pd.DataFrame, counts: pd.DataFrame) -> pd.DataFrame:
    """
    Final table rules:
      - Class rows: show Class (exact), Funded, Enrolled, Waitlist, Lacking, %, and BLANK Applied/Accepted.
      - Center Total rows: sums + Applied/Accepted (from A/A).
      - Agency Total row: overall sums.
    """
    merged = vf_tidy.merge(counts, on="Center", how="left").fillna({"Accepted": 0, "Applied": 0})

    rows = []
    for center, group in merged.groupby("Center", sort=True):
        # Class rows
        for _, r in group.iterrows():
            waitlist, lacking = calc_waitlist_lacking(int(r["Funded"]), int(r["Enrolled"]))
            pct = int(round(r["Enrolled"] / r["Funded"] * 100, 0)) if r["Funded"] > 0 else pd.NA
            rows.append({
                "Center": r["Center"],
                "Class": r["Class"],   # exact "Class NN" from VF
                "Funded": int(r["Funded"]),
                "Enrolled": int(r["Enrolled"]),
                "Applied": "",         # blank on class rows
                "Accepted": "",        # blank on class rows
                "Waitlist": waitlist,
                "Lacking": lacking,
                "% Enrolled of Funded": pct
            })

        # Center Total row
        funded_sum   = int(group["Funded"].sum())
        enrolled_sum = int(group["Enrolled"].sum())
        wait_sum, lack_sum = calc_waitlist_lacking(funded_sum, enrolled_sum)
        pct_total = int(round(enrolled_sum / funded_sum * 100, 0)) if funded_sum > 0 else pd.NA

        rows.append({
            "Center": f"{center} Total",
            "Class": "",  # totals row keeps Class blank
            "Funded": funded_sum,
            "Enrolled": enrolled_sum,
            "Applied": int(group["Applied"].max()),
            "Accepted": int(group["Accepted"].max()),
            "Waitlist": wait_sum,
            "Lacking": lack_sum,
            "% Enrolled of Funded": pct_total
        })

    final = pd.DataFrame(rows)

    # Agency Total row
    center_total_mask = final["Center"].astype(str).str.endswith(" Total", na=False)
    agency_funded   = int(final.loc[center_total_mask, "Funded"].sum())
    agency_enrolled = int(final.loc[center_total_mask, "Enrolled"].sum())
    wait_agency, lack_agency = calc_waitlist_lacking(agency_funded, agency_enrolled)
    agency_applied  = int(merged["Applied"].sum())
    agency_accepted = int(merged["Accepted"].sum())
    agency_pct      = int(round(agency_enrolled / agency_funded * 100, 0)) if agency_funded > 0 else pd.NA

    final = pd.concat([
        final,
        pd.DataFrame([{
            "Center": "Agency Total",
            "Class": "",
            "Funded": agency_funded,
            "Enrolled": agency_enrolled,
            "Applied": agency_applied,
            "Accepted": agency_accepted,
            "Waitlist": wait_agency,
            "Lacking": lack_agency,
            "% Enrolled of Funded": agency_pct
        }])
    ], ignore_index=True)

    final = final[["Center","Class","Funded","Enrolled","Applied","Accepted","Waitlist","Lacking","% Enrolled of Funded"]]
    return final


def to_styled_excel(df: pd.DataFrame) -> bytes:
    """
    Excel:
      - Black title/subtitle
      - Blue header band
      - Filters, frozen header
      - % column formatted with a % sign for ALL rows (including totals)
      - % < 100 red, % > 100 blue
      - Lacking > 0 red
      - Bold totals rows
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Formatted", startrow=3)
        wb = writer.book
        ws = writer.sheets["Formatted"]

        # Titles (black)
        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        ws.merge_range(0, 0, 0, len(df.columns)-1, "Hidalgo County Head Start Program", title_fmt)
        ws.merge_range(1, 0, 1, len(df.columns)-1, "2025–2026 Campus Classroom Enrollment", subtitle_fmt)

        # Blue header band
        header_fmt = wb.add_format({"bold": True, "bg_color": "#B7DEE8"})
        for c, col in enumerate(df.columns):
            ws.write(3, c, col, header_fmt)

        # Filters + freeze header
        last_row = len(df) + 3
        last_col = len(df.columns) - 1
        ws.autofilter(3, 0, last_row, last_col)
        ws.freeze_panes(4, 0)

        # % column formatting (whole number with % sign)
        percent_col_idx = df.columns.get_loc("% Enrolled of Funded")
        percent_fmt = wb.add_format({"num_format": '0"%"'})
        ws.set_column(percent_col_idx, percent_col_idx, 18, percent_fmt)

        # Helper: convert 0-based col idx to Excel letter(s)
        def colnum_string(n: int) -> str:
            s = ""
            while n >= 0:
                s = chr(n % 26 + 65) + s
                n = n // 26 - 1
            return s

        # Conditional formatting for % column
        percent_letter = colnum_string(percent_col_idx)
        percent_range = f"{percent_letter}5:{percent_letter}{last_row+1}"

        # % < 100 red
        ws.conditional_format(percent_range, {
            "type": "cell", "criteria": "<", "value": 100,
            "format": wb.add_format({"font_color": "red"})
        })
        # % > 100 blue
        ws.conditional_format(percent_range, {
            "type": "cell", "criteria": ">", "value": 100,
            "format": wb.add_format({"font_color": "blue"})
        })
        # (Exactly 100% stays default black)

        # Lacking > 0 red
        lacking_idx = df.columns.get_loc("Lacking")
        lacking_letter = colnum_string(lacking_idx)
        lacking_range = f"{lacking_letter}5:{lacking_letter}{last_row+1}"
        ws.conditional_format(lacking_range, {
            "type": "cell", "criteria": ">", "value": 0,
            "format": wb.add_format({"font_color": "red"})
        })

        # Bold totals (center totals + agency total)
        bold_fmt = wb.add_format({"bold": True})
        for ridx, val in enumerate(df["Center"].tolist()):
            if (isinstance(val, str) and val.endswith(" Total")) or (val == "Agency Total"):
                ws.set_row(ridx + 4, None, bold_fmt)

    return output.getvalue()

# ----------------------------
# Main
# ----------------------------
if st.button("Process & Download"):
    if not vf_file or not aa_file:
        st.warning("Please upload BOTH files to proceed.")
    else:
        try:
            # Read both files as raw (header=None) for the VF parser, headerless-safe for A/A too
            vf_raw = pd.read_excel(vf_file, sheet_name=0, header=None)
            aa_raw = pd.read_excel(aa_file, sheet_name=0, header=None)

            vf_tidy = parse_vf(vf_raw)                 # Center, Class (exact), Funded, Enrolled
            aa_counts = parse_applied_accepted(aa_raw)  # Center-level Applied/Accepted

            final_df = build_output_table(vf_tidy, aa_counts)

            st.success("Preview below. Use the download button to get the Excel file.")
            st.dataframe(final_df, use_container_width=True)

            xlsx_bytes = to_styled_excel(final_df)
            st.download_button(
                "Download Formatted Excel",
                data=xlsx_bytes,
                file_name="HCHSP_Enrollment_Formatted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Processing error: {e}")

