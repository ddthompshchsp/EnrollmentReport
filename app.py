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
    logo = Image.open("header_logo.png")  # optional; safe to remove
    st.image(logo, width=300)
except Exception:
    pass

st.title("HCHSP Enrollment Report Formatter (2025–2026)")
st.divider()

# ---- File Uploads (stacked exactly as requested) ----
vf_file = st.file_uploader("Upload *VF_Average_Funded_Enrollment_Level.xlsx*", type=["xlsx"], key="vf")
aa_file = st.file_uploader("Upload *25-26 Applied/Accepted.xlsx*", type=["xlsx"], key="aa")

# ----------------------------
# Utilities
# ----------------------------
def parse_vf(vf_df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Parse VF report (header=None) into per-class rows:
    Center | Class | Funded | Enrolled

    Assumptions for the VF export:
    - Column A contains markers like "HCHSP -- Center Name", "Class X", and a "Class Total..." line per class.
    - On the 'Class Total' line:
        Enrolled is at column index 3,
        Funded is at column index 4.
    (If your export differs, update the indices below.)
    """
    records = []
    current_center = None
    current_class_name = None

    for i in range(len(vf_df_raw)):
        c0 = vf_df_raw.iloc[i, 0]

        # Center marker
        if isinstance(c0, str) and c0.strip().startswith("HCHSP --"):
            current_center = c0.strip()

        # Class marker, e.g. "Class 1", "Class 2A"
        elif isinstance(c0, str) and re.match(r"^\s*Class\s+\w+", c0, flags=re.I):
            # keep ONLY the class token (no "Class " prefix in output)
            current_class_name = re.sub(r"^\s*Class\s+", "", c0, flags=re.I).strip()

        # Totals row per class (where we read the numbers)
        if isinstance(c0, str) and c0.strip().lower().startswith("class total") and current_center and current_class_name:
            row = vf_df_raw.iloc[i]
            funded   = pd.to_numeric(row.iloc[4], errors="coerce")  # Number of Federal Slots Available
            enrolled = pd.to_numeric(row.iloc[3], errors="coerce")  # Number of Children Enrolled
            center_clean = re.sub(r"^HCHSP --\s*", "", current_center).strip()

            records.append({
                "Center": center_clean,
                "Class": current_class_name,   # <-- just the name (no "Class " prefix)
                "Funded": 0 if pd.isna(funded) else int(funded),
                "Enrolled": 0 if pd.isna(enrolled) else int(enrolled),
            })

    tidy = pd.DataFrame(records)
    if tidy.empty:
        raise ValueError("Could not find any class totals rows in the VF file. Check the file format.")
    return tidy


def parse_applied_accepted(aa_df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Parse Applied/Accepted report (header=None) to per-center counts.
    Excludes rows with a Status End Date.
    """
    idx = aa_df_raw.index[aa_df_raw.iloc[:, 0].astype(str).str.startswith("ST: Participant PID", na=False)]
    if len(idx) == 0:
        raise ValueError("Could not find header row in Applied/Accepted report (expected a row starting with 'ST: Participant PID').")
    header_row_idx = int(idx[0])

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
    counts = counts[["Accepted", "Applied"]].astype(int).reset_index().rename(columns={center_col: "Center"})
    return counts


def calc_waitlist_lacking(funded: int, enrolled: int) -> tuple[int, int]:
    """Return (waitlist, lacking) using your rule set."""
    if enrolled > funded:
        return enrolled - funded, 0
    elif funded > enrolled:
        return 0, funded - enrolled
    else:
        return 0, 0


def build_output_table(vf_tidy: pd.DataFrame, counts: pd.DataFrame) -> pd.DataFrame:
    """
    Build final table:
    - Class rows show Class (from VF), Funded, Enrolled, Waitlist, Lacking, %, and BLANK Applied/Accepted.
    - Center Total rows show sums + Applied/Accepted from the A/A file.
    - Agency Total final row shows overall sums.
    """
    merged = vf_tidy.merge(counts, on="Center", how="left").fillna({"Accepted": 0, "Applied": 0})

    rows = []
    for center, group in merged.groupby("Center", sort=True):
        # ---- Class rows ----
        for _, r in group.iterrows():
            waitlist, lacking = calc_waitlist_lacking(int(r["Funded"]), int(r["Enrolled"]))
            pct = int(round(r["Enrolled"] / r["Funded"] * 100, 0)) if r["Funded"] > 0 else pd.NA
            rows.append({
                "Center": r["Center"],
                "Class": r["Class"],   # exact class name from VF
                "Funded": int(r["Funded"]),
                "Enrolled": int(r["Enrolled"]),
                "Applied": "",         # blank on class rows
                "Accepted": "",        # blank on class rows
                "Waitlist": waitlist,
                "Lacking": lacking,
                "% Enrolled of Funded": pct
            })

        # ---- Center Total ----
        funded_sum   = int(group["Funded"].sum())
        enrolled_sum = int(group["Enrolled"].sum())
        wait_sum, lack_sum = calc_waitlist_lacking(funded_sum, enrolled_sum)
        pct_total = int(round(enrolled_sum / funded_sum * 100, 0)) if funded_sum > 0 else pd.NA

        rows.append({
            "Center": f"{center} Total",
            "Class": "",  # blank on totals
            "Funded": funded_sum,
            "Enrolled": enrolled_sum,
            "Applied": int(group["Applied"].max()),
            "Accepted": int(group["Accepted"].max()),
            "Waitlist": wait_sum,
            "Lacking": lack_sum,
            "% Enrolled of Funded": pct_total
        })

    final = pd.DataFrame(rows)

    # ---- Agency Total ----
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
    Styled Excel with:
    - Black title/subtitle
    - Blue header band
    - Filters, frozen header
    - % with percent sign on every row
    - % < 100 red, % > 100 blue (100 = black)
    - Lacking > 0 red
    - Bold totals
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

        # % column formatting
        percent_col_idx = df.columns.get_loc("% Enrolled of Funded")
        percent_fmt = wb.add_format({"num_format": '0"%"'})  # whole-number percent with % sign
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
        # (Exact 100% will remain default black)

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
            vf_raw = pd.read_excel(vf_file, sheet_name=0, header=None)
            aa_raw = pd.read_excel(aa_file, sheet_name=0, header=None)

            vf_tidy = parse_vf(vf_raw)               # gets Center, Class (as in VF), Funded, Enrolled
            aa_counts = parse_applied_accepted(aa_raw)  # gets Applied, Accepted per Center
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


