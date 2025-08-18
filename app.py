import io
import re
from pathlib import Path
from datetime import date
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="HCHSP Enrollment", layout="wide")

# =========================
# Centered header (logo + title on the app only)
# =========================
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
        "<p style='text-align:center; font-size:16px; margin-top:0;'>"
        "Upload the VF Average Funded Enrollment report and the 25–26 Applied/Accepted report. "
        "This produces a styled Excel with titles, bold headers, filters, center totals, an agency total, "
        "and color highlighting for percentages.</p>",
        unsafe_allow_html=True,
    )

st.divider()

# =========================
# Centered inputs
# =========================
vf_file = None
aa_file = None
process = False
inp_l, inp_c, inp_r = st.columns([1, 2, 1])
with inp_c:
    vf_file = st.file_uploader("Upload *VF_Average_Funded_Enrollment_Level.xlsx*", type=["xlsx"], key="vf")
    aa_file = st.file_uploader("Upload *25-26 Applied/Accepted.xlsx*", type=["xlsx"], key="aa")
    process = st.button("Process & Download")

# ----------------------------
# Utilities
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
    Parse Applied/Accepted (header=None) to per-center counts.
    KEEP ONLY rows with a BLANK 'ST: Status End Date' (NaN or empty/whitespace).
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

    is_blank_date = body[date_col].isna() | body[date_col].astype(str).str.strip().eq("")
    body = body[is_blank_date].copy()

    body[center_col] = body[center_col].astype(str).str.replace(r"^HCHSP --\s*", "", regex=True)

    counts = body.groupby(center_col)[status_col].value_counts().unstack(fill_value=0)
    for c in ["Accepted", "Applied"]:
        if c not in counts.columns:
            counts[c] = 0

    counts = counts[["Accepted", "Applied"]].astype(int).reset_index().rename(columns={center_col:"Center"})
    return counts


def build_output_table(vf_tidy: pd.DataFrame, counts: pd.DataFrame) -> pd.DataFrame:
    """
    Merge class rows with center counts; add Center Totals and Agency Total.
    Class rows: Applied/Accepted/Waitlist/Lacking blank.
    Totals rows: show Applied/Accepted + Waitlist + Lacking.
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
                "Waitlist": "",
                "Lacking": "",
                "Applied": "",
                "Accepted": "",
                "% Enrolled of Funded": int(r["% Enrolled of Funded"]) if pd.notna(r["% Enrolled of Funded"]) else pd.NA
            })

        funded_sum = int(group["Funded"].sum())
        enrolled_sum = int(group["Enrolled"].sum())
        pct_total = int(round(enrolled_sum / funded_sum * 100, 0)) if funded_sum > 0 else pd.NA

        waitlist_total = enrolled_sum - funded_sum
        waitlist_total = "" if waitlist_total <= 0 else int(waitlist_total)

        lacking_total = funded_sum - enrolled_sum
        lacking_total = "" if lacking_total <= 0 else int(lacking_total)

        rows.append({
            "Center": f"{center} Total",
            "Class": "",
            "Funded": funded_sum,
            "Enrolled": enrolled_sum,
            "Waitlist": waitlist_total,
            "Lacking": lacking_total,
            "Applied": int(applied_by_center.get(center, 0)),
            "Accepted": int(accepted_by_center.get(center, 0)),
            "% Enrolled of Funded": pct_total
        })

    final = pd.DataFrame(rows)

    # Agency totals computed directly from filtered center-level counts
    agency_funded   = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Funded"].sum())
    agency_enrolled = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Enrolled"].sum())
    agency_applied  = int(counts["Applied"].sum())
    agency_accepted = int(counts["Accepted"].sum())
    agency_pct = int(round(agency_enrolled / agency_funded * 100, 0)) if agency_funded > 0 else pd.NA

    agency_waitlist = agency_enrolled - agency_funded
    agency_waitlist = "" if agency_waitlist <= 0 else int(agency_waitlist)

    agency_lacking = agency_funded - agency_enrolled
    agency_lacking = "" if agency_lacking <= 0 else int(agency_lacking)

    final = pd.concat([final, pd.DataFrame([{
        "Center": "Agency Total",
        "Class": "",
        "Funded": agency_funded,
        "Enrolled": agency_enrolled,
        "Waitlist": agency_waitlist,
        "Lacking": agency_lacking,
        "Applied": agency_applied,
        "Accepted": agency_accepted,
        "% Enrolled of Funded": agency_pct
    }])], ignore_index=True)

    # Final column order
    final = final[[
        "Center","Class","Funded","Enrolled","Waitlist","Lacking",
        "Applied","Accepted","% Enrolled of Funded"
    ]]
    return final


def to_styled_excel(df: pd.DataFrame) -> bytes:
    """Write styled Excel with boxed title+table; rest of sheet normal."""
    # Helpers for A1 ranges
    def col_letter(n: int) -> str:
        s = ""
        while n >= 0:
            s = chr(n % 26 + 65) + s
            n = n // 26 - 1
        return s

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Formatted", startrow=3)
        wb = writer.book
        ws = writer.sheets["Formatted"]

        # Titles with today's date in M.D.YY (IN the box)
        d = date.today()
        date_str = f"{d.month}.{d.day}.{str(d.year % 100).zfill(2)}"
        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center", "bg_color": "#FFFFFF"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center", "bg_color": "#FFFFFF"})
        ws.merge_range(0, 0, 0, len(df.columns)-1, "Hidalgo County Head Start Program", title_fmt)
        ws.merge_range(1, 0, 1, len(df.columns)-1, f"2025-2026 Campus Classroom Enrollment — {date_str}", subtitle_fmt)

        # Header bar (dark blue)
        header_fmt = wb.add_format({
            "bold": True, "font_color": "white", "bg_color": "#305496",
            "align": "center", "valign": "vcenter", "text_wrap": True
        })
        ws.set_row(3, 26)
        for c, col in enumerate(df.columns):
            ws.write(3, c, col, header_fmt)

        # Compute ranges
        last_row_0 = len(df) + 3                # 0-based last data row
        last_col_0 = len(df.columns) - 1
        first_row_excel = 1                     # include title row 1 in box
        first_data_excel = 5                    # first data row number
        last_excel_row = last_row_0 + 1         # convert to Excel numbering
        first_col_letter = "A"
        last_col_letter = col_letter(last_col_0)

        # Filters + freeze header (still at header)
        ws.autofilter(3, 0, last_row_0, last_col_0)
        ws.freeze_panes(4, 0)

        # Base formats (white fill inside the box to suppress gridlines)
        pct_idx = df.columns.get_loc("% Enrolled of Funded")
        pct_fmt = wb.add_format({'num_format': '0"%"', 'align': 'center', 'bg_color': '#FFFFFF'})
        int_fmt = wb.add_format({'num_format': '0', 'align': 'right', 'bg_color': '#FFFFFF'})
        text_fmt = wb.add_format({'align': 'left', 'bg_color': '#FFFFFF'})

        # Column widths & defaults (apply white fill as default inside box)
        ws.set_column(0, 0, 28, text_fmt)  # Center
        ws.set_column(1, 1, 14, text_fmt)  # Class
        for name, width in [("Funded", 12), ("Enrolled", 12), ("Waitlist", 12),
                            ("Lacking", 12), ("Applied", 12), ("Accepted", 12)]:
            idx = df.columns.get_loc(name)
            ws.set_column(idx, idx, width, int_fmt)
        ws.set_column(pct_idx, pct_idx, 16, pct_fmt)

        # Zebra striping on data rows ONLY (inside the box)
        data_range  = f"{first_col_letter}{first_data_excel}:{last_col_letter}{last_excel_row}"
        band_fmt = wb.add_format({"bg_color": "#F2F2F2"})
        ws.conditional_format(data_range, {
            "type": "formula",
            "criteria": '=MOD(ROW(),2)=0',
            "format": band_fmt
        })

        # % column colors (inside the box)
        pct_letter = col_letter(pct_idx)
        pct_range = f"{pct_letter}{first_data_excel}:{pct_letter}{last_excel_row}"
        ws.conditional_format(pct_range, {
            "type": "cell", "criteria": "<", "value": 100,
            "format": wb.add_format({"font_color": "red"})
        })
        ws.conditional_format(pct_range, {
            "type": "cell", "criteria": ">", "value": 100,
            "format": wb.add_format({"font_color": "blue"})
        })

        # Total rows (inside the box)
        total_fmt = wb.add_format({"bold": True, "bg_color": "#D9D9D9"})
        for ridx, val in enumerate(df["Center"].tolist()):
            if isinstance(val, str) and (val.endswith(" Total") or val == "Agency Total"):
                ws.set_row(ridx + 4, None, total_fmt)

        # ===== Box around title + header + data (everything we formatted) =====
        top_border    = wb.add_format({"top": 2})
        bottom_border = wb.add_format({"bottom": 2})
        left_border   = wb.add_format({"left": 2})
        right_border  = wb.add_format({"right": 2})

        # Apply to edges unconditionally so merges/blanks are included
        # Top (row 1), Bottom (last data row)
        ws.conditional_format(f"{first_col_letter}{first_row_excel}:{last_col_letter}{first_row_excel}",
                              {"type": "formula", "criteria": "TRUE", "format": top_border})
        ws.conditional_format(f"{first_col_letter}{last_excel_row}:{last_col_letter}{last_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": bottom_border})
        # Left and Right edges from row 1 through last data row
        ws.conditional_format(f"{first_col_letter}{first_row_excel}:{first_col_letter}{last_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": left_border})
        ws.conditional_format(f"{last_col_letter}{first_row_excel}:{last_col_letter}{last_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": right_border})

        # Note: We DID NOT hide gridlines, so everything OUTSIDE the box remains normal.

    return output.getvalue()


# ----------------------------
# Main
# ----------------------------
if process and vf_file and aa_file:
    try:
        vf_raw = pd.read_excel(vf_file, sheet_name=0, header=None)
        aa_raw = pd.read_excel(aa_file, sheet_name=0, header=None)

        vf_tidy = parse_vf(vf_raw)
        aa_counts = parse_applied_accepted(aa_raw)
        final_df = build_output_table(vf_tidy, aa_counts)

        # Preview with % signs on the % column (all rows, including totals)
        st.success("Preview below. Use the download button to get the Excel file.")
        preview_df = final_df.copy()
        pct_col = "% Enrolled of Funded"
        preview_df[pct_col] = preview_df[pct_col].apply(lambda v: "" if pd.isna(v) else f"{int(v)}%")
        st.dataframe(preview_df, use_container_width=True)

        # Export WITHOUT logo
        xlsx_bytes = to_styled_excel(final_df)
        st.download_button(
            "Download Formatted Excel",
            data=xlsx_bytes,
            file_name="HCHSP_Enrollment_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Processing error: {e}")



