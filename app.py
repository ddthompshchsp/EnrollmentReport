import io
import re
from pathlib import Path
from datetime import datetime, date
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="HCHSP Enrollment", layout="wide")

# ----------------------------
# Header (Streamlit UI only)
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
    process = st.button("Process & Download")

# ----------------------------
# Parsers
# ----------------------------
def parse_vf(vf_df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Parse VF report (header=None) into per-class rows:
    Center | Class | Funded | Enrolled
    Supports class codes with letters (e.g., A01, E117).
    """
    records = []
    current_center = None
    current_class = None

    re_center = re.compile(r"^\s*HCHSP\s*--\s*(.+)$", re.I)
    re_class  = re.compile(r"^\s*Class\s+([A-Za-z0-9\-]+)\s*$", re.I)

    for i in range(len(vf_df_raw)):
        c0 = vf_df_raw.iloc[i, 0]
        c0_str = c0 if isinstance(c0, str) else str(c0)

        m_center = re_center.match(c0_str)
        if m_center:
            current_center = m_center.group(1).strip()
            continue

        m_class = re_class.match(c0_str)
        if m_class:
            current_class = m_class.group(1).strip()
            continue

        if isinstance(c0, str) and c0.strip().lower() == "class totals:" and current_center and current_class:
            row = vf_df_raw.iloc[i]
            enrolled = pd.to_numeric(row.iloc[3], errors="coerce")
            funded   = pd.to_numeric(row.iloc[4], errors="coerce")

            records.append({
                "Center": current_center,
                "Class": f"Class {current_class}",
                "Funded": 0.0 if pd.isna(funded) else float(funded),
                "Enrolled": 0.0 if pd.isna(enrolled) else float(enrolled),
            })

    tidy = pd.DataFrame(records)
    if tidy.empty:
        raise ValueError("Could not parse VF report (check class/center markers and column indices).")
    return tidy


def parse_applied_accepted(aa_df_raw: pd.DataFrame) -> pd.DataFrame:
    header_row_idx = aa_df_raw.index[aa_df_raw.iloc[:, 0].astype(str).str.startswith("ST: Participant PID", na=False)]
    if len(header_row_idx) == 0:
        raise ValueError("Could not find header row in Applied/Accepted report.")
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

    return counts[["Accepted", "Applied"]].astype(int).reset_index().rename(columns={center_col: "Center"})

# ----------------------------
# Builder
# ----------------------------
def build_output_table(vf_tidy: pd.DataFrame, counts: pd.DataFrame) -> pd.DataFrame:
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

    for center, group in merged.groupby("Center", sort=True):
        # Totals first
        funded_sum   = int(group["Funded"].sum())
        enrolled_sum = int(group["Enrolled"].sum())
        pct_total    = int(round(enrolled_sum / funded_sum * 100, 0)) if funded_sum > 0 else pd.NA
        accepted_val = int(accepted_by_center.get(center, 0))
        applied_val  = int(applied_by_center.get(center, 0))
        waitlist_val = accepted_val if enrolled_sum > funded_sum else ""
        lacking_over = funded_sum - enrolled_sum
        if waitlist_val != "": 
            waitlist_totals += waitlist_val

        rows.append({
            "Center": f"{center} Total", "Class": "",
            "Funded": funded_sum, "Enrolled": enrolled_sum,
            "Applied": applied_val, "Accepted": accepted_val,
            "Lacking/Overage": lacking_over, "Waitlist": waitlist_val,
            "% Enrolled of Funded": pct_total
        })

        # Then the classes
        for _, r in group.iterrows():
            rows.append({
                "Center": r["Center"], "Class": r["Class"],
                "Funded": int(r["Funded"]), "Enrolled": int(r["Enrolled"]),
                "Applied": "", "Accepted": "", "Lacking/Overage": "", "Waitlist": "",
                "% Enrolled of Funded": int(r["% Enrolled of Funded"]) if pd.notna(r["% Enrolled of Funded"]) else pd.NA
            })

    final = pd.DataFrame(rows)

    # Agency total at the bottom
    agency_funded   = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Funded"].sum())
    agency_enrolled = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Enrolled"].sum())
    agency_applied  = int(counts["Applied"].sum())
    agency_accepted = int(counts["Accepted"].sum())
    agency_pct      = int(round(agency_enrolled / agency_funded * 100, 0)) if agency_funded > 0 else pd.NA
    agency_lacking  = agency_funded - agency_enrolled

    final = pd.concat([final, pd.DataFrame([{
        "Center": "Agency Total", "Class": "",
        "Funded": agency_funded, "Enrolled": agency_enrolled,
        "Applied": agency_applied, "Accepted": agency_accepted,
        "Lacking/Overage": agency_lacking, "Waitlist": waitlist_totals,
        "% Enrolled of Funded": agency_pct
    }])], ignore_index=True)

    return final[[
        "Center","Class","Funded","Enrolled","Applied","Accepted","Lacking/Overage","Waitlist","% Enrolled of Funded"
    ]]

# ----------------------------
# Excel Writer
# ----------------------------
def to_styled_excel(df: pd.DataFrame) -> bytes:
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

        ws.hide_gridlines(0)

        ws.set_row(0, 24)
        ws.set_row(1, 22)
        ws.set_row(2, 20)

        # Logo
        logo = Path("header_logo.png")
        if logo.exists():
            ws.set_column(1, 1, 6)
            ws.insert_image(0, 1, str(logo), {
                "x_offset": 2, "y_offset": 2,
                "x_scale": 0.53, "y_scale": 0.53,
                "object_position": 1
            })

        # Titles with date + time (Central)
        now_ct = datetime.now(ZoneInfo("America/Chicago"))
        date_str = now_ct.strftime("%m.%d.%y %I:%M %p CT")

        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        red_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": "#C00000"})

        last_col_0 = len(df.columns) - 1
        last_col_letter = col_letter(last_col_0)

        ws.merge_range(0, 2, 0, last_col_0, "Hidalgo County Head Start Program", title_fmt)
        ws.merge_range(1, 2, 1, last_col_0, "", subtitle_fmt)
        ws.write_rich_string(1, 2,
            subtitle_fmt, "Head Start - 2025-2026 Campus Classroom Enrollment as of ",
            red_fmt, f"({date_str})",
            subtitle_fmt
        )

        # --- Header (blue) ---
        header_fmt = wb.add_format({
            "bold": True, "font_color": "white", "bg_color": "#305496",
            "align": "center", "valign": "vcenter", "text_wrap": True,
            "border": 1
        })
        ws.set_row(3, 26)
        for c, col in enumerate(df.columns):
            ws.write(3, c, col, header_fmt)

        last_row_0 = len(df) + 3
        last_excel_row = last_row_0 + 1
        ws.autofilter(3, 0, last_row_0, last_col_0)

        widths = {"Center": 28, "Class": 14, "Funded": 12, "Enrolled": 12,
                  "Applied": 12, "Accepted": 12, "Lacking/Overage": 14, "Waitlist": 12}
        for name, width in widths.items():
            if name in df.columns:
                idx = df.columns.get_loc(name)
                ws.set_column(idx, idx, width)
        ws.set_column(df.columns.get_loc("% Enrolled of Funded"), df.columns.get_loc("% Enrolled of Funded"), 16)

        border_all = wb.add_format({"border": 1})
        ws.conditional_format(f"A4:{last_col_letter}{last_excel_row}", {"type": "formula", "criteria": "TRUE", "format": border_all})

        pct_idx = df.columns.get_loc("% Enrolled of Funded")
        pct_letter = col_letter(pct_idx)
        pct_range = f"{pct_letter}5:{pct_letter}{last_excel_row}"
        ws.conditional_format(pct_range, {"type": "cell", "criteria": "<", "value": 100, "format": wb.add_format({"font_color": "red"})})
        ws.conditional_format(pct_range, {"type": "cell", "criteria": ">", "value": 100, "format": wb.add_format({"font_color": "blue"})})
        ws.conditional_format(pct_range, {"type": "formula", "criteria": "TRUE", "format": wb.add_format({'num_format': '0"%"', 'align': 'center'})})

        bold_row = wb.add_format({"bold": True})
        for ridx, val in enumerate(df["Center"].tolist()):
            if isinstance(val, str) and (val.endswith(" Total") or val == "Agency Total"):
                ws.set_row(ridx + 4, None, bold_row)

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

        st.success("Preview below. Use the download button to get the Excel file.")
        preview_df = final_df.copy()
        pct_col = "% Enrolled of Funded"
        preview_df[pct_col] = preview_df[pct_col].apply(lambda v: "" if pd.isna(v) else f"{int(v)}%")
        preview_df = preview_df[["Center","Class","Funded","Enrolled","Applied","Accepted","Lacking/Overage","Waitlist",pct_col]]
        st.dataframe(preview_df, use_container_width=True)

        xlsx_bytes = to_styled_excel(final_df)
        st.download_button(
            "Download Formatted Excel",
            data=xlsx_bytes,
            file_name="HCHSP_Enrollment_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Processing error: {e}")


