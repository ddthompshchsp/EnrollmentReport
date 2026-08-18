
import io
import re
from pathlib import Path
from datetime import datetime
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
        Upload the VF Average Funded Enrollment report and the 2026–2027 Applied/Accepted report.
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
    aa_file = st.file_uploader("Upload *2026-2027 Applied/Accepted.xlsx*", type=["xlsx"], key="aa")
    process = st.button("Process & Download")

# ----------------------------
# HARD-CODED LICENSED CAPACITY
# ----------------------------
# 2026-2027 license capacities.
# Slash values are stored as text so Excel displays them exactly as provided.
# Use None only for a center that should remain blank.
LIC_CAP = {
    "alvarez": 138,
    "arnold": 94,
    "castro": 88,
    "chapa": 112,
    "cesar chavez": 187,
    "donna ehs": "100/40",
    "garza": "188/34",
    "garza ehs academy": "188/34",
    "edinburg": 232,
    "edinburg north": 147,
    "escandon": 109,
    "farias": 132,
    "guerra": 144,
    "guzman": 373,
    "longoria": 151,
    "mercedes": 182,
    "mission ehs academy": "111/26",
    "monte alto": 100,
    "munoz": 88,
    "palacios": 135,
    "palmview": 153,
    "roosevelt": None,
    "salinas": 90,
    "sam fordyce": 97,
    "sam houston": 108,
    "san carlos": 105,
    "san juan ehs academy": "144/26",
    "seguin": 112,
    "singleterry": 204,
    "thigpen": 136,
    "wilson": 96,
}

# accept ASCII hyphen + Unicode dashes
DASH_CLASS = r"[-‐-‒–—]"

# ----------------------------
# Normalization / matching
# ----------------------------
def _norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()

def _canonicalize_center(s: str) -> str:
    if s is None:
        return ""
    txt = str(s)
    txt = re.sub(rf"^\s*HCHSP\s*{DASH_CLASS}{{1,}}\s*", "", txt, flags=re.I)  # strip "HCHSP — "
    txt = re.sub(r"\([^)]*\)", " ", txt)                                      # remove "(...)"
    txt = txt.lower()
    txt = re.sub(r"[^a-z0-9\s]", " ", txt)

    # These words are removed so center names match even when the source report includes
    # district names, school type, or descriptors.
    filler = {
        "head", "start", "headstart", "center", "campus", "elementary", "elem",
        "school", "program", "isd", "cisd", "psja", "mcallen", "mission",
        "donna", "edinburg", "mercedes", "academy", "hs"
    }
    tokens = [t for t in txt.split() if t and t not in filler]
    return " ".join(tokens).strip()

_CANON_TO_OFFICIAL = {_canonicalize_center(k): k for k in LIC_CAP}

# Extra aliases for names as they appear in the 2026-2027 classroom report.
CENTER_ALIASES = {
    "arnold": "arnold",
    "alvarez": "alvarez",
    "castro": "castro",
    "chapa": "chapa",
    "cesar chavez": "cesar chavez",
    "chavez": "cesar chavez",
    "donna ehs": "donna ehs",
    "donna": "donna ehs",
    "edinburg": "edinburg",
    "edinburg north": "edinburg north",
    "north": "edinburg north",
    "escandon": "escandon",
    "farias": "farias",
    "garza": "garza",
    "garza ehs": "garza ehs academy",
    "garza academy": "garza ehs academy",
    "guerra": "guerra",
    "guzman": "guzman",
    "longoria": "longoria",
    "mel": "mercedes",
    "mercedes": "mercedes",
    "mission ehs": "mission ehs academy",
    "mission": "mission ehs academy",
    "robert garate monte alto": "monte alto",
    "garate monte alto": "monte alto",
    "monte alto": "monte alto",
    "munoz": "munoz",
    "palacios": "palacios",
    "palmview": "palmview",
    "roosevelt": "roosevelt",
    "salinas mission": "salinas",
    "salinas": "salinas",
    "sam fordyce": "sam fordyce",
    "sam houston": "sam houston",
    "houston": "sam houston",
    "san carlos": "san carlos",
    "san juan ehs": "san juan ehs academy",
    "san juan": "san juan ehs academy",
    "seguin": "seguin",
    "singleterry": "singleterry",
    "thigpen": "thigpen",
    "thigpen zavala": "thigpen",
    "zavala": "thigpen",
    "wilson": "wilson",
}

# ----------------------------
# Program assignment
# ----------------------------
PROGRAM_HS = "Head Start"
PROGRAM_EHS = "Early Head Start"

FUNDED_TARGETS = {
    "All": 2720,
    PROGRAM_HS: 2480,
    PROGRAM_EHS: 240,
}

EHS_CENTER_PHRASES = (
    "donna early head start",
    "mission early head start",
    "san juan early head start",
    "donna ehs",
    "mission ehs",
    "san juan ehs",
)

def _plain_name(value) -> str:
    """Lowercase text with punctuation and repeated whitespace removed."""
    text = "" if value is None or pd.isna(value) else str(value)
    text = re.sub(rf"^\s*hchsp\s*{DASH_CLASS}{{1,}}\s*", "", text, flags=re.I)
    text = re.sub(r"[^a-zA-Z0-9\s]", " ", text).lower()
    return re.sub(r"\s+", " ", text).strip()

def program_for(center_name, class_name="") -> str:
    """
    Assign a center/classroom row to Head Start or Early Head Start.

    Donna, Mission, and San Juan Early Head Start are EHS centers. Garza is
    split at the classroom level: Infant/Toddler rooms are EHS and every other
    Garza room is Head Start.
    """
    center = _plain_name(center_name)
    classroom = _plain_name(class_name)

    if "garza" in center:
        return PROGRAM_EHS if re.search(r"\b(infant|toddler)\b", classroom) else PROGRAM_HS

    if any(phrase in center for phrase in EHS_CENTER_PHRASES):
        return PROGRAM_EHS

    return PROGRAM_HS

def lic_cap_for(center_name: str):
    if not isinstance(center_name, str):
        return ""

    raw = _norm_ws(center_name).lower()
    raw = re.sub(rf"^\s*hchsp\s*{DASH_CLASS}{{1,}}\s*", "", raw, flags=re.I)
    raw = re.sub(r"[^a-z0-9\s]", " ", raw)
    raw = re.sub(r"\s+", " ", raw).strip()

    # Check the most specific aliases first so, for example,
    # "San Juan EHS Academy" does not get matched too early by "san juan".
    for alias, official in sorted(CENTER_ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
        if alias in raw:
            val = LIC_CAP.get(official)
            return "" if val is None else val

    canon = _canonicalize_center(center_name)
    if canon in CENTER_ALIASES:
        val = LIC_CAP.get(CENTER_ALIASES[canon])
        return "" if val is None else val

    if canon in _CANON_TO_OFFICIAL:
        val = LIC_CAP[_CANON_TO_OFFICIAL[canon]]
        return "" if val is None else val

    best_key, best_len = None, 0
    for canon_k, off in _CANON_TO_OFFICIAL.items():
        if canon_k and (canon_k in canon or canon in canon_k):
            if len(canon_k) > best_len:
                best_key, best_len = off, len(canon_k)

    val = LIC_CAP.get(best_key) if best_key else ""
    return "" if val is None else val

# ----------------------------
# Helpers (parsing)
# ----------------------------
def _first_nonempty_strings(row, max_cols=8):
    vals = []
    for j in range(min(max_cols, row.shape[0])):
        v = row.iloc[j]
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s:
            vals.append(s)
    return vals

def _row_has_totals(cells_lower: list[str]) -> bool:
    joined = " | ".join(cells_lower)
    return (
        "class totals" in joined
        or "totals for class" in joined
        or re.search(r"\bclass\s*total", joined) is not None
    )

def _last_two_numbers(row):
    nums = []
    for v in row:
        x = pd.to_numeric(v, errors="coerce")
        if pd.notna(x):
            nums.append(float(x))
    if len(nums) >= 2:
        return nums[-2], nums[-1]
    return None, None

# ----------------------------
# Parsers
# ----------------------------
def parse_vf(vf_df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Output columns: Center | Class | Funded | Enrolled | PctRatio

    Rules:
    - Keep 'Class X' exactly as 'Class X'
    - Keep 'Home Based 01' exactly as 'Home Based 01'
    - Do NOT auto-add 'Class' to Home Based
    - Class names may repeat across centers (allowed)
    """
    records = []
    current_center = None
    current_class  = None

    re_center = re.compile(rf"^\s*HCHSP\s*{DASH_CLASS}+\s*(.+)$", re.I)

    # Matches: "Class PK3", "Class 101 (Spanish)"
    re_class = re.compile(r"^\s*(Class\s+(?!Totals).+)$", re.I)

    # Matches: "Home Based 01", "Home  Based   001 (EHS)"
    re_home_based = re.compile(r"^\s*(Home\s*Based\s*0*\d+\b.*)$", re.I)

    for i in range(len(vf_df_raw)):
        row = vf_df_raw.iloc[i, :]
        cells = _first_nonempty_strings(row, max_cols=8)
        if not cells:
            continue

        first = cells[0]
        lower_cells = [c.lower() for c in cells]

        # ---- Center header ----
        m_center = re_center.match(first)
        if m_center:
            current_center = _norm_ws(m_center.group(1))
            current_class = None
            continue

        # ---- Totals row ----
        if _row_has_totals(lower_cells) and current_center and current_class:
            enrolled = pd.to_numeric(row.iloc[3], errors="coerce")
            funded   = pd.to_numeric(row.iloc[4], errors="coerce")
            pct_ratio= pd.to_numeric(row.iloc[6], errors="coerce")

            if pd.isna(enrolled) or pd.isna(funded):
                e, f = _last_two_numbers(row)
                enrolled = e if e is not None else enrolled
                funded   = f if f is not None else funded

            records.append({
                "Center": current_center,
                "Class": current_class,   # ← EXACT label preserved
                "Funded": 0.0 if pd.isna(funded) else float(funded),
                "Enrolled": 0.0 if pd.isna(enrolled) else float(enrolled),
                "PctRatio": None if pd.isna(pct_ratio) else float(pct_ratio),
            })
            continue

        # ---- Class headers ----
        m_class = re_class.match(first)
        if m_class:
            current_class = _norm_ws(m_class.group(1))  # keeps "Class PK3"
            continue

        m_home = re_home_based.match(first)
        if m_home:
            current_class = _norm_ws(m_home.group(1))   # keeps "Home Based 01"
            continue

    tidy = pd.DataFrame(records)
    if tidy.empty:
        raise ValueError("Could not parse VF report (check class/center markers).")

    tidy["Center"] = tidy["Center"].map(_norm_ws)
    tidy["Program"] = tidy.apply(lambda r: program_for(r["Center"], r["Class"]), axis=1)
    return tidy



def parse_applied_accepted(aa_df_raw: pd.DataFrame) -> pd.DataFrame:
    header_row_idx = aa_df_raw.index[
        aa_df_raw.iloc[:, 0].astype(str).str.startswith("ST: Participant PID", na=False)
    ]
    if len(header_row_idx) == 0:
        raise ValueError("Could not find header row in Applied/Accepted report.")

    header_row_idx = int(header_row_idx[0])
    headers = aa_df_raw.iloc[header_row_idx].tolist()
    body = pd.DataFrame(aa_df_raw.iloc[header_row_idx + 1:].values, columns=headers)

    center_col = "ST: Center Name"
    status_col = "ST: Status"
    date_col = "ST: Status End Date"
    target_py_col = "APF: Target PY"

    required_cols = [center_col, status_col, date_col, target_py_col]
    missing_cols = [c for c in required_cols if c not in body.columns]
    if missing_cols:
        raise ValueError(f"Missing required column(s) in Applied/Accepted report: {', '.join(missing_cols)}")

    # Same original filter: only active rows with blank Status End Date.
    is_blank_date = body[date_col].isna() | body[date_col].astype(str).str.strip().eq("")

    # Fixed rule: Target PY must contain 2026-2027.
    # The report may show the program year as:
    # (Hidalgo County Head Start Program)2026-2027 (07/01/2026--06/30/2027)
    # so we use contains instead of an exact match.
    is_target_py = (
        body[target_py_col]
        .astype(str)
        .str.strip()
        .str.contains("2026-2027", na=False)
    )

    body = body[is_blank_date & is_target_py].copy()

    body[center_col] = (
        body[center_col]
        .astype(str)
        .str.replace(rf"^\s*HCHSP\s*{DASH_CLASS}{{1,}}\s*", "", regex=True)
        .map(_norm_ws)
    )

    # Use a classroom/room field when one is present. This is what allows the
    # Garza Applied/Accepted records to follow the same Infant/Toddler split as
    # the funded-enrollment report. If no such field exists, blank/non-matching
    # Garza records follow the stated rule and remain Head Start.
    room_col = None
    preferred_room_cols = (
        "ST: Classroom Name", "ST: Class Name", "ST: Classroom",
        "ST: Class", "ST: Room Name", "ST: Room",
    )
    for candidate in preferred_room_cols:
        if candidate in body.columns:
            room_col = candidate
            break
    if room_col is None:
        for col in body.columns:
            col_text = str(col).lower()
            if "classroom" in col_text or re.search(r"\broom\b", col_text):
                room_col = col
                break

    if room_col is None:
        body["Program"] = body[center_col].map(lambda c: program_for(c, ""))
    else:
        body["Program"] = body.apply(lambda r: program_for(r[center_col], r[room_col]), axis=1)

    counts = (
        body.groupby([center_col, "Program"])[status_col]
        .value_counts()
        .unstack(fill_value=0)
    )
    for c in ["Accepted", "Applied"]:
        if c not in counts.columns:
            counts[c] = 0
    return (
        counts[["Accepted", "Applied"]]
        .astype(int)
        .reset_index()
        .rename(columns={center_col: "Center"})
    )

# ----------------------------
# Builder
# ----------------------------
def build_output_table(
    vf_tidy: pd.DataFrame,
    counts: pd.DataFrame,
    funded_target: int,
    program_filter: str | None = None,
) -> pd.DataFrame:
    """Build the all-agency, Head Start-only, or EHS-only display table."""
    vf_work = vf_tidy.copy()
    counts_work = counts.copy()

    if program_filter is not None:
        vf_work = vf_work[vf_work["Program"] == program_filter].copy()
        counts_work = counts_work[counts_work["Program"] == program_filter].copy()

    # Main sheet combines the program-level Applied/Accepted counts back into
    # one count per center. Program sheets retain only their own counts.
    counts_by_center = (
        counts_work.groupby("Center", as_index=False)[["Accepted", "Applied"]].sum()
        if not counts_work.empty
        else pd.DataFrame(columns=["Center", "Accepted", "Applied"])
    )

    # Prefer percentage from the VF sheet when provided.
    if "PctRatio" in vf_work.columns and vf_work["PctRatio"].notna().any():
        vf_work["PctInt"] = pd.array((vf_work["PctRatio"] * 100).round(0), dtype="Int64")
    else:
        pct = (vf_work["Enrolled"] * 100).div(pd.Series(vf_work["Funded"]).replace(0, np.nan))
        vf_work["PctInt"] = pd.array(pct.round(0), dtype="Int64")

    merged = vf_work.merge(counts_by_center, on="Center", how="left").fillna({"Accepted": 0, "Applied": 0})
    applied_by_center = merged.groupby("Center")["Applied"].max() if not merged.empty else pd.Series(dtype=float)
    accepted_by_center = merged.groupby("Center")["Accepted"].max() if not merged.empty else pd.Series(dtype=float)

    rows = []
    waitlist_total = 0

    for center, group in merged.groupby("Center", sort=True):
        funded_sum = int(group["Funded"].sum())
        enrolled_sum = int(group["Enrolled"].sum())
        pct_total = int(round(enrolled_sum / funded_sum * 100, 0)) if funded_sum > 0 else pd.NA
        accepted_val = int(accepted_by_center.get(center, 0))
        applied_val = int(applied_by_center.get(center, 0))

        # If Enrolled >= Funded, set Waitlist = Accepted.
        waitlist_val = accepted_val if funded_sum > 0 and enrolled_sum >= funded_sum else ""
        lacking_over = funded_sum - enrolled_sum
        if waitlist_val != "":
            waitlist_total += waitlist_val

        rows.append({
            "Center": f"{center} Total",
            "Room#/Age/Lang": "",
            "Lic Cap.": lic_cap_for(center),
            "Funded": funded_sum,
            "Enrolled": enrolled_sum,
            "Applied": applied_val,
            "Accepted": accepted_val,
            "Lacking/Overage": lacking_over,
            "Waitlist": waitlist_val,
            "% Enrolled of Funded": pct_total,
            "Comments": "",
        })

        # Class labels are preserved exactly as parsed.
        for _, r in group.iterrows():
            rows.append({
                "Center": r["Center"],
                "Room#/Age/Lang": r["Class"],
                "Lic Cap.": "",
                "Funded": int(r["Funded"]),
                "Enrolled": int(r["Enrolled"]),
                "Applied": "",
                "Accepted": "",
                "Lacking/Overage": "",
                "Waitlist": "",
                "% Enrolled of Funded": int(r["PctInt"]) if pd.notna(r["PctInt"]) else pd.NA,
                "Comments": "",
            })

    center_rows = [r for r in rows if str(r["Center"]).endswith(" Total")]
    agency_enrolled = sum(int(r["Enrolled"]) for r in center_rows)
    agency_applied = sum(int(r["Applied"]) for r in center_rows)
    agency_accepted = sum(int(r["Accepted"]) for r in center_rows)
    agency_pct = int(round(agency_enrolled / funded_target * 100, 0)) if funded_target > 0 else pd.NA

    rows.append({
        "Center": "Agency Total",
        "Room#/Age/Lang": "",
        "Lic Cap.": "",
        "Funded": int(funded_target),
        "Enrolled": agency_enrolled,
        "Applied": agency_applied,
        "Accepted": agency_accepted,
        "Lacking/Overage": int(funded_target) - agency_enrolled,
        "Waitlist": waitlist_total,
        "% Enrolled of Funded": agency_pct,
        "Comments": "",
    })

    final = pd.DataFrame(rows)
    return final[[
        "Center", "Room#/Age/Lang", "Lic Cap.", "Funded", "Enrolled",
        "Applied", "Accepted", "Lacking/Overage", "Waitlist",
        "% Enrolled of Funded", "Comments",
    ]]

# ----------------------------
# Excel Writer (three program views + filter-responsive totals)
# ----------------------------
def to_styled_excel(sheet_specs: list[dict]) -> bytes:
    def idx_to_letter0(idx0: int) -> str:
        n = idx0
        letters = ""
        while True:
            n, remainder = divmod(n, 26)
            letters = chr(remainder + 65) + letters
            if n == 0:
                break
            n -= 1
        return letters

    def visible_center_total_formula(column_letter: str, data_end_excel: int) -> str:
        """Sum visible center-total rows only, excluding duplicated class rows."""
        return (
            f'=SUMPRODUCT(SUBTOTAL(109,OFFSET(${column_letter}$5,'
            f'ROW(${column_letter}$5:${column_letter}${data_end_excel})-ROW(${column_letter}$5),0,1)),'
            f'--(RIGHT($A$5:$A${data_end_excel},6)=" Total"))'
        )

    def visible_class_count_expression(data_end_excel: int) -> str:
        """Count visible classroom rows so either common Center filter style works."""
        return (
            f'SUMPRODUCT(SUBTOTAL(103,OFFSET($B$5,'
            f'ROW($B$5:$B${data_end_excel})-ROW($B$5),0,1)),'
            f'--($B$5:$B${data_end_excel}<>""))'
        )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        wb.set_calc_mode("auto")
        now_ct = datetime.now(ZoneInfo("America/Chicago"))
        date_str = now_ct.strftime("%m.%d.%y %I:%M %p CT")

        for spec in sheet_specs:
            df = spec["df"]
            sheet_name = spec["sheet_name"]
            subtitle = spec["subtitle"]
            funded_target = int(spec["funded_target"])

            df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=3)
            ws = writer.sheets[sheet_name]

            ws.hide_gridlines(0)
            ws.set_row(0, 24)
            ws.set_row(1, 22)
            ws.set_row(2, 20)
            ws.freeze_panes(4, 0)

            if logo_path.exists():
                ws.set_column(1, 1, 6)
                ws.insert_image(0, 1, str(logo_path), {
                    "x_offset": 2,
                    "y_offset": 2,
                    "x_scale": 0.53,
                    "y_scale": 0.53,
                    "object_position": 1,
                })

            title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
            subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
            red_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": "#C00000"})
            header_fmt = wb.add_format({
                "bold": True,
                "font_color": "white",
                "bg_color": "#305496",
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "border": 1,
            })
            border_all = wb.add_format({"border": 1})
            bold_row = wb.add_format({"bold": True})
            comments_fmt = wb.add_format({
                "border": 1,
                "text_wrap": True,
                "valign": "top",
                "align": "left",
            })
            agency_number_fmt = wb.add_format({"bold": True, "border": 1, "align": "center"})
            agency_pct_fmt = wb.add_format({
                "bold": True,
                "border": 1,
                "align": "center",
                "num_format": '0"%"',
            })

            last_col_0 = len(df.columns) - 1
            last_col_letter = idx_to_letter0(last_col_0)

            ws.merge_range(0, 2, 0, last_col_0, "Hidalgo County Head Start Program", title_fmt)
            ws.merge_range(1, 2, 1, last_col_0, "", subtitle_fmt)
            ws.write_rich_string(
                1,
                2,
                subtitle_fmt,
                f"{subtitle} - 2026-2027 Campus Classroom Enrollment as of ",
                red_fmt,
                f"({date_str})",
                subtitle_fmt,
            )

            ws.set_row(3, 26)
            for column_index, column_name in enumerate(df.columns):
                ws.write(3, column_index, column_name, header_fmt)

            widths = {
                "Center": 28,
                "Room#/Age/Lang": 22,
                "Lic Cap.": 10,
                "Funded": 12,
                "Enrolled": 12,
                "Applied": 12,
                "Accepted": 12,
                "Lacking/Overage": 14,
                "Waitlist": 12,
                "% Enrolled of Funded": 16,
                "Comments": 48,
            }
            for column_name, width in widths.items():
                if column_name in df.columns:
                    column_index = df.columns.get_loc(column_name)
                    ws.set_column(column_index, column_index, width)

            agency_idx = int(df.index[df["Center"] == "Agency Total"][0])
            agency_ws_row = agency_idx + 4
            agency_excel_row = agency_ws_row + 1
            data_end_ws_row = agency_ws_row - 1
            data_end_excel = data_end_ws_row + 1
            last_excel_row = len(df) + 4
            center_total_idxs = [
                i for i, name in enumerate(df["Center"].tolist())
                if isinstance(name, str) and name.endswith(" Total") and name != "Agency Total"
            ]

            # Hidden helper values live only on classroom rows. They let Agency
            # Total work whether the Center filter selects "Garza", searches
            # for all Garza values, or selects only "Garza Total".
            helper_start_col = len(df.columns)
            helper_names = [
                "_Dynamic Funded", "_Dynamic Enrolled", "_Dynamic Applied",
                "_Dynamic Accepted", "_Dynamic Waitlist",
            ]
            for helper_offset, helper_name in enumerate(helper_names):
                ws.write(3, helper_start_col + helper_offset, helper_name)
                ws.set_column(helper_start_col + helper_offset, helper_start_col + helper_offset, None, None, {"hidden": True})
            for data_index in range(agency_idx):
                for helper_offset in range(len(helper_names)):
                    ws.write_number(4 + data_index, helper_start_col + helper_offset, 0)
            for group_index, total_idx in enumerate(center_total_idxs):
                next_total_idx = (
                    center_total_idxs[group_index + 1]
                    if group_index + 1 < len(center_total_idxs)
                    else agency_idx
                )
                class_idxs = [
                    idx for idx in range(total_idx + 1, next_total_idx)
                    if str(df.loc[idx, "Room#/Age/Lang"]).strip()
                ]
                for class_idx in class_idxs:
                    ws.write_number(4 + class_idx, helper_start_col, int(df.loc[class_idx, "Funded"]))
                    ws.write_number(4 + class_idx, helper_start_col + 1, int(df.loc[class_idx, "Enrolled"]))
                if class_idxs:
                    first_class_idx = class_idxs[0]
                    ws.write_number(4 + first_class_idx, helper_start_col + 2, int(df.loc[total_idx, "Applied"]))
                    ws.write_number(4 + first_class_idx, helper_start_col + 3, int(df.loc[total_idx, "Accepted"]))
                    waitlist_value = df.loc[total_idx, "Waitlist"]
                    ws.write_number(
                        4 + first_class_idx,
                        helper_start_col + 4,
                        int(waitlist_value) if waitlist_value != "" else 0,
                    )

            # Exclude Agency Total from the filter range so it remains visible.
            if data_end_ws_row >= 4:
                ws.autofilter(3, 0, data_end_ws_row, last_col_0)

            ws.conditional_format(f"A4:{last_col_letter}{last_excel_row}", {
                "type": "formula",
                "criteria": "TRUE",
                "format": border_all,
            })

            pct_idx = df.columns.get_loc("% Enrolled of Funded")
            pct_letter = idx_to_letter0(pct_idx)
            pct_range = f"{pct_letter}5:{pct_letter}{last_excel_row}"
            ws.conditional_format(pct_range, {
                "type": "cell", "criteria": "<", "value": 100,
                "format": wb.add_format({"font_color": "red"}),
            })
            ws.conditional_format(pct_range, {
                "type": "cell", "criteria": ">", "value": 100,
                "format": wb.add_format({"font_color": "blue"}),
            })
            ws.conditional_format(pct_range, {
                "type": "formula", "criteria": "TRUE",
                "format": wb.add_format({"num_format": '0"%"', "align": "center"}),
            })

            for row_index, center_name in enumerate(df["Center"].tolist()):
                if isinstance(center_name, str) and center_name.endswith(" Total"):
                    ws.set_row(row_index + 4, None, bold_row)

            # Keep the target when all schools are visible. When a filter hides
            # rows, Funded and the other Agency Total values recalculate from
            # only the visible center-total rows.
            if data_end_excel >= 5:
                no_filter_test = f'SUBTOTAL(103,$A$5:$A${data_end_excel})=COUNTA($A$5:$A${data_end_excel})'
                class_count = visible_class_count_expression(data_end_excel)
                helper_letters = [idx_to_letter0(helper_start_col + i) for i in range(len(helper_names))]

                def filtered_total_formula(display_letter: str, helper_letter: str) -> str:
                    center_total_sum = visible_center_total_formula(display_letter, data_end_excel)[1:]
                    return (
                        f'=IF({class_count}>0,'
                        f'SUBTOTAL(109,${helper_letter}$5:${helper_letter}${data_end_excel}),'
                        f'{center_total_sum})'
                    )

                filtered_funded = filtered_total_formula("D", helper_letters[0])
                funded_formula = (
                    f'=IF({no_filter_test},{funded_target},'
                    f'{filtered_funded[1:]})'
                )
                enrolled_formula = filtered_total_formula("E", helper_letters[1])
                applied_formula = filtered_total_formula("F", helper_letters[2])
                accepted_formula = filtered_total_formula("G", helper_letters[3])
                waitlist_formula = filtered_total_formula("I", helper_letters[4])

                cached = df.loc[agency_idx]
                ws.write_formula(agency_ws_row, 3, funded_formula, agency_number_fmt, int(cached["Funded"]))
                ws.write_formula(agency_ws_row, 4, enrolled_formula, agency_number_fmt, int(cached["Enrolled"]))
                ws.write_formula(agency_ws_row, 5, applied_formula, agency_number_fmt, int(cached["Applied"]))
                ws.write_formula(agency_ws_row, 6, accepted_formula, agency_number_fmt, int(cached["Accepted"]))
                ws.write_formula(
                    agency_ws_row,
                    7,
                    f"=D{agency_excel_row}-E{agency_excel_row}",
                    agency_number_fmt,
                    int(cached["Lacking/Overage"]),
                )
                ws.write_formula(agency_ws_row, 8, waitlist_formula, agency_number_fmt, int(cached["Waitlist"]))
                ws.write_formula(
                    agency_ws_row,
                    9,
                    f'=IFERROR(ROUND(E{agency_excel_row}/D{agency_excel_row}*100,0),"")',
                    agency_pct_fmt,
                    int(cached["% Enrolled of Funded"]),
                )

            # Merge one Comments area for each center, preserving the existing layout.
            comments_col_idx = df.columns.get_loc("Comments")
            for group_index, start_idx in enumerate(center_total_idxs):
                next_start = (
                    center_total_idxs[group_index + 1]
                    if group_index + 1 < len(center_total_idxs)
                    else agency_idx
                )
                end_idx = next_start - 1
                r0 = 4 + start_idx
                r1 = max(r0, 4 + end_idx)
                ws.merge_range(r0, comments_col_idx, r1, comments_col_idx, "", comments_fmt)

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
        all_df = build_output_table(vf_tidy, aa_counts, FUNDED_TARGETS["All"])
        hs_df = build_output_table(vf_tidy, aa_counts, FUNDED_TARGETS[PROGRAM_HS], PROGRAM_HS)
        ehs_df = build_output_table(vf_tidy, aa_counts, FUNDED_TARGETS[PROGRAM_EHS], PROGRAM_EHS)

        sheet_specs = [
            {
                "sheet_name": "Head Start Enrollment",
                "subtitle": "Head Start/EHS",
                "funded_target": FUNDED_TARGETS["All"],
                "df": all_df,
            },
            {
                "sheet_name": "Head Start Only",
                "subtitle": "Head Start",
                "funded_target": FUNDED_TARGETS[PROGRAM_HS],
                "df": hs_df,
            },
            {
                "sheet_name": "Early Head Start",
                "subtitle": "Early Head Start",
                "funded_target": FUNDED_TARGETS[PROGRAM_EHS],
                "df": ehs_df,
            },
        ]

        st.success("Three worksheet previews are below. Use the download button to get the Excel file.")
        preview_tabs = st.tabs(["All (2,720)", "Head Start (2,480)", "Early Head Start (240)"])
        for tab, spec in zip(preview_tabs, sheet_specs):
            with tab:
                preview_df = spec["df"].copy()
                pct_col = "% Enrolled of Funded"
                preview_df[pct_col] = preview_df[pct_col].apply(
                    lambda value: "" if pd.isna(value) else f"{int(value)}%"
                )
                st.dataframe(preview_df, use_container_width=True)

        xlsx_bytes = to_styled_excel(sheet_specs)
        st.download_button(
            "Download Formatted Excel",
            data=xlsx_bytes,
            file_name="HCHSP_Enrollment_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Processing error: {e}")
