
}
# Force overwrite all rows for matching centers (recommended while debugging)
FORCE_REPLACE_ALL_ROWS_FOR_MATCHING_CENTERS = True

# =========================
# Header / UI
# =========================
try:
    logo = Image.open("header_logo.png")
    st.image(logo, width=300)
except Exception:
    pass

st.title("HCHSP Enrollment Checklist Formatter (2025–2026)")
st.markdown("Upload your **Enrollment.xlsx**. This version **hard-codes** lettered class names for specific centers, replacing any 'Class 30' leftovers.")

uploaded_file = st.file_uploader("Upload Enrollment.xlsx", type=["xlsx"])

# =========================
# Helpers
# =========================
def coerce_to_dt(v):
    if pd.isna(v):
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return from_excel(v)
        except Exception:
            return None
    if isinstance(v, str):
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip(), fmt)
            except Exception:
                continue
    return None

def most_recent(series):
    dates, texts = [], []
    for v in pd.unique(series.dropna()):
        dt = coerce_to_dt(v)
        if dt:
            dates.append(dt)
        else:
            texts.append(v)
    if dates:
        return max(dates)
    return texts[0] if texts else None

def find_header_row(ws, probe="ST: Participant PID", search_rows=160):
    for row in ws.iter_rows(min_row=1, max_row=search_rows):
        for cell in row:
            if isinstance(cell.value, str) and probe in cell.value:
                return cell.row
    return None

def norm_center(x: str) -> str:
    x = (x or "").lower().strip()
    x = re.sub(r'[\u2013\u2014\-]+', '-', x)     # normalize dashes
    x = re.sub(r'[^a-z0-9\s\-]', '', x)         # strip punctuation
    x = re.sub(r'\s+', ' ', x)
    for tail in [" isd"," elementary"," elem"," school"]:
        if x.endswith(tail):
            x = x[: -len(tail)]
    return x.strip()

# normalized copy of hard-coded map (CRITICAL)
HARD_CODED_CLASSES_N: Dict[str, List[str]] = { norm_center(k): v for k, v in HARD_CODED_CLASSES.items() }

# broader detection: ANY column whose header contains "class" (minus well-known non-class fields)
EXCLUDE_CLASSY = ("classification", "class size", "capacity")
def find_all_class_columns(cols):
    picks = []
    for c in cols:
        if not isinstance(c, str):
            continue
        low = c.lower().strip()
        if "class" in low and not any(bad in low for bad in EXCLUDE_CLASSY):
            picks.append(c)
    # if none found, create a new "Class Name"
    if not picks:
        picks = ["Class Name"]
    # dedupe keep order
    seen, keep = set(), []
    for c in picks:
        if c not in seen:
            seen.add(c); keep.append(c)
    return keep

def find_center_column(cols):
    for name in ["Center Name","Site","Campus","Center","ST: Center Name"]:
        if name in cols: return name
    for c in cols:
        if isinstance(c, str) and "center" in c.lower():
            return c
    return None

def deterministic_assign(block: pd.DataFrame, classes: List[str], pid_col="Participant PID") -> pd.Series:
    if len(classes) == 0 or len(block) == 0:
        return pd.Series([""]*len(block), index=block.index)
    tmp = block.copy()
    pid = pid_col if pid_col in tmp.columns else tmp.columns[0]
    tmp["__pid"] = tmp[pid].astype(str).fillna("")
    tmp = tmp.sort_values("__pid", kind="mergesort")
    out = []
    k = len(classes)
    for i, idx in enumerate(tmp.index):
        out.append(classes[i % k])
    return pd.Series(out, index=tmp.index).reindex(block.index)

PLACEHOLDER_ANY = re.compile(r"^\s*class(?:room)?\s*\d+\s*$", re.IGNORECASE)
ONLY_NUMERIC    = re.compile(r"^\s*\d+\s*$")

# =========================
# Main
# =========================
if uploaded_file:
    # 1) Detect header row
    wb_src = load_workbook(uploaded_file, data_only=True)
    ws_src = wb_src.active
    header_row = find_header_row(ws_src)
    if not header_row:
        st.error("Couldn't find 'ST: Participant PID' in the file.")
        st.stop()
    uploaded_file.seek(0)

    # 2) Load table
    df = pd.read_excel(uploaded_file, header=header_row - 1)
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]

    if "Participant PID" not in df.columns:
        st.error("The file is missing 'Participant PID'.")
        st.stop()

    # keep only most recent per PID (your original)
    df = (
        df.dropna(subset=["Participant PID"])
          .groupby("Participant PID", as_index=False)
          .agg(most_recent)
    )

    # 3) HARD-CODED class override (robust)
    center_col = find_center_column(df.columns)
    if center_col is None:
        st.warning("Could not find a Center column in Enrollment; skipping class override.")
    else:
        class_cols = find_all_class_columns(df.columns)
        # ensure any new class column exists
        for c in class_cols:
            if c not in df.columns:
                df[c] = ""

        # normalized center text for matching keys
        df["__CenterClean"] = (
            df[center_col].astype(str)
              .str.replace(r"^HCHSP\s*--\s*", "", regex=True)
              .str.strip()
        )

        debug_rows = []
        for e_center in df["__CenterClean"].dropna().unique():
            key = norm_center(e_center)
            classes = HARD_CODED_CLASSES_N.get(key, [])
            if not classes:
                continue

            mask_center = df["__CenterClean"] == e_center

            # choose rows to modify
            if FORCE_REPLACE_ALL_ROWS_FOR_MATCHING_CENTERS:
                need_mask = mask_center
            else:
                # replace only placeholders/numeric/empty
                any_mask = pd.Series(False, index=df.index)
                for col in class_cols:
                    vals = df[col].astype(str).fillna("")
                    any_mask |= (
                        vals.str.strip().eq("") |
                        vals.str.match(PLACEHOLDER_ANY) |
                        vals.str.match(ONLY_NUMERIC)
                    )
                need_mask = mask_center & any_mask

            block = df.loc[need_mask, ["Participant PID"]].copy()
            if block.empty:
                debug_rows.append({"Center": e_center, "Matched Key": key, "Rows changed": 0, "Columns": ", ".join(class_cols)})
                continue

            assigned = deterministic_assign(block, classes, pid_col="Participant PID")
            # write to ALL class-ish columns so we can't miss the one you're viewing
            for col in class_cols:
                df.loc[assigned.index, col] = assigned.values

            debug_rows.append({"Center": e_center, "Matched Key": key, "Rows changed": int(need_mask.sum()), "Columns": ", ".join(class_cols)})

        # drop temp
        df.drop(columns=["__CenterClean"], inplace=True, errors="ignore")

        # Debug panel
        if debug_rows:
            st.subheader("Class override — what changed")
            st.dataframe(pd.DataFrame(debug_rows))

            # quick visual for Farias/Guerra
            try:
                mask_fg = df[center_col].astype(str).str.contains("Farias|Guerra", case=False, na=False)
                preview_cols = ["Participant PID", center_col] + class_cols
                st.subheader("Preview: Farias & Guerra after override")
                st.dataframe(df.loc[mask_fg, preview_cols].head(40))
            except Exception:
                pass

    # 4) Write temp workbook
    title_text = "Enrollment Checklist 2025–2026"
    central_now = datetime.now(ZoneInfo("America/Chicago"))
    timestamp_text = central_now.strftime("Generated on %B %d, %Y at %I:%M %p %Z")

    temp_path = "Enrollment_Cleaned.xlsx"
    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        pd.DataFrame([[title_text]]).to_excel(writer, index=False, header=False, startrow=0)
        pd.DataFrame([[timestamp_text]]).to_excel(writer, index=False, header=False, startrow=1)
        df.to_excel(writer, index=False, startrow=3)

    # 5) Style with openpyxl (original)
    wb = load_workbook(temp_path)
    ws = wb.active

    filter_row = 4
    data_start = filter_row + 1
    data_end = ws.max_row
    max_col = ws.max_column

    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(max_col)}{data_end}"

    title_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    ts_fill    = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)

    tcell = ws.cell(row=1, column=1); tcell.value = title_text
    tcell.font = Font(size=14, bold=True); tcell.alignment = Alignment(horizontal="center", vertical="center"); tcell.fill = title_fill
    scell = ws.cell(row=2, column=1); scell.value = timestamp_text
    scell.font = Font(size=10, italic=True, color="555555"); scell.alignment = Alignment(horizontal="center", vertical="center"); scell.fill = ts_fill

    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[filter_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill

    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"), bottom=Side(style="thin"))
    red_font    = Font(color="FF0000", bold=True)

    # Identify immun & name cols
    immun_col = None
    name_col_idx = None
    headers = [ws.cell(row=filter_row, column=c).value for c in range(1, max_col + 1)]
    for idx, h in enumerate(headers, start=1):
        if isinstance(h, str):
            low = h.lower()
            if immun_col is None and "immun" in low:
                immun_col = idx
            if name_col_idx is None and "name" in low:
                name_col_idx = idx
    if name_col_idx is None:
        name_col_idx = 2

    cutoff_date  = datetime(2025, 5, 11)
    immun_cutoff = datetime(2024, 5, 11)

    for r in range(data_start, data_end + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val  = cell.value
            cell.border = thin_border

            if val in (None, "", "nan", "NaT"):
                cell.value = "X"; cell.font = red_font
                continue

            dt = coerce_to_dt(val)
            if dt:
                if c == immun_col and dt < immun_cutoff:
                    cell.value = dt; cell.number_format = "m/d/yy"; cell.font = red_font
                elif dt < cutoff_date:
                    cell.value = "X"; cell.font = red_font
                else:
                    cell.value = dt; cell.number_format = "m/d/yy"

    width_map = {1: 16, 2: 22}
    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = width_map.get(c, 14)

    # 6) Totals at the bottom (original)
    total_row = ws.max_row + 2
    ws.cell(row=total_row, column=1, value="Grand Total").font = Font(bold=True)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left", vertical="center")

    center = Alignment(horizontal="center", vertical="center")
    for c in range(1, max_col + 1):
        if c <= name_col_idx:
            continue
        valid_count = 0
        for r in range(data_start, data_end + 1):
            if ws.cell(row=r, column=c).value != "X":
                valid_count += 1
        cell = ws.cell(row=total_row, column=c, value=valid_count)
        cell.alignment = center; cell.font = Font(bold=True)
        cell.border = Border(top=Side(style="thin"))

    # 7) Save & download
    final_output = "Formatted_Enrollment_Checklist.xlsx"
    wb.save(final_output)
    with open(final_output, "rb") as f:
        st.download_button("📥 Download Formatted Excel", f, file_name=final_output)




