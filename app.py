import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="Enrollment Report", layout="wide")

st.title("HCHSP Enrollment Report")

# ---- File uploaders ----
uploaded_vf = st.file_uploader("Upload VF_Average_Funded_Enrollment_Level.xlsx", type=["xlsx"])
uploaded_detail = st.file_uploader("Upload Detail Report.xlsx", type=["xlsx"])

if uploaded_vf and uploaded_detail:
    try:
        # --- Load both files ---
        vf_df = pd.read_excel(uploaded_vf)
        detail_df = pd.read_excel(uploaded_detail)

        # --- Keep only rows that look like class names ---
        vf_classes = vf_df[vf_df["Class"].astype(str).str.contains(r"Class \d+", case=False, na=False)].copy()

        # --- Ensure numeric ---
        for col in ["Funded Enrollment", "Enrolled"]:
            if col in vf_classes.columns:
                vf_classes[col] = pd.to_numeric(vf_classes[col], errors="coerce").fillna(0)

        # --- Add Percentage Column ---
        vf_classes["% Enrolled of Funded"] = np.where(
            vf_classes["Funded Enrollment"] > 0,
            (vf_classes["Enrolled"] / vf_classes["Funded Enrollment"]) * 100,
            np.nan
        )

        # --- Add Waitlist + Lacking if exist in VF ---
        if "Waitlist" in vf_df.columns:
            vf_classes["Waitlist"] = vf_df.loc[vf_classes.index, "Waitlist"]
        else:
            vf_classes["Waitlist"] = np.nan

        if "Lacking" in vf_df.columns:
            vf_classes["Lacking"] = vf_df.loc[vf_classes.index, "Lacking"]
        else:
            vf_classes["Lacking"] = np.nan

        # --- Reorder columns ---
        report_df = vf_classes[["Class", "Funded Enrollment", "Enrolled", "% Enrolled of Funded", "Waitlist", "Lacking"]]

        # --- Style function ---
        def highlight_percent(val):
            try:
                if float(val) > 100:
                    return "color: blue"
            except:
                return ""
            return ""

        # --- Display styled table ---
        st.dataframe(
            report_df.style.format({
                "Funded Enrollment": "{:,.0f}",
                "Enrolled": "{:,.0f}",
                "% Enrolled of Funded": "{:.1f}%",
                "Waitlist": "{:,.0f}",
                "Lacking": "{:,.0f}",
            }).applymap(highlight_percent, subset=["% Enrolled of Funded"])
        )

        # --- Downloadable Excel ---
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            report_df.to_excel(writer, index=False, sheet_name="Enrollment Report")
        st.download_button("Download Excel Report", output.getvalue(), file_name="Enrollment_Report.xlsx")

    except Exception as e:
        st.error(f"Processing error: {e}")

