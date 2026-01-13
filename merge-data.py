import streamlit as st
import pandas as pd
import warnings
from io import BytesIO
import re

warnings.filterwarnings("ignore", category=UserWarning)

# ================= CONFIG =================
ROWS_TO_DELETE = 10
FIXED_COLUMNS = ["B", "C", "D", "E", "F"]
NEW_COLUMN = "Panchayat Name"
# =========================================

st.set_page_config(
    page_title="Draft Roll Control Chart Merger",
    layout="centered"
)

st.title("📊 Draft Roll Control Chart – Excel Merger")
st.markdown("""
Upload **multiple A1 Excel files**.  
✔ Draft Roll sheet is auto-detected  
✔ Panchayat name is auto-added from filename  
""")

uploaded_files = st.file_uploader(
    "📂 Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

def extract_panchayat_name(filename):
    """
    Extracts Panchayat name from filename
    Example:
    Aamalva-Format-A1_(Name of PRI)_SEC_Rajasthan.xlsx → Aamalva
    """
    match = re.match(r"(.*?)-Format-A1", filename)
    return match.group(1).strip() if match else "UNKNOWN"

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded")

    if st.button("🚀 Clean & Merge", use_container_width=True):

        merged_rows = []
        total_rows = 0

        with st.spinner("Processing files..."):

            for file in uploaded_files:
                panchayat = extract_panchayat_name(file.name)
                st.write(f"➡ Processing **{file.name}** → **{panchayat}**")

                # -------- AUTO DETECT SHEET --------
                try:
                    xls = pd.ExcelFile(file, engine="openpyxl")
                except Exception as e:
                    st.warning(f"⚠ Cannot read file: {file.name}")
                    continue

                target_sheet = None
                for sheet in xls.sheet_names:
                    sheet_clean = sheet.lower().replace("\u00a0", " ").strip()
                    if "draft" in sheet_clean and "roll" in sheet_clean:
                        target_sheet = sheet
                        break

                if not target_sheet:
                    st.warning(f"⚠ Draft Roll sheet not found: {file.name}")
                    continue

                # -------- READ SHEET --------
                df = pd.read_excel(
                    file,
                    sheet_name=target_sheet,
                    header=None,
                    engine="openpyxl"
                )

                # -------- CLEAN DATA --------
                df = df.iloc[ROWS_TO_DELETE:]        # remove top rows
                df = df.iloc[:, 1:6]                 # columns B–F
                df.columns = FIXED_COLUMNS

                df = df.dropna(how="all", subset=FIXED_COLUMNS)

                if df.empty:
                    st.warning(f"⚠ No valid rows in {file.name}")
                    continue

                # -------- ADD PANCHAYAT --------
                df.insert(0, NEW_COLUMN, panchayat)

                rows = len(df)
                total_rows += rows
                merged_rows.append(df)

                st.write(f"✔ Rows added: {rows}")

        if not merged_rows:
            st.error("❌ No valid data found in uploaded files")
            st.stop()

        final_df = pd.concat(merged_rows, ignore_index=True)

        # -------- EXPORT --------
        output = BytesIO()
        final_df.to_excel(output, index=False)
        output.seek(0)

        st.success("✅ Merge completed successfully")
        st.metric("📊 Total Rows Merged", total_rows)
        st.metric("📁 Files Processed", len(uploaded_files))

        st.download_button(
            label="⬇ Download Merged Excel",
            data=output,
            file_name="MERGED_DRAFT_ROLL_CONTROL_CHART.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

else:
    st.info("👆 Upload Excel files to begin")
