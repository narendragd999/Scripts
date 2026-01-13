import streamlit as st
import pandas as pd
import warnings
from io import BytesIO
import re

warnings.filterwarnings("ignore", category=UserWarning)

# ================= CONFIG =================
ROWS_TO_DELETE = 10
TARGET_SHEET = "Draft Roll Control Chart"
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
Panchayat name will be **auto-detected from filename** and added as a new column.
""")

uploaded_files = st.file_uploader(
    "📂 Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

def extract_panchayat_name(filename):
    """
    Extracts text before '-Format-A1'
    Example:
    Aamalva-Format-A1_(Name of PRI)_SEC_Rajasthan.xlsx -> Aamalva
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

                try:
                    df = pd.read_excel(
                        file,
                        sheet_name=TARGET_SHEET,
                        header=None,
                        engine="openpyxl"
                    )
                except Exception:
                    st.warning(f"⚠ Sheet not found: {file.name}")
                    continue

                # Remove top 10 rows
                df = df.iloc[ROWS_TO_DELETE:]

                # Keep only columns B–F
                df = df.iloc[:, 1:6]
                df.columns = FIXED_COLUMNS

                # Drop empty rows
                df = df.dropna(how="all", subset=FIXED_COLUMNS)

                if df.empty:
                    st.warning(f"⚠ No valid data in {file.name}")
                    continue

                # Add Panchayat column
                df.insert(0, NEW_COLUMN, panchayat)

                rows = len(df)
                total_rows += rows
                merged_rows.append(df)

                st.write(f"✔ Rows added: {rows}")

        if not merged_rows:
            st.error("❌ No valid data found in uploaded files")
            st.stop()

        final_df = pd.concat(merged_rows, ignore_index=True)

        # Write output to memory
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
