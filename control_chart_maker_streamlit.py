import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- Streamlit Page Setup ---
st.set_page_config(page_title="Control Chart Excel Updater (QA OMAC Tools)", layout="wide")
st.markdown("<h1 style='text-align:left;'>Control Chart Excel Updater (QA OMAC Tools)</h1>", unsafe_allow_html=True)
st.markdown("---")

# --- Sidebar: Upload & Selection ---
st.sidebar.header("Upload & Setup")
uploaded_file = st.sidebar.file_uploader(
    "Upload an Excel workbook (.xlsx)", type=["xlsx"], key="file_upload"
)

if uploaded_file is not None:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheets = xl.sheet_names
    except Exception as e:
        st.sidebar.error(f"Unable to read workbook: {e}")
        st.stop()

    sheet_name = st.sidebar.selectbox(
        "Select the sheet containing your data", options=sheets, key="sheet_select"
    )

    if sheet_name:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=0)
        st.write("**Preview of selected sheet:**")
        st.dataframe(df.head())

        # --- Time/Batch column ---
        time_col = st.sidebar.selectbox(
            "Select Time / Batch column", options=df.columns, key="time_col"
        )

        # --- Parameter columns ---
        param_cols = st.sidebar.multiselect(
            "Select one or more parameter columns",
            options=[c for c in df.columns if c != time_col],
            default=[c for c in df.columns if c != time_col],
            key="param_cols"
        )

        if len(param_cols) == 0:
            st.info("Please select at least one parameter column to proceed.")
        else:
            if st.sidebar.button("Process & Update Excel", key="process_button"):
                working = df.copy()

                # --- Clean numeric columns ---
                for col in param_cols:
                    working[col] = (
                        working[col]
                        .astype(str)
                        .str.replace(',', '', regex=False)
                        .str.strip()
                        .replace(['-', '–', '—', 'N/A', 'na', '', None], np.nan)
                    )
                    working[col] = pd.to_numeric(working[col], errors='coerce')

                    # --- Compute Control Chart Metrics ---
                    col_values = working[col]
                    CL = col_values.mean(skipna=True)
                    MR = col_values.diff().abs()
                    MRbar = MR[1:].mean() if len(MR) > 1 else 0
                    d2 = 1.128
                    UCL = CL + 3 * (MRbar / d2)
                    LCL = max(0, CL - 3 * (MRbar / d2))

                    # --- Add columns to dataframe ---
                    working[f'{col}_CL'] = CL
                    working[f'{col}_MR'] = MR
                    working[f'{col}_MRbar'] = MRbar
                    working[f'{col}_UCL'] = UCL
                    working[f'{col}_LCL'] = LCL

                # --- Write updated Excel to memory ---
                out_xlsx = BytesIO()
                with pd.ExcelWriter(out_xlsx, engine='openpyxl') as writer:
                    working.to_excel(writer, sheet_name=sheet_name, index=False)
                out_xlsx.seek(0)

                st.success("Excel updated with Control Chart metrics!")
                st.download_button(
                    "Download updated Excel",
                    data=out_xlsx,
                    file_name="control_chart_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.info("Upload an Excel workbook to get started.")
