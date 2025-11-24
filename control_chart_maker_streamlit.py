import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Control Chart Excel Updater (QA OMAC Tools)", layout="wide")
st.markdown("<h1 style='text-align:left;'>Control Chart Excel Updater (QA OMAC Tools)</h1>", unsafe_allow_html=True)
st.markdown("---")

# --- Upload Excel ---
st.sidebar.header("Upload & Setup")
uploaded_file = st.sidebar.file_uploader("Upload an Excel workbook (.xlsx)", type=["xlsx"], key="file_upload")

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    sheet_name = st.sidebar.selectbox("Select the sheet with your data", xl.sheet_names)

    # Ask user which row is header
    header_row = st.sidebar.number_input("Header row (0 = first row)", min_value=0, value=0, step=1)

    if sheet_name:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=header_row)
        st.write("**Preview of sheet:**")
        st.dataframe(df.head())

        # Ask user which column is Time/Batch
        time_col = st.sidebar.selectbox("Select Time / Batch column", df.columns)

        # Automatically take all other columns as numeric parameters
        param_cols = [c for c in df.columns if c != time_col]
        st.sidebar.write("Parameters detected:", param_cols)

        if param_cols and st.sidebar.button("Process & Update Excel"):
            working = df.copy()

            for col in param_cols:
                # Ensure numeric (skip cleaning since data is already numeric)
                col_values = pd.to_numeric(working[col], errors='coerce')

                # Compute Control Chart metrics
                CL = col_values.mean(skipna=True)
                MR = col_values.diff().abs()
                MRbar = MR[1:].mean() if len(MR) > 1 else 0
                d2 = 1.128
                UCL = CL + 3 * (MRbar / d2)
                LCL = max(0, CL - 3 * (MRbar / d2))

                # Add new columns
                working[f'{col}_CL'] = CL
                working[f'{col}_MR'] = MR
                working[f'{col}_MRbar'] = MRbar
                working[f'{col}_UCL'] = UCL
                working[f'{col}_LCL'] = LCL

            # Save to Excel in memory
            out_xlsx = BytesIO()
            with pd.ExcelWriter(out_xlsx, engine='openpyxl') as writer:
                working.to_excel(writer, sheet_name=sheet_name, index=False)
            out_xlsx.seek(0)

            st.success("Excel updated with control chart metrics!")
            st.download_button(
                "Download updated Excel",
                data=out_xlsx,
                file_name="control_chart_updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Upload an Excel workbook to get started.")
