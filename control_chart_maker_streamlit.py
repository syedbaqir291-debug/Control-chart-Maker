import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from datetime import datetime

st.set_page_config(page_title="Control Chart Maker (QA OMAC Tools)", layout="wide")

# --- Header / Footer styling ---
st.markdown("<h1 style='text-align:left;'>Control chart maker (QA OMAC Tools)</h1>", unsafe_allow_html=True)
st.markdown("---")

st.sidebar.header("Upload & Setup")
uploaded_file = st.sidebar.file_uploader("Upload an Excel workbook (.xlsx)", type=["xlsx"]) 

if uploaded_file is not None:
    # read workbook and list sheets
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheets = xl.sheet_names
    except Exception as e:
        st.sidebar.error(f"Unable to read workbook: {e}")
        st.stop()

    sheet_name = st.sidebar.selectbox("Select sheet that contains data", options=sheets)
    if sheet_name:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=0)
        st.write("**Preview of selected sheet**")
        st.dataframe(df.head())

        # Select time column
        st.sidebar.subheader("Time / Grouping")
        time_col = st.sidebar.selectbox("Select the Time/Batch column (used for x-axis / grouping)", options=df.columns)
        time_slot = st.sidebar.selectbox("Select time slot / aggregation",
                                         options=["As-is (no aggregation)", "Monthly", "Quarterly", "Yearly", "Batch - as unique values"]) 

        # Select parameter columns (allow multiple)
        st.sidebar.subheader("Parameters for Control Chart")
        param_cols = st.sidebar.multiselect("Select one or more parameter columns",
                                            options=[c for c in df.columns if c != time_col])

        if len(param_cols) == 0:
            st.info("Select at least one parameter column from the sidebar to proceed.")
        else:
            st.sidebar.markdown("---")
            if st.sidebar.button("Process & Generate Outputs"):
                working = df.copy()

                # Attempt to parse time column to datetime if possible
                x = working[time_col].copy()
                parsed = pd.to_datetime(x, infer_datetime_format=True, errors='coerce')

                if time_slot != "As-is (no aggregation)":
                    if parsed.notna().all():
                        working['_time_dt'] = parsed
                        if time_slot == 'Monthly':
                            working['_period'] = working['_time_dt'].dt.to_period('M').dt.to_timestamp()
                        elif time_slot == 'Quarterly':
                            working['_period'] = working['_time_dt'].dt.to_period('Q').dt.to_timestamp()
                        elif time_slot == 'Yearly':
                            working['_period'] = working['_time_dt'].dt.to_period('Y').dt.to_timestamp()
                        else:
                            working['_period'] = working[time_col]
                    else:
                        working['_period'] = working[time_col]
                else:
                    working['_period'] = working[time_col]

                # Prepare output workbook in memory
                out_xlsx = BytesIO()
                writer = pd.ExcelWriter(out_xlsx, engine='openpyxl')

                # Build PPTX
                prs = Presentation()
                title_slide_layout = prs.slide_layouts[5]

                # For each selected parameter compute stats
                for col in param_cols:

                    # -----------------------------
                    # SAFE CLEANING OF NUMERIC DATA
                    # -----------------------------
                    working[col] = (
                        working[col]
                        .astype(str)
                        .str.replace(',', '', regex=False)
                        .str.strip()
                        .replace(['-', '–', '—', 'N/A', 'na', '', None], np.nan)
                    )

                    # Convert to numeric safely
                    working[col] = pd.to_numeric(working[col], errors='coerce')

                    col_values = working[col]

                    # Drop all NaN rows for statistical calculation
                    clean_vals = col_values.dropna()

                    if clean_vals.empty:
                        st.error(f"Column '{col}' has no numeric data.")
                        continue

                    # CL: mean
                    CL = clean_vals.mean()

                    # Moving Range (MR)
                    MR = clean_vals.diff().abs()
                    MRbar = MR[1:].mean()

                    # UCL / LCL calculation
                    d2 = 1.128
                    if not np.isnan(MRbar):
                        UCL = CL + 3 * (MRbar / d2)
                        LCL = CL - 3 * (MRbar / d2)
                        if LCL < 0:
                            LCL = 0
                    else:
                        UCL = CL
                        LCL = CL

                    # Add new columns to working df
                    working[f'{col}_CL'] = CL
                    working[f'{col}_MR'] = MR
                    working[f'{col}_MRbar'] = MRbar
                    working[f'{col}_UCL'] = UCL
                    working[f'{col}_LCL'] = LCL

                    # Grouping for chart
                    chart_df = working[[time_col, '_period', col]].copy()
                    if time_slot != 'As-is (no aggregation)':
                        chart_grouped = chart_df.groupby('_period')[col].mean().reset_index()
                        x_vals = chart_grouped['_period']
                        y_vals = chart_grouped[col]
                    else:
                        x_vals = chart_df[time_col]
                        y_vals = chart_df[col]

                    # Create chart
                    fig, ax = plt.subplots(figsize=(10,4))
                    ax.plot(x_vals, y_vals, marker='o', label=col)
                    ax.axhline(CL, linestyle='--', label='CL')
                    ax.axhline(UCL, color='r', linestyle='--', label='UCL')
                    ax.axhline(LCL, color='r', linestyle='--', label='LCL')
                    ax.set_title(f'Control Chart - {col}')
                    ax.set_xlabel(time_col)
                    ax.set_ylabel(col)
                    ax.legend()
                    plt.xticks(rotation=45)
                    plt.tight_layout()

                    img_bytes = BytesIO()
                    fig.savefig(img_bytes, format='png', dpi=150)
                    img_bytes.seek(0)
                    plt.close(fig)

                    slide = prs.slides.add_slide(title_slide_layout)
                    slide.shapes.add_picture(img_bytes, Inches(0.5), Inches(0.7), width=Inches(9))

                # Write updated df to Excel
                working.to_excel(writer, sheet_name=sheet_name, index=False)
                writer.save()
                out_xlsx.seek(0)

                # Save PPTX
                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)

                st.success('Processing complete. Download the outputs below:')
                st.download_button('Download updated Excel', data=out_xlsx,
                                   file_name='control_chart_output.xlsx',
                                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                st.download_button('Download PowerPoint (charts as images)', data=pptx_io,
                                   file_name='control_charts.pptx',
                                   mime='application/vnd.openxmlformats-officedocument.presentation.mspresentation')

                st.markdown("---")
                st.caption('Footer: OMAC Developer by SM Baqir 2025')

else:
    st.info('Upload an Excel workbook to get started.')
