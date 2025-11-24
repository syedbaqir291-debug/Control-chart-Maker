# runcharts_streamlit_revamped.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

st.set_page_config(page_title="Premium Control Chart Tool", layout="wide")
st.title("📊 OMAC Control Chart Tool")
st.markdown("""
This tool generates control charts from your data.
- Upload Excel/CSV
- Select column(s)
- Automatically considers all numeric data
""")

# -----------------------------
# Upload Data
# -----------------------------
uploaded_file = st.file_uploader("Upload your Excel/CSV file", type=["xlsx", "csv"])
if uploaded_file:
    # Read Excel or CSV
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file, header=0)
    else:
        df = pd.read_excel(uploaded_file, header=0)

    st.success("File uploaded successfully!")
    
    # Select numeric columns only
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    
    if not numeric_cols:
        st.error("No numeric columns found!")
    else:
        selected_col = st.selectbox("Select numeric column for control chart", numeric_cols)

        # Extract numeric data, flatten if needed
        data = df[selected_col].dropna().values

        # -----------------------------
        # Control Chart Calculations
        # -----------------------------
        mean_val = np.mean(data)
        sigma_val = np.std(data, ddof=1)
        ucl = mean_val + 3*sigma_val
        lcl = mean_val - 3*sigma_val

        # Plot Control Chart
        fig, ax = plt.subplots(figsize=(12, 5))
        ax.plot(data, marker='o', linestyle='-', color='blue', label='Data')
        ax.axhline(mean_val, color='green', linestyle='--', label='Mean')
        ax.axhline(ucl, color='red', linestyle='--', label='UCL (Mean + 3σ)')
        ax.axhline(lcl, color='red', linestyle='--', label='LCL (Mean - 3σ)')

        # Highlight Out-of-Control Points
        out_of_control = np.where((data > ucl) | (data < lcl))[0]
        ax.scatter(out_of_control, data[out_of_control], color='red', s=100, label='Out of Control')

        ax.set_title(f"Control Chart for {selected_col}")
        ax.set_xlabel("Sample Number / Sequence")
        ax.set_ylabel("Value")
        ax.legend()
        ax.grid(True)
        st.pyplot(fig)

        # Summary
        st.markdown("### Control Chart Summary")
        st.write(f"Mean: {mean_val:.2f}")
        st.write(f"Standard Deviation: {sigma_val:.2f}")
        st.write(f"UCL: {ucl:.2f}, LCL: {lcl:.2f}")
        st.write(f"Number of points out of control: {len(out_of_control)}")
