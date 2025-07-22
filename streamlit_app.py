import streamlit as st
import pandas as pd
from io import BytesIO
import re
from master_pipeline import run_master_pipeline

st.title("🧾 Mutual Fund Summary Generator")
st.markdown("Upload one or more mutual fund Excel files. The app will detect the AMC and return a cleaned summary.")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    file_dict = {file.name: file.read() for file in uploaded_files}
    results = run_master_pipeline(file_dict)

    valid_results = {k: v for k, v in results.items() if isinstance(v, pd.DataFrame)}
    error_results = {k: v for k, v in results.items() if not isinstance(v, pd.DataFrame)}

    if error_results:
        st.subheader("❌ Errors")
        for name, error in error_results.items():
            st.error(f"{name}: {error}")

    if valid_results:
        st.subheader("✅ Fund Summaries")
        for name, df in valid_results.items():
            with st.expander(f"{name.title()} Summary"):
                st.dataframe(df)

        # Download button
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for name, df in valid_results.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)
        output.seek(0)
        st.download_button("📥 Download Combined Excel", output, file_name="all_funds_summary.xlsx")
    else:
        st.warning("No valid dataframes to display or download.")
