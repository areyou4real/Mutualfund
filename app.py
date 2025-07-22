
import streamlit as st
import pandas as pd
from master_pipeline import run_master_pipeline

st.set_page_config(page_title="Mutual Fund Summary Tool", layout="centered")
st.title("🧾 Mutual Fund Summary Processor")

st.markdown("Upload one or more Excel files below. The tool will detect and process supported fund formats and return a single Excel file with each fund on its own sheet.")

uploaded_files = st.file_uploader("📤 Upload your fund Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    uploaded = {file.name.rsplit(".", 1)[0].lower().replace(" ", ""): file.read() for file in uploaded_files}

    if st.button("🚀 Run Processing"):
        with st.spinner("Processing all uploaded funds..."):
            try:
                results = run_master_pipeline(uploaded)
                if results:
                    output_path = "/tmp/all_funds_summary.xlsx"
                    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                        for sheet_name, df in results.items():
                            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

                    with open(output_path, "rb") as f:
                        st.success("✅ All funds processed successfully!")
                        st.download_button("📥 Download Summary Excel", f, file_name="all_funds_summary.xlsx")
                else:
                    st.warning("⚠️ No valid sheets generated. Please check your files.")
            except Exception as e:
                st.error(f"❌ Error during processing: {e}")
