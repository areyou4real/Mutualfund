import streamlit as st
import pandas as pd
from master_pipeline import run_master_pipeline
from io import BytesIO

st.set_page_config(page_title="Mutual Fund Summary Tool", layout="centered")

# ---- HEADER ----
st.markdown("""
    <style>
        .title {
            font-size: 40px;
            font-weight: 700;
            color: #3C3C3C;
        }
        .footer {
            position: fixed;
            bottom: 10px;
            font-size: 13px;
            color: gray;
            width: 100%;
            text-align: center;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="title">📊 Mutual Fund Summary Generator</p>', unsafe_allow_html=True)
st.write("Upload one or more Excel files below. The app will process each fund and compile the results into a single Excel file.")

# ---- FILE UPLOAD ----
uploaded_files = st.file_uploader(
    "📁 Upload your Excel files", accept_multiple_files=True, type=["xlsx"]
)

if uploaded_files:
    with st.spinner("Processing uploaded files..."):
        byte_data = {file.name: file.read() for file in uploaded_files}
        results = run_master_pipeline(byte_data)

    st.success("✅ All files processed!")

    # ---- PREVIEW RESULTS ----
    for fund_name, result in results.items():
        with st.expander(f"📄 {fund_name.upper()} Summary", expanded=False):
            if isinstance(result, pd.DataFrame):
                st.dataframe(result, use_container_width=True)
            else:
                st.error(result)

    # ---- EXPORT BUTTON ----
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
    wrote_at_least_one_sheet = False
    for fund_name, df in results.items():
        if isinstance(df, pd.DataFrame) and not df.empty:
            df.to_excel(writer, index=False, sheet_name=fund_name[:31])
            wrote_at_least_one_sheet = True

    if not wrote_at_least_one_sheet:
        st.error("❌ No valid sheets to export. Please check uploaded files.")
        st.stop()
    output.seek(0)

    st.download_button(
        label="📥 Download Combined Excel File",
        data=output,
        file_name="all_funds_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---- FOOTER ----
st.markdown('<div class="footer">Built with ❤️ using Streamlit | For internal use only</div>', unsafe_allow_html=True)
