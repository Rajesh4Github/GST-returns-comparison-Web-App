import os
import streamlit as st
import pandas as pd
from backend import execute_final_consolidation as execute_unified_consolidation

st.set_page_config(
    page_title="GST Returns Comparison & Audit App",
    page_icon="📊",
    layout="wide"
)

st.title("📊 GST Returns Comparison & Multi-Sheet Audit App")
st.markdown("Upload your monthly and annual GST return PDFs or schedules, run all reconciliation models, and download auditor-ready Excel workbooks.")

tab_upload, tab_analytics, tab_download = st.tabs([
    "📁 Document Upload Panel", 
    "📊 Audit Algorithms & Analytics", 
    "📥 Download Auditor-Ready Excel Working Papers"
])

with tab_upload:
    st.subheader("Upload Monthly & Annual GST Returns")
    uploaded_files = st.file_uploader(
        "Upload GSTR PDFs, GSTR-2B/3B Excel files, or GSTR-9 working papers", 
        type=["pdf", "xlsx"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        os.makedirs("data", exist_ok=True)
        for uploaded_file in uploaded_files:
            file_path = os.path.join("data", uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        st.success(f"Successfully uploaded {len(uploaded_files)} file(s)!")

    if st.button("🚀 Execute Master Audit Consolidation", type="primary"):
        output_file = "GST_Consolidated_Audit_Report.xlsx"
        with st.spinner("Processing files and building audit schedules..."):
            try:
                report_path = execute_unified_consolidation(output_path=output_file)
                st.session_state["report_ready"] = True
                st.session_state["output_file"] = report_path
                st.success("Consolidation complete!")
            except Exception as e:
                st.error(f"Error during consolidation: {e}")

with tab_analytics:
    st.subheader("🔍 Audit & Compliance Reconciliation Schedules")
    
    output_file = "GST_Consolidated_Audit_Report.xlsx"
    if not os.path.exists(output_file):
        execute_unified_consolidation(output_path=output_file)
        st.session_state["report_ready"] = True

    if os.path.exists(output_file):
        xls = pd.ExcelFile(output_file)
        selected_schedule = st.selectbox("Select Audit Schedule to Inspect:", xls.sheet_names)
        
        if selected_schedule:
            df_display = pd.read_excel(xls, sheet_name=selected_schedule)
            st.markdown(f"### Schedule View: `{selected_schedule}`")
            st.dataframe(df_display, use_container_width=True)

with tab_download:
    st.subheader("📥 Download Auditor-Ready Excel Working Papers")
    output_file = "GST_Consolidated_Audit_Report.xlsx"
    
    if not os.path.exists(output_file):
        execute_unified_consolidation(output_path=output_file)
        
    if os.path.exists(output_file):
        with open(output_file, "rb") as f:
            st.download_button(
                label="⬇️ Download Complete Multi-Sheet Audit Report (.xlsx)",
                data=f,
                file_name="GST_Consolidated_Audit_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
    else:
        st.error("Report file not found. Please run the consolidation pipeline first.")
