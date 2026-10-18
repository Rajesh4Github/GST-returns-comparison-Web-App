import os
import streamlit as st
import pandas as pd
from backend import execute_unified_consolidation

st.set_page_config(
    page_title="GST Returns Comparison & Audit App",
    page_icon="📊",
    layout="wide"
)

st.title("📊 GST Returns Comparison & Multi-Sheet Audit App")
st.markdown("Upload your monthly and annual GST return PDFs or schedules, run all 10 data analytical audit models, and download auditor-ready Excel workbooks.")

# Navigation / Sections using Tabs
tab_upload, tab_analytics, tab_download = st.tabs([
    "📁 Document Upload Panel", 
    "📊 Audit Algorithms & Analytics (10 Modules)", 
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
        st.success(f"Successfully uploaded {len(uploaded_files)} file(s) to working environment!")

    if st.button("🚀 Execute All 10 Audit & Reconciliation Algorithms", type="primary"):
        output_file = "GST_Consolidated_Audit_Report.xlsx"
        with st.spinner("Running all 10 analytical reconciliation models..."):
            try:
                report_path = execute_unified_consolidation(output_path=output_file)
                st.session_state["report_ready"] = True
                st.session_state["output_file"] = report_path
                st.success("All 10 audit algorithms executed successfully!")
            except Exception as e:
                st.error(f"Error during execution: {e}")

with tab_analytics:
    st.subheader("🔍 10 Audit & Compliance Reconciliation Schedules")
    
    output_file = "GST_Consolidated_Audit_Report.xlsx"
    if not os.path.exists(output_file) and not st.session_state.get("report_ready", False):
        st.info("⚠️ Please upload your files and run the algorithms in the **Document Upload Panel** first, or view sample audit results below:")
        # Generate sample immediately if not run yet
        execute_unified_consolidation(output_path=output_file)
        st.session_state["report_ready"] = True

    if os.path.exists(output_file):
        xls = pd.ExcelFile(output_file)
        selected_schedule = st.selectbox("Select Audit Algorithm / Schedule to Inspect:", xls.sheet_names)
        
        if selected_schedule:
            df_display = pd.read_excel(xls, sheet_name=selected_schedule)
            st.markdown(f"### Schedule View: `{selected_schedule}`")
            st.dataframe(df_display, use_container_width=True)
            
            # Summary metrics or status
            if "Risk Status" in df_display.columns:
                risks = df_display[df_display["Risk Status"].str.contains("Risk|Action", case=False, na=False)]
                if len(risks) > 0:
                    st.warning(f"⚠️ Found {len(risks)} period(s) requiring attention in this schedule.")
                else:
                    st.success("✅ All periods appear clean and satisfactory.")
            elif "Status" in df_display.columns:
                issues = df_display[df_display["Status"].str.contains("Action|Risk", case=False, na=False)]
                if len(issues) > 0:
                    st.warning(f"⚠️ Found {len(issues)} item(s) flagged for review.")
                else:
                    st.success("✅ Audit checks passed successfully.")

with tab_download:
    st.subheader("📥 Download Auditor-Ready Excel Working Papers")
    output_file = "GST_Consolidated_Audit_Report.xlsx"
    
    # Ensure file exists
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
        st.markdown("---")
        st.info("💡 **Tip:** This Excel report contains all 10 segregated working papers, ready to be handed directly to statutory or tax auditors.")
    else:
        st.error("Report file not found. Please run the consolidation pipeline first.")
