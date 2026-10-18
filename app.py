import os
import streamlit as st
import pandas as pd
import altair as alt
from backend import execute_unified_consolidation

st.set_page_config(
    page_title="GST Returns Comparison & Audit App",
    page_icon="📊",
    layout="wide"
)

st.title("📊 GST Returns Comparison & Multi-Sheet Audit App")
st.markdown("Upload your monthly and annual GST return PDFs, run reconciliation models, and download auditor-ready Excel workbooks.")

# Tabs
tab_upload, tab_summary, tab_analytics, tab_download = st.tabs([
    "📁 Upload Documents",
    "📊 Summary Dashboard",
    "📑 Detailed Audit Schedules",
    "📥 Download Report"
])

# ---------------- Upload Tab ----------------
with tab_upload:
    st.subheader("Upload Monthly & Annual GST Returns")
    uploaded_files = st.file_uploader(
        "Upload GSTR PDFs (1, 2B, 3B, 9)",
        type=["pdf"],
        accept_multiple_files=True
    )

    if uploaded_files:
        os.makedirs("data", exist_ok=True)
        for uploaded_file in uploaded_files:
            file_path = os.path.join("data", uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        st.success(f"Uploaded {len(uploaded_files)} file(s)!")

    if st.button("🚀 Run Consolidation", type="primary"):
        output_file = "GST_Consolidated_Audit_Report.xlsx"
        with st.spinner("Processing files and building audit schedules..."):
            try:
                report_path = execute_unified_consolidation(output_path=output_file)
                st.session_state["report_ready"] = True
                st.session_state["output_file"] = report_path
                st.success("Consolidation complete!")
            except Exception as e:
                st.error(f"Error during consolidation: {e}")

# ---------------- Summary Dashboard Tab ----------------
with tab_summary:
    st.subheader("📊 Consolidated Summary Dashboard")
    output_file = "GST_Consolidated_Audit_Report.xlsx"
    if os.path.exists(output_file):
        xls = pd.ExcelFile(output_file)
        if "Summary Dashboard" in xls.sheet_names:
            df_summary = pd.read_excel(xls, sheet_name="Summary Dashboard")
            st.dataframe(df_summary, use_container_width=True)

            # KPIs
            cols = st.columns(len(df_summary))
            for i, row in df_summary.iterrows():
                cols[i].metric(label=row["Metric"], value=row["Value"])

            # Charts: Monthly deviations from ITC Supplier Recon
            if "1. ITC Supplier Recon" in xls.sheet_names:
                df_recon = pd.read_excel(xls, sheet_name="1. ITC Supplier Recon")
                chart = alt.Chart(df_recon).mark_bar(color="orange").encode(
                    x=alt.X("Period", sort=None),
                    y="Deviation",
                    tooltip=["Period", "Deviation", "Risk Status"]
                ).properties(title="Excess ITC Availment vs 2B")
                st.altair_chart(chart, use_container_width=True)

            # Charts: ITC composition pie chart
            if "GSTR-2B Summary (12M)" in xls.sheet_names:
                df_itc = pd.read_excel(xls, sheet_name="GSTR-2B Summary (12M)")
                totals = df_itc.groupby("Details / Field").sum(numeric_only=True).reset_index()
                pie = alt.Chart(totals).mark_arc().encode(
                    theta="FY TOTAL",
                    color="Details / Field",
                    tooltip=["Details / Field", "FY TOTAL"]
                ).properties(title="ITC Composition (2B)")
                st.altair_chart(pie, use_container_width=True)

# ---------------- Detailed Analytics Tab ----------------
with tab_analytics:
    st.subheader("🔍 Detailed Audit Schedules")
    output_file = "GST_Consolidated_Audit_Report.xlsx"
    if os.path.exists(output_file):
        xls = pd.ExcelFile(output_file)
        selected_schedule = st.selectbox("Select Audit Schedule:", xls.sheet_names)
        if selected_schedule:
            df_display = pd.read_excel(xls, sheet_name=selected_schedule)
            st.markdown(f"### Schedule View: `{selected_schedule}`")
            st.dataframe(df_display, use_container_width=True)

# ---------------- Download Tab ----------------
with tab_download:
    st.subheader("📥 Download Auditor-Ready Excel Report")
    output_file = "GST_Consolidated_Audit_Report.xlsx"
    if os.path.exists(output_file):
        with open(output_file, "rb") as f:
            st.download_button(
                label="⬇️ Download Full Audit Report (.xlsx)",
                data=f,
                file_name="GST_Consolidated_Audit_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
    else:
        st.error("Report file not found. Please run consolidation first.")
