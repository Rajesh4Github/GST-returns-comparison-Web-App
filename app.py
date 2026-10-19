import streamlit as st
from backend import execute_final_consolidation

st.title("GST Returns Consolidation & Audit")

st.write("Generate a unified GST Consolidated Audit Report from uploaded returns.")

if st.button("Run Consolidation"):
    try:
        output_file = execute_final_consolidation()
        st.success(f"✅ Report generated successfully at: {output_file}")
        st.download_button(
            label="Download Consolidated Report",
            data=open(output_file, "rb").read(),
            file_name="GST_Consolidated_Audit_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Error during consolidation: {e}")
