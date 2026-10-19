import streamlit as st
import os
from backend import execute_final_consolidation

st.title("GST Returns Consolidation & Audit")

uploaded_files = st.file_uploader(
    "Upload your GSTR returns (PDF/Excel)", 
    type=["pdf","xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    os.makedirs("uploads", exist_ok=True)
    for f in uploaded_files:
        with open(os.path.join("uploads", f.name), "wb") as out:
            out.write(f.read())
    st.success(f"Uploaded {len(uploaded_files)} files successfully!")

if st.button("Run Consolidation"):
    output_file = execute_final_consolidation(output_path="./GST_Consolidated_Audit_Report.xlsx")
    st.success(f"Report generated: {output_file}")
    st.download_button(
        label="Download Consolidated Report",
        data=open(output_file, "rb").read(),
        file_name="GST_Consolidated_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
