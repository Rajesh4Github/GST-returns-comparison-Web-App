import streamlit as st
from backend import execute_final_consolidation

st.title("GST Returns Consolidation & Audit")

uploaded_files = st.file_uploader(
    "Upload your GSTR returns (PDF/Excel)", 
    type=["pdf","xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    # Save uploaded files locally
    for f in uploaded_files:
        with open(f.name, "wb") as out:
            out.write(f.read())
    st.success(f"Uploaded {len(uploaded_files)} files successfully!")

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
