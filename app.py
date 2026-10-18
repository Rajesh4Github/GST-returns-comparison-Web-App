# app.py
import os
import streamlit as st
import pandas as pd
from pathlib import Path
from pypdf import PdfReader

st.set_page_config(
    page_title="GST Detailed Audit & Analytics Portal",
    page_icon="🔍",
    layout="wide"
)

st.title("🔍 GST Detailed Audit — Automated Desk Review & Reconciliation Portal")
st.markdown("Upload your monthly and annual GST return PDFs, run data analytical models, and download auditor-ready Excel workbooks.")

# Sidebar Settings
st.sidebar.header("⚙️ Audit Parameters")
gstin_input = st.sidebar.text_input("Taxpayer GSTIN", value="32AAAJU0328F1ZJ")
audit_year = st.sidebar.selectbox("Audit Year", ["2024-25", "2023-24"])
commissionerate = st.sidebar.text_input("Commissionerate", value="CGST Commissionerate Thiruvananthapuram")

tab_upload, tab_analytics, tab_downloads = st.tabs([
    "📁 Document Upload Panel", 
    "📊 Audit Algorithms & Analytics", 
    "📥 Download Consolidated Excel Working Papers"
])

workspace_dir = Path("./app_workspace")
returns_staging = workspace_dir / "returns_folder"
workspace_dir.mkdir(exist_ok=True)
returns_staging.mkdir(exist_ok=True)

with tab_upload:
    st.subheader("Document Upload Panel (Returns & Statements)")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 1. Monthly & Annual Returns (PDFs)")
        uploaded_3b = st.file_uploader("GSTR-3B Returns (Up to 12 files)", type=["pdf"], accept_multiple_files=True)
        uploaded_gstr1 = st.file_uploader("GSTR-1 Returns (Up to 12 files)", type=["pdf"], accept_multiple_files=True)
        uploaded_gstr2b = st.file_uploader("GSTR-2B Statements (Up to 12 files)", type=["pdf"], accept_multiple_files=True)
        uploaded_gstr9 = st.file_uploader("GSTR-9 Annual Returns", type=["pdf"], accept_multiple_files=True)

    def save_uploaded_files(files, prefix):
        if not files:
            return
        for file in files:
            dest = returns_staging / f"{prefix}_{file.name}"
            with open(dest, "wb") as f:
                f.write(file.getbuffer())

    if st.button("🚀 Process Documents & Initialize Staging", type="primary"):
        with st.spinner("Saving uploaded files to workspace..."):
            save_uploaded_files(uploaded_3b, "GSTR3B")
            save_uploaded_files(uploaded_gstr1, "GSTR1")
            save_uploaded_files(uploaded_gstr2b, "GSTR2B")
            save_uploaded_files(uploaded_gstr9, "GSTR9")
            st.success("Files successfully staged in app workspace! Ready for audit analytics.")

with tab_analytics:
    st.subheader("📊 Audit Analytics & Data Reconciliation Models")
    
    st.markdown("### ITC Analytical Reconciliation Model (GSTR-3B vs GSTR-2B / GSTR-9)")
    col_a, col_b = st.columns(2)
    with col_a:
        v_3b_4a5 = st.number_input("GSTR-3B 4A(5) ITC Availed", value=18752270.59)
        v_3b_4d1 = st.number_input("GSTR-3B 4D(1) Reclaim", value=18752270.59)
        v_g9_6h  = st.number_input("GSTR-9 6H Reclaimed ITC", value=15165555.85)
        v_8c_curr = st.number_input("GSTR-9 8C Current Year", value=0.0)
    with col_b:
        v_8c_prev = st.number_input("GSTR-9 8C Previous Year", value=0.0)
        v_2b_3ai = st.number_input("GSTR-2B Table 3A(I) / 3B", value=8047583.23)
        v_2b_3bi = st.number_input("GSTR-2B Table 3B(I) Credit Notes", value=0.0)
        v_2b_4b  = st.number_input("GSTR-2B Table 4B Not Available", value=0.0)
        v_g9_t12 = st.number_input("GSTR-9 Table 12", value=0.0)

    if st.button("Run ITC Analytical Reconciliation Algorithm"):
        higher_rev = max(v_3b_4d1, v_g9_6h)
        deviation = (
            v_3b_4a5 
            - higher_rev 
            - v_8c_prev 
            + v_8c_curr 
            - v_2b_3ai 
            - v_2b_3bi 
            - v_2b_4b 
            + v_g9_t12
        )
        
        result_df = pd.DataFrame([{
            "GSTR-3B 4A(5) ITC": v_3b_4a5,
            "GSTR-3B 4D(1) reclaim": v_3b_4d1,
            "GSTR-9 6H reclaimed ITC": v_g9_6h,
            "GSTR-9 8C current year": v_8c_curr,
            "GSTR-9 8C previous year": v_8c_prev,
            "GSTR-2B 3A(I)": v_2b_3ai,
            "GSTR-2B 3B(I)": v_2b_3bi,
            "GSTR-2B 4B": v_2b_4b,
            "GSTR-9 Table 12": v_g9_t12,
            "Deviation": round(deviation, 2)
        }])
        st.dataframe(result_df, use_container_width=True)

with tab_downloads:
    st.subheader("📥 Download Auditor-Ready Excel Working Papers")
    w_dir = Path("./app_workspace")
    
    d1, d2, d3 = st.columns(3)
    
    f_3b = w_dir / "GSTR3B_Consolidated.xlsx"
    if f_3b.exists():
        with open(f_3b, "rb") as f:
            d1.download_button("Download GSTR-3B Working Papers (Excel)", f, file_name="GSTR3B_Consolidated_Working_Papers.xlsx")
            
    f_g1 = w_dir / "GSTR1_Consolidated.xlsx"
    if f_g1.exists():
        with open(f_g1, "rb") as f:
            d2.download_button("Download GSTR-1 Working Papers (Excel)", f, file_name="GSTR1_Consolidated_Working_Papers.xlsx")
            
    f_g9 = w_dir / "GSTR9_Comparative.xlsx"
    if f_g9.exists():
        with open(f_g9, "rb") as f:
            d3.download_button("Download GSTR-9 Comparative Working Papers (Excel)", f, file_name="GSTR9_Comparative_Working_Papers.xlsx")