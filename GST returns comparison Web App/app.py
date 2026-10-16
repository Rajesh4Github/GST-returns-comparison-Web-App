import streamlit as st
import pandas as pd
import glob
from gst_parser import parse_and_populate_gstr1_pdf, parse_and_populate_gstr2b_pdf, parse_and_populate_gstr3b_pdf

st.title("GST Returns Audit & Reconciliation")

uploaded_files = st.file_uploader("Upload GST PDFs", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    # Temporary storage
    for f in uploaded_files:
        with open(f.name, "wb") as out:
            out.write(f.getbuffer())

    # Initialize empty DataFrames
    gstr1_df = pd.DataFrame()
    gstr2b_df = pd.DataFrame()
    gstr3b_df = pd.DataFrame()

    # Run parsers
    for f in uploaded_files:
        if "GSTR1" in f.name.upper():
            gstr1_df = parse_and_populate_gstr1_pdf(f.name, gstr1_df)
        elif "GSTR2B" in f.name.upper():
            gstr2b_df = parse_and_populate_gstr2b_pdf(f.name, gstr2b_df)
        elif "GSTR3B" in f.name.upper():
            gstr3b_df = parse_and_populate_gstr3b_pdf(f.name, gstr3b_df)

    # Display results
    if not gstr1_df.empty:
        st.subheader("📊 GSTR-1 Extract")
        st.dataframe(gstr1_df)

    if not gstr2b_df.empty:
        st.subheader("📊 GSTR-2B Extract")
        st.dataframe(gstr2b_df)

    if not gstr3b_df.empty:
        st.subheader("📊 GSTR-3B Extract")
        st.dataframe(gstr3b_df)

    # Consolidated Excel download
    with pd.ExcelWriter("GST_Returns_Consolidated.xlsx") as writer:
        if not gstr1_df.empty: gstr1_df.to_excel(writer, sheet_name="GSTR-1", index=False)
        if not gstr2b_df.empty: gstr2b_df.to_excel(writer, sheet_name="GSTR-2B", index=False)
        if not gstr3b_df.empty: gstr3b_df.to_excel(writer, sheet_name="GSTR-3B", index=False)

    with open("GST_Returns_Consolidated.xlsx", "rb") as f:
        st.download_button("⬇️ Download Consolidated Workbook", f, file_name="GST_Returns_Consolidated.xlsx")
