import os
import re
import glob
import pdfplumber
import pandas as pd
import numpy as np
import openpyxl
from pypdf import PdfReader

def clean_value(val):
    if val is None:
        return 0.0
    val_str = str(val).replace(",", "").replace(" ", "").replace("\n", "").strip()
    if val_str in ["", "-", ".", "0.00"]:
        return 0.0
    match = re.search(r"[-+]?\d*\.\d+|[-+]?\d+", val_str)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return 0.0
    return 0.0

def get_month_key(filename, pdf_text=""):
    fname_lower = filename.lower()
    months_map = {
        'apr':'APR','april':'APR','04':'APR',
        'may':'MAY','05':'MAY',
        'jun':'JUN','june':'JUN','06':'JUN',
        'jul':'JUL','july':'JUL','07':'JUL',
        'aug':'AUG','august':'AUG','08':'AUG',
        'sep':'SEP','september':'SEP','09':'SEP',
        'oct':'OCT','october':'OCT','10':'OCT',
        'nov':'NOV','november':'NOV','11':'NOV',
        'dec':'DEC','december':'DEC','12':'DEC',
        'jan':'JAN','january':'JAN','01':'JAN',
        'feb':'FEB','february':'FEB','02':'FEB',
        'mar':'MAR','march':'MAR','03':'MAR'
    }
    for k,v in months_map.items():
        if k in fname_lower:
            return v
    match = re.search(r"_(\d{2})\d{4}", filename)
    if match:
        return months_map.get(match.group(1), None)
    text_lower = pdf_text.lower()
    for k,v in months_map.items():
        if k in text_lower:
            return v
    return None
    
def execute_unified_consolidation(output_path="GST_Consolidated_Audit_Report.xlsx"):
    MONTHS = ["APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC","JAN","FEB","MAR"]
    output_sheets = {}
    data_dir = "data" if os.path.exists("data") else "."

    # ---------------- GSTR-1 ----------------
    gstr1_layout = [
        ("4A - B2B Regular","Taxable Value"),("4A - B2B Regular","IGST"),("4A - B2B Regular","CGST"),("4A - B2B Regular","SGST"),("4A - B2B Regular","Cess"),
        ("4B - B2B Reverse Charge","Taxable Value"),("4B - B2B Reverse Charge","IGST"),("4B - B2B Reverse Charge","CGST"),("4B - B2B Reverse Charge","SGST"),("4B - B2B Reverse Charge","Cess"),
        ("5 - B2CL (Large)","Taxable Value"),("5 - B2CL (Large)","IGST"),("5 - B2CL (Large)","Cess"),
        ("6A - Exports","Value"),("6A - Exports","IGST"),("6A - Exports","Cess"),
        ("6B - SEZ Supplies","Value"),("6B - SEZ Supplies","IGST"),("6B - SEZ Supplies","CGST"),("6B - SEZ Supplies","SGST"),("6B - SEZ Supplies","Cess"),
        ("6C - Deemed Exports","Value"),("6C - Deemed Exports","IGST"),("6C - Deemed Exports","CGST"),("6C - Deemed Exports","SGST"),("6C - Deemed Exports","Cess"),
        ("7 - B2CS (Others)","Net Value"),("7 - B2CS (Others)","IGST"),("7 - B2CS (Others)","CGST"),("7 - B2CS (Others)","SGST"),("7 - B2CS (Others)","Cess"),
        ("8 - Nil Rated, Exempted & Non-GST","Nil Rated Value"),("8 - Nil Rated, Exempted & Non-GST","Exempted Value"),("8 - Nil Rated, Exempted & Non-GST","Non-GST Value"),
        ("Total Liability (Outward supplies other than Reverse charge)","Taxable Value"),("Total Liability (Outward supplies other than Reverse charge)","IGST"),("Total Liability (Outward supplies other than Reverse charge)","CGST"),("Total Liability (Outward supplies other than Reverse charge)","SGST"),("Total Liability (Outward supplies other than Reverse charge)","Cess")
    ]
    df_g1 = pd.DataFrame([{"Table Section":h,"Details / Field":f,**{m:0.0 for m in MONTHS}} for h,f in gstr1_layout])
    for f_path in glob.glob(os.path.join(data_dir,"*GSTR1*.pdf")):
        reader=PdfReader(f_path)
        full_text="\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
        col=get_month_key(os.path.basename(f_path),full_text)
        if col and col in MONTHS:
            df_g1=map_pdf_tables_to_layout(f_path,df_g1,col)
    df_g1["FY TOTAL"]=df_g1[MONTHS].sum(axis=1)
    output_sheets["GSTR-1 Summary (12M)"]=df_g1

    # ---------------- GSTR-2B ----------------
    gstr2b_layout = [
        ("3. ITC Available - Part A (All Other ITC - 4(A)(5))","Integrated Tax (IGST)"),
        ("3. ITC Available - Part A (All Other ITC - 4(A)(5))","Central Tax (CGST)"),
        ("3. ITC Available - Part A (All Other ITC - 4(A)(5))","State/UT Tax (SGST)"),
        ("3. ITC Available - Part A (All Other ITC - 4(A)(5))","Cess"),
        ("3. ITC Available - Part B (Credit Notes - Others 4(A))","Integrated Tax (IGST)"),
        ("3. ITC Available - Part B (Credit Notes - Others 4(A))","Central Tax (CGST)"),
        ("3. ITC Available - Part B (Credit Notes - Others 4(A))","State/UT Tax (SGST)"),
        ("3. ITC Available - Part B (Credit Notes - Others 4(A))","Cess"),
        ("4. ITC Not Available - Part A (All Other ITC - 4(D)(2))","Integrated Tax (IGST)"),
        ("4. ITC Not Available - Part A (All Other ITC - 4(D)(2))","Central Tax (CGST)"),
        ("4. ITC Not Available - Part A (All Other ITC - 4(D)(2))","State/UT Tax (SGST)"),
        ("4. ITC Not Available - Part A (All Other ITC - 4(D)(2))","Cess"),
        ("6. ITC Rejected - Part A (IMS Rejected)","Integrated Tax (IGST)"),
        ("6. ITC Rejected - Part A (IMS Rejected)","Central Tax (CGST)"),
        ("6. ITC Rejected - Part A (IMS Rejected)","State/UT Tax (SGST)"),
        ("6. ITC Rejected - Part A (IMS Rejected)","Cess")
    ]
    df_g2b = pd.DataFrame([{"Table Section":h,"Details / Field":f,**{m:0.0 for m in MONTHS}} for h,f in gstr2b_layout])
    for f_path in glob.glob(os.path.join(data_dir,"*GSTR2B*.pdf")):
        reader=PdfReader(f_path)
        full_text="\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
        col=get_month_key(os.path.basename(f_path),full_text)
        if col and col in MONTHS:
            df_g2b=map_pdf_tables_to_layout(f_path,df_g2b,col)
    df_g2b["FY TOTAL"]=df_g2b[MONTHS].sum(axis=1)
    output_sheets["GSTR-2B Summary (12M)"]=df_g2b

    # ---------------- GSTR-3B ----------------
    gstr3b_layout = [
        ("3.1 (a) Outward taxable supplies","Taxable Value"),("3.1 (a) Outward taxable supplies","IGST"),("3.1 (a) Outward taxable supplies","CGST"),("3.1 (a) Outward taxable supplies","SGST"),("3.1 (a) Outward taxable supplies","CESS"),
        ("3.1 (b) Outward taxable supplies (zero rated)","Taxable Value"),("3.1 (b) Outward taxable supplies (zero rated)","IGST"),("3.1 (b) Outward taxable supplies (zero rated)","CESS"),
        ("3.1 (c) Other outward supplies","Value"),
        ("3.1 (d) Inward supplies liable to reverse charge","Taxable Value"),("3.1 (d) Inward supplies liable to reverse charge","IGST"),("3.1 (d) Inward supplies liable to reverse charge","CGST"),("3.1 (d) Inward supplies liable to reverse charge","SGST"),("3.1 (d) Inward supplies liable to reverse charge","CESS"),
        ("3.1 (e) Non-GST outward supplies","Value"),
        ("4 (A) ITC Available","IGST"),("4 (A) ITC Available","CGST"),("4 (A) ITC Available","SGST"),("4 (A) ITC Available","CESS"),
        ("4 (B) ITC Reversed","IGST"),("4 (B) ITC Reversed","CGST"),("4 (B) ITC Reversed","SGST"),("4 (B) ITC Reversed","CESS"),
        ("4 (C) Net ITC Available","IGST"),("4 (C) Net ITC Available","CGST"),("4 (C) Net ITC Available","SGST"),("4 (C) Net ITC Available","CESS"),
        ("4 (D) Ineligible ITC","IGST"),("4 (D) Ineligible ITC","CGST"),("4 (D) Ineligible ITC","SGST"),("4 (D) Ineligible ITC","CESS"),
        ("5.1 Interest and Late fee","IGST"),("5.1 Interest and Late fee","CGST"),("5.1 Interest and Late fee","SGST"),("5.1 Interest and Late fee","CESS"),
                ("6.1 Payment of Tax - Using ITC","IGST"),
        ("6.1 Payment of Tax - Using ITC","CGST"),
        ("6.1 Payment of Tax - Using ITC","SGST"),
        ("6.1 Payment of Tax - Using ITC","CESS"),
        ("6.1 Payment of Tax - Cash Forward Charge","IGST"),
        ("6.1 Payment of Tax - Cash Forward Charge","CGST"),
        ("6.1 Payment of Tax - Cash Forward Charge","SGST"),
        ("6.1 Payment of Tax - Cash Forward Charge","CESS"),
        ("6.1 Payment of Tax - Cash Reverse Charge","IGST"),
        ("6.1 Payment of Tax - Cash Reverse Charge","CGST"),
        ("6.1 Payment of Tax - Cash Reverse Charge","SGST"),
        ("6.1 Payment of Tax - Cash Reverse Charge","CESS")
    ]
    df_g3b = pd.DataFrame([{"HEADING":h,"Details / Field":f,**{m:0.0 for m in MONTHS}} for h,f in gstr3b_layout])
    for f_path in glob.glob(os.path.join(data_dir,"*GSTR3B*.pdf")):
        reader=PdfReader(f_path)
        full_text="\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
        col=get_month_key(os.path.basename(f_path),full_text)
        if col and col in MONTHS:
            df_g3b=map_pdf_tables_to_layout(f_path,df_g3b,col)
    df_g3b["FY TOTAL"]=df_g3b[MONTHS].sum(axis=1)
    output_sheets["GSTR-3B Summary (12M)"]=df_g3b

    # ---------------- GSTR-9 ----------------
    gstr9_layout = [
        ("Table 4 - B2B Supplies","Value"),
        ("Table 6A - Total ITC Availed","Value"),
        ("Table 7I - Total ITC Reversed","Value")
    ]
    df_g9 = pd.DataFrame([{"Parameter":h,"Details / Field":f,"Value":0.0} for h,f in gstr9_layout])
    for f_path in glob.glob(os.path.join(data_dir,"*GSTR9*.pdf")):
        reader=PdfReader(f_path)
        full_text="\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
        # For GSTR-9 we can still use the generalized parser
        df_g9=map_pdf_tables_to_layout(f_path,df_g9,"Value")
    output_sheets["GSTR-9 Summary (Multi-Year)"]=df_g9

    # ---------------- SAVE MASTER REPORT ----------------
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df_data in output_sheets.items():
            if isinstance(df_data, pd.DataFrame) and not df_data.empty:
                df_data.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Included Sheet: {sheet_name}")

    print(f"✅ Master report created at: {output_path}")
    return output_path
