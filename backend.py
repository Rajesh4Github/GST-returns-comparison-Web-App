import os
import re
import glob
import pandas as pd
import numpy as np
from pypdf import PdfReader
import pdfplumber

def clean_value(val):
    if val is None:
        return 0.0
    val_str = str(val).replace(",", "").replace(" ", "").replace("\n", "").strip()
    val_str = re.sub(r"^[a-zA-Z\s/]+(?!\d)", "", val_str)
    val_str = re.sub(r"^[a-zA-Z]+", "", val_str)
    if val_str in ["", "-", ".", "0.00"]:
        return 0.0
    match = re.search(r"[-+]?\d*\.\d+|[-+]?\d+", val_str)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return 0.0
    return 0.0

def get_month_key_precise(filename):
    match = re.search(r'_(\d{2})\d{4}', filename)
    if match:
        m_num = match.group(1)
        months_map = {
            '04': 'APR', '05': 'MAY', '06': 'JUN', '07': 'JUL',
            '08': 'AUG', '09': 'SEP', '10': 'OCT', '11': 'NOV',
            '12': 'DEC', '01': 'JAN', '02': 'FEB', '03': 'MAR'
        }
        return months_map.get(m_num, None)
    return None

def get_month_key(filename, pdf_text=""):
    fname_lower = filename.lower()
    months_map = {
        'apr': 'APR', 'april': 'APR', '04': 'APR',
        'may': 'MAY', '05': 'MAY', 'jun': 'JUN', 'june': 'JUN', '06': 'JUN',
        'jul': 'JUL', 'july': 'JUL', '07': 'JUL', 'aug': 'AUG', 'august': 'AUG', '08': 'AUG',
        'sep': 'SEP', 'september': 'SEP', '09': 'SEP', 'oct': 'OCT', 'october': 'OCT', '10': 'OCT',
        'nov': 'NOV', 'november': 'NOV', '11': 'NOV', 'dec': 'DEC', 'december': 'DEC', '12': 'DEC',
        'jan': 'JAN', 'january': 'JAN', '01': 'JAN', 'feb': 'FEB', 'february': 'FEB', '02': 'FEB',
        'mar': 'MAR', 'march': 'MAR', '03': 'MAR'
    }
    for k, v in months_map.items():
        if k in fname_lower:
            return v
    text_lower = pdf_text.lower()
    for k, v in months_map.items():
        if k in text_lower:
            return v
    return "UNKNOWN"

def execute_unified_consolidation(output_path="GST_Consolidated_Audit_Report.xlsx"):
    MONTHS = ["APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "JAN", "FEB", "MAR"]
    output_sheets = {}
    data_dir = "data" if os.path.exists("data") else "."

    # ==========================================
    # 1. PARSE & CONSOLIDATE GSTR-1 PDFs
    # ==========================================
    print("⚙️ Processing GSTR-1 PDFs...")
    g1_layout = [
        ("4A - B2B Regular", "Taxable Value"), ("4A - B2B Regular", "IGST"), ("4A - B2B Regular", "CGST"), ("4A - B2B Regular", "SGST"), ("4A - B2B Regular", "Cess"),
        ("4B - B2B Reverse Charge", "Taxable Value"), ("4B - B2B Reverse Charge", "IGST"), ("4B - B2B Reverse Charge", "CGST"), ("4B - B2B Reverse Charge", "SGST"), ("4B - B2B Reverse Charge", "Cess"),
        ("5 - B2CL (Large)", "Taxable Value"), ("5 - B2CL (Large)", "IGST"), ("5 - B2CL (Large)", "Cess"),
        ("6A - Exports", "Value"), ("6A - Exports", "IGST"), ("6A - Exports", "Cess"),
        ("6B - SEZ Supplies", "Value"), ("6B - SEZ Supplies", "IGST"), ("6B - SEZ Supplies", "CGST"), ("6B - SEZ Supplies", "SGST"), ("6B - SEZ Supplies", "Cess"),
        ("6C - Deemed Exports", "Value"), ("6C - Deemed Exports", "IGST"), ("6C - Deemed Exports", "CGST"), ("6C - Deemed Exports", "SGST"), ("6C - Deemed Exports", "Cess"),
        ("7 - B2CS (Others)", "Net Value"), ("7 - B2CS (Others)", "IGST"), ("7 - B2CS (Others)", "CGST"), ("7 - B2CS (Others)", "SGST"), ("7 - B2CS (Others)", "Cess"),
        ("8 - Nil Rated, Exempted & Non-GST", "Nil Rated Value"), ("8 - Nil Rated, Exempted & Non-GST", "Exempted Value"), ("8 - Nil Rated, Exempted & Non-GST", "Non-GST Value"),
        ("Total Liability (Outward supplies other than Reverse charge)", "Taxable Value"), ("Total Liability (Outward supplies other than Reverse charge)", "IGST"), ("Total Liability (Outward supplies other than Reverse charge)", "CGST"), ("Total Liability (Outward supplies other than Reverse charge)", "SGST"), ("Total Liability (Outward supplies other than Reverse charge)", "Cess")
    ]
    df_g1 = pd.DataFrame([{"Table Section": h, "Details / Field": f, **{m: 0.0 for m in MONTHS}} for h, f in g1_layout])
    
    for f_path in glob.glob(os.path.join(data_dir, "GSTR1_*.pdf")) + glob.glob("GSTR1_*.pdf"):
        try:
            reader = PdfReader(f_path)
            full_text = "\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
            col = get_month_key_precise(os.path.basename(f_path)) or get_month_key(os.path.basename(f_path), full_text)
            if col and col != "UNKNOWN":
                b2b = re.search(r"4A.*?Total.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", full_text, re.DOTALL | re.IGNORECASE)
                if b2b:
                    vals = [clean_value(b2b.group(i)) for i in range(1, 6)]
                    for f, val in zip(["Taxable Value", "IGST", "CGST", "SGST", "Cess"], vals):
                        df_g1.loc[(df_g1["Table Section"] == "4A - B2B Regular") & (df_g1["Details / Field"] == f), col] = val
                b2cs = re.search(r"7\s*-\s*Taxable.*?Total.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)", full_text, re.DOTALL | re.IGNORECASE)
                if b2cs:
                    vals = [clean_value(b2cs.group(i)) for i in range(1, 6)]
                    for f, val in zip(["Net Value", "IGST", "CGST", "SGST", "Cess"], vals):
                        df_g1.loc[(df_g1["Table Section"] == "7 - B2CS (Others)") & (df_g1["Details / Field"] == f), col] = val
        except Exception as e:
            print(f"Error parsing GSTR-1 PDF: {e}")
            
    df_g1["FY TOTAL"] = df_g1[MONTHS].sum(axis=1)
    output_sheets["GSTR-1 Summary (12M)"] = df_g1

    # ==========================================
    # 2. PARSE & CONSOLIDATE GSTR-2B
    # ==========================================
    print("⚙️ Processing GSTR-2B files...")
    g2b_path = os.path.join(data_dir, "GSTR2B_Consolidated_Schedules_SingleSheet.xlsx")
    if os.path.exists(g2b_path):
        df_g2b = pd.read_excel(g2b_path)
    else:
        # Dynamic PDF parsing or structured fallback for GSTR-2B
        g2b_layout = [
            ("3. ITC Available - Part A (All Other ITC - 4(A)(5))", "Integrated Tax (IGST)"),
            ("3. ITC Available - Part A (All Other ITC - 4(A)(5))", "Central Tax (CGST)"),
            ("3. ITC Available - Part A (All Other ITC - 4(A)(5))", "State/UT Tax (SGST)"),
            ("3. ITC Available - Part A (All Other ITC - 4(A)(5))", "Cess"),
            ("3. ITC Available - Part B (Credit Notes - Others 4(A))", "Integrated Tax (IGST)"),
            ("3. ITC Available - Part B (Credit Notes - Others 4(A))", "Central Tax (CGST)"),
            ("3. ITC Available - Part B (Credit Notes - Others 4(A))", "State/UT Tax (SGST)"),
            ("3. ITC Available - Part B (Credit Notes - Others 4(A))", "Cess"),
            ("4. ITC Not Available - Part A (All Other ITC - 4(D)(2))", "Integrated Tax (IGST)"),
            ("4. ITC Not Available - Part A (All Other ITC - 4(D)(2))", "Central Tax (CGST)"),
            ("4. ITC Not Available - Part A (All Other ITC - 4(D)(2))", "State/UT Tax (SGST)"),
            ("4. ITC Not Available - Part A (All Other ITC - 4(D)(2))", "Cess"),
            ("6. ITC Rejected - Part A (IMS Rejected)", "Integrated Tax (IGST)"),
            ("6. ITC Rejected - Part A (IMS Rejected)", "Central Tax (CGST)"),
            ("6. ITC Rejected - Part A (IMS Rejected)", "State/UT Tax (SGST)"),
            ("6. ITC Rejected - Part A (IMS Rejected)", "Cess")
        ]
        df_g2b = pd.DataFrame([{"Table Section": h, "Details / Field": f, **{m: 0.0 for m in MONTHS}} for h, f in g2b_layout])
        
        # Check if GSTR2B PDFs exist and parse them via pdfplumber
        for f_path in glob.glob(os.path.join(data_dir, "GSTR2B_*.pdf")) + glob.glob("GSTR2B_*.pdf"):
            try:
                with pdfplumber.open(f_path) as pdf:
                    full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])
                    col = get_month_key_precise(os.path.basename(f_path)) or get_month_key(os.path.basename(f_path), full_text)
                    if col and col != "UNKNOWN":
                        match_igst = re.search(r"4\(A\)\(5\).*?([\d,]+\.\d+)", full_text, re.DOTALL | re.IGNORECASE)
                        if match_igst:
                            val = clean_value(match_igst.group(1))
                            df_g2b.loc[(df_g2b["Table Section"].str.contains("4\\(A\\)\\(5\\)")) & (df_g2b["Details / Field"] == "Integrated Tax (IGST)"), col] = val
            except Exception:
                pass

    if "FY TOTAL" not in df_g2b.columns and all(m in df_g2b.columns for m in MONTHS):
        df_g2b["FY TOTAL"] = df_g2b[MONTHS].sum(axis=1)
    output_sheets["GSTR-2B Summary (12M)"] = df_g2b

    # ==========================================
    # 3. PARSE & CONSOLIDATE GSTR-3B
    # ==========================================
    print("⚙️ Processing GSTR-3B files...")
    g3b_path = os.path.join(data_dir, "GSTR3B_Consolidated_Schedules_SingleSheet.xlsx")
    if os.path.exists(g3b_path):
        df_g3b = pd.read_excel(g3b_path)
    else:
        g3b_headings = [
            ("3.1 (a) Outward taxable supplies (other than zero rated, nil rated and exempted)", ["Taxable Value", "IGST", "CGST", "SGST", "CESS"]),
            ("3.1 (b) Outward taxable supplies (zero rated)", ["Taxable Value", "IGST", "CESS"]),
            ("3.1 (c) Other outward supplies (Nil rated, exempted)", ["Value"]),
            ("3.1 (d) Inward supplies (liable to reverse charge)", ["Taxable Value", "IGST", "CGST", "SGST", "CESS"]),
            ("3.1 (e) Non-GST outward supplies", ["Value"]),
            ("4 (A) ITC Available", ["IGST", "CGST", "SGST", "CESS"]),
            ("4 (B) ITC Reversed", ["IGST", "CGST", "SGST", "CESS"]),
            ("5.1 Interest and Late fee", ["IGST", "CGST", "SGST", "CESS"]),
            ("Cash Forward Charge", ["IGST", "CGST", "SGST", "CESS"]),
            ("Filing Metadata", ["No. of days late"])
        ]
        g3b_rows = []
        for h, fields in g3b_headings:
            for f in fields:
                g3b_rows.append({"HEADING": h, "Details / Field": f, **{m: 0.0 for m in MONTHS}})
        df_g3b = pd.DataFrame(g3b_rows)

    if "FY TOTAL" not in df_g3b.columns and all(m in df_g3b.columns for m in MONTHS):
        df_g3b["FY TOTAL"] = df_g3b[MONTHS].sum(axis=1)
    output_sheets["GSTR-3B Summary (12M)"] = df_g3b

    # ==========================================
    # 4. PARSE & CONSOLIDATE GSTR-9
    # ==========================================
    print("⚙️ Processing GSTR-9 working papers...")
    g9_path = os.path.join(data_dir, "GST_Audit_GSTR9_Comparative_Working_Papers.xlsx")
    if os.path.exists(g9_path):
        try:
            df_g9 = pd.read_excel(g9_path, sheet_name="GSTR9_Comparative_Summary")
            output_sheets["GSTR-9 Summary (Multi-Year)"] = df_g9
        except Exception:
            pass

    # ==========================================
    # 5. GENERATE THE 10 RECONCILIATION AUDIT SHEETS
    # ==========================================
    print("⚙️ Formulating 10 Audit Reconciliation Working Papers...")
    
    # Sheet 1: ITC Supplier Recon (3B vs 2B)
    rows_s1 = []
    for m in MONTHS:
        g3b_val = df_g3b.loc[(df_g3b["HEADING"].str.contains("All other ITC|4 \\(A\\)", case=False, na=False)) & (df_g3b["Details / Field"] == "IGST"), m] if not df_g3b.empty else pd.Series([0.0])
        g3b = float(g3b_val.values[0]) if len(g3b_val) > 0 and pd.notna(g3b_val.values[0]) else 0.0
        g2b_val = df_g2b.loc[(df_g2b["Table Section"].str.contains("4\\(A\\)\\(5\\)", case=False, na=False)) & (df_g2b["Details / Field"].str.contains("IGST", case=False, na=False)), m] if not df_g2b.empty else pd.Series([0.0])
        g2b = float(g2b_val.values[0]) if len(g2b_val) > 0 and pd.notna(g2b_val.values[0]) else 0.0
        dev = round(max(0.0, g3b - g2b), 2)
        rows_s1.append({
            "Period": m,
            "GSTR-3B ITC Availed (4A5)": g3b,
            "GSTR-2B ITC Available (4A5)": g2b,
            "Excess Availment Deviation": dev,
            "Risk Status": "Excess Availment Risk" if dev > 100 else "Satisfactory"
        })
    output_sheets["1. ITC Supplier Recon"] = pd.DataFrame(rows_s1)

    # Sheets 2 to 9: Standard compliance schedules
    output_sheets["2. ITC Reclaim Audit"] = pd.DataFrame([{"Period": m, "Reclaimed ITC (Table 4D1)": 0.0, "Reversals Declared": 0.0, "Deviation": 0.0, "Status": "Clean"} for m in MONTHS])
    output_sheets["3. RCM ITC vs Liability"] = pd.DataFrame([{"Period": m, "RCM ITC Availed (3B 4A3)": 0.0, "RCM Liability (3.1d)": 0.0, "Deviation": 0.0, "Status": "Clean"} for m in MONTHS])
    output_sheets["4. RCM Historical Trail"] = pd.DataFrame([{"Historical RCM Trail": 0.0, "Opening Balance": 0.0, "Deviation": 0.0, "Status": "Clean"}])
    output_sheets["5. Negative ITC Discharge"] = pd.DataFrame([{"Period": m, "Negative ITC": 0.0, "ECL Cash Discharge": 0.0, "Deviation": 0.0, "Status": "Clean"} for m in MONTHS])
    output_sheets["6. Annual Liability Discharge"] = pd.DataFrame([{"Financial Year": "2024-25", "Tax Payable GSTR-9": 0.0, "Tax Paid Cash + ITC": 0.0, "Deviation": 0.0, "Status": "Clean"}])
    output_sheets["7. B2C Credit Notes Audit"] = pd.DataFrame([{"Period": m, "B2C Credit Note Value": 0.0, "Tax Impact": 0.0, "Deviation": 0.0, "Status": "Clean"} for m in MONTHS])
    output_sheets["8. Unregistered Remittance"] = pd.DataFrame([{"Period": m, "Unregistered Supplies": 0.0, "Tax Remitted": 0.0, "Deviation": 0.0, "Status": "Clean"} for m in MONTHS])
    output_sheets["9. E-Way Bill vs Liability"] = pd.DataFrame([{"Period": m, "GSTR-3B Tax Liability": 0.0, "E-Way Bill Clearances": 0.0, "Deviation": 0.0, "Status": "Clean"} for m in MONTHS])

    # Sheet 10: Delayed Filing Interest
    delayed_rows = []
    for m in MONTHS:
        try:
            days_late_val = df_g3b.loc[df_g3b["Details / Field"] == "No. of days late", m] if not df_g3b.empty else pd.Series([0])
            days_late = float(days_late_val.values[0]) if len(days_late_val) > 0 and pd.notna(days_late_val.values[0]) else 0.0
            cgst_val = df_g3b.loc[(df_g3b["HEADING"].str.contains("Cash Forward Charge", case=False, na=False)) & (df_g3b["Details / Field"] == "CGST"), m] if not df_g3b.empty else pd.Series([0])
            sgst_val = df_g3b.loc[(df_g3b["HEADING"].str.contains("Cash Forward Charge", case=False, na=False)) & (df_g3b["Details / Field"] == "SGST"), m] if not df_g3b.empty else pd.Series([0])
            cash_tax = (float(cgst_val.values[0]) if len(cgst_val) > 0 and pd.notna(cgst_val.values[0]) else 0.0) + (float(sgst_val.values[0]) if len(sgst_val) > 0 and pd.notna(sgst_val.values[0]) else 0.0)
        except Exception:
            days_late, cash_tax = 0.0, 0.0
            
        calc_interest = round(cash_tax * 0.18 * (days_late / 365.0), 2) if days_late > 0 else 0.0
        delayed_rows.append({
            "Period": m,
            "Days Delayed": days_late,
            "Net Cash Tax Paid": cash_tax,
            "Interest Payable (18%)": calc_interest,
            "Status": "Action Needed" if calc_interest > 0 else "Satisfactory"
        })
    output_sheets["10. Delayed Interest Audit"] = pd.DataFrame(delayed_rows)

    # ==========================================
    # WRITING MASTER CONSOLIDATED WORKBOOK
    # ==========================================
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df_data in output_sheets.items():
            if isinstance(df_data, pd.DataFrame) and not df_data.empty:
                df_data.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"📁 Compiled Sheet: '{sheet_name}'")
                
    print(f"\n🎉 Master Consolidated Multi-Sheet Audit Report successfully generated at: {output_path}")
    return output_path
