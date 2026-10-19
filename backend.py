import os, re, glob
import pandas as pd
import pdfplumber
from backend_utils import clean_value, get_month_key

MONTHS = ["APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC","JAN","FEB","MAR"]

# ---------------- Parsing Functions ----------------

def parse_gstr1_pdfs():
    """Parse GSTR-1 PDFs into a 12-month DataFrame."""
    layout = [
        ("4A - B2B Regular", ["Taxable Value","IGST","CGST","SGST","Cess"]),
        ("4B - B2B Reverse Charge", ["Taxable Value","IGST","CGST","SGST","Cess"]),
        ("5 - B2CL (Large)", ["Taxable Value","IGST","Cess"]),
        ("6A - Exports", ["Value","IGST","Cess"]),
        ("6B - SEZ Supplies", ["Value","IGST","CGST","SGST","Cess"]),
        ("6C - Deemed Exports", ["Value","IGST","CGST","SGST","Cess"]),
        ("7 - B2CS (Others)", ["Net Value","IGST","CGST","SGST","Cess"]),
    ]
    df = pd.DataFrame([{"Table Section": sec, "Details / Field": f, **{m:0.0 for m in MONTHS}} 
                       for sec, fields in layout for f in fields])
    for f_path in glob.glob("GSTR1_*.pdf"):
        try:
            with pdfplumber.open(f_path) as pdf:
                full_text = "\n".join([p.extract_text() or "" for p in pdf.pages])
            col = get_month_key(os.path.basename(f_path), full_text)
            if col and col != "UNKNOWN":
                # Regex for B2B Regular
                match_b2b = re.search(
                    r"B2B Regular.*?Total.*?Invoice\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)",
                    full_text, re.DOTALL
                )
                if match_b2b:
                    taxable, igst, cgst, sgst, cess = [clean_value(match_b2b.group(i)) for i in range(1,6)]
                    df.loc[(df["Table Section"]=="4A - B2B Regular") & (df["Details / Field"]=="Taxable Value"), col] = taxable
                    df.loc[(df["Table Section"]=="4A - B2B Regular") & (df["Details / Field"]=="IGST"), col] = igst
                    df.loc[(df["Table Section"]=="4A - B2B Regular") & (df["Details / Field"]=="CGST"), col] = cgst
                    df.loc[(df["Table Section"]=="4A - B2B Regular") & (df["Details / Field"]=="SGST"), col] = sgst
                    df.loc[(df["Table Section"]=="4A - B2B Regular") & (df["Details / Field"]=="Cess"), col] = cess
                # Regex for B2CS (Others)
                match_b2cs = re.search(
                    r"B2CS.*?Total.*?Invoice\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)",
                    full_text, re.DOTALL
                )
                if match_b2cs:
                    netval, igst, cgst, sgst, cess = [clean_value(match_b2cs.group(i)) for i in range(1,6)]
                    df.loc[(df["Table Section"]=="7 - B2CS (Others)") & (df["Details / Field"]=="Net Value"), col] = netval
                    df.loc[(df["Table Section"]=="7 - B2CS (Others)") & (df["Details / Field"]=="IGST"), col] = igst
                    df.loc[(df["Table Section"]=="7 - B2CS (Others)") & (df["Details / Field"]=="CGST"), col] = cgst
                    df.loc[(df["Table Section"]=="7 - B2CS (Others)") & (df["Details / Field"]=="SGST"), col] = sgst
                    df.loc[(df["Table Section"]=="7 - B2CS (Others)") & (df["Details / Field"]=="Cess"), col] = cess
        except Exception as e:
            print("Error parsing GSTR-1:", e)
    df["FY TOTAL"] = df[MONTHS].sum(axis=1)
    return df


def parse_gstr2b_pdfs():
    """Parse GSTR-2B PDFs into a 12-month DataFrame."""
    layout = [
        ("3. ITC Available - Part A (All Other ITC - 4(A)(5))", ["IGST","CGST","SGST","Cess"])
    ]
    df = pd.DataFrame([{"Table Section": sec, "Details / Field": f, **{m:0.0 for m in MONTHS}} 
                       for sec, fields in layout for f in fields])
    for f_path in glob.glob("GSTR2B_*.pdf"):
        try:
            with pdfplumber.open(f_path) as pdf:
                full_text = "\n".join([p.extract_text() or "" for p in pdf.pages])
            col = get_month_key(os.path.basename(f_path), full_text)
            if col and col != "UNKNOWN":
                match_itc = re.search(
                    r"4\(A\)\(5\)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)",
                    full_text
                )
                if match_itc:
                    igst, cgst, sgst, cess = [clean_value(match_itc.group(i)) for i in range(1,5)]
                    df.loc[(df["Details / Field"]=="IGST"), col] = igst
                    df.loc[(df["Details / Field"]=="CGST"), col] = cgst
                    df.loc[(df["Details / Field"]=="SGST"), col] = sgst
                    df.loc[(df["Details / Field"]=="Cess"), col] = cess
        except Exception as e:
            print("Error parsing GSTR-2B:", e)
    df["FY TOTAL"] = df[MONTHS].sum(axis=1)
    return df


def parse_gstr3b_pdfs():
    """Parse GSTR-3B PDFs into a 12-month DataFrame."""
    layout = [
        ("3.1 (a) Outward taxable supplies", ["Taxable Value","IGST","CGST","SGST","CESS"]),
        ("4A ITC Availed", ["IGST","CGST","SGST","Cess"]),
        ("Late Filing", ["No. of days late"])
    ]
    df = pd.DataFrame([{"HEADING": sec, "Details / Field": f, **{m:0.0 for m in MONTHS}} 
                       for sec, fields in layout for f in fields])
    for f_path in glob.glob("GSTR3B_*.pdf"):
        try:
            with pdfplumber.open(f_path) as pdf:
                full_text = "\n".join([p.extract_text() or "" for p in pdf.pages])
            col = get_month_key(os.path.basename(f_path), full_text)
            if col and col != "UNKNOWN":
                match_liab = re.search(
                    r"Outward taxable supplies.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)",
                    full_text
                )
                if match_liab:
                    taxable, igst, cgst, sgst, cess = [clean_value(match_liab.group(i)) for i in range(1,6)]
                    df.loc[(df["HEADING"]=="3.1 (a) Outward taxable supplies") & (df["Details / Field"]=="Taxable Value"), col] = taxable
                    df.loc[(df["HEADING"]=="3.1 (a) Outward taxable supplies") & (df["Details / Field"]=="IGST"), col] = igst
                    df.loc[(df["HEADING"]=="3.1 (a) Outward taxable supplies") & (df["Details / Field"]=="CGST"), col] = cgst
                    df.loc[(df["HEADING"]=="3.1 (a) Outward taxable supplies") & (df["Details / Field"]=="SGST"), col] = sgst
                    df.loc[(df["HEADING"]=="3.1 (a) Outward taxable supplies") & (df["Details / Field"]=="CESS"), col] = cess
                match_late = re.search(r"No\. of days late.*?(\d+)", full_text)
                if match_late:
                    df.loc[(df["Details / Field"]=="No. of days late"), col] = clean_value(match_late.group(1))
        except Exception as e:
            print("Error parsing GSTR-3B:", e)
    return df


def parse_gstr9_pdfs():
    """Parse filed GSTR-9 PDFs into a multi-year DataFrame."""
    records = []
    for f_path in glob.glob("GSTR9_*.pdf"):
        fy = "2023-24" if "2023" in f_path else "2024-25"
        data = {
            "Financial Year": fy,
            "Table 4 - B2B Supplies": 0.0,
            "Table 6A - Total ITC Availed": 0.0,
            "Table 7I - Total ITC Reversed": 0.0
        }
        try:
            with pdfplumber.open(f_path) as pdf:
                full_text = "\n".join([p.extract_text() or "" for p in pdf.pages])

                # Skip system-computed GSTR-9 (no legal validity)
                if "System Computed GSTR-9" in full_text:
                    print(f"Skipping system-computed GSTR-9: {f_path}")
                    continue

                # Match B2B supplies line
                match_b2b = re.search(
                    r"Supplies made to registered persons.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)",
                    full_text
                )
                if match_b2b:
                    taxable, cgst, sgst, igst, cess = [clean_value(match_b2b.group(i)) for i in range(1,6)]
                    data["Table 4 - B2B Supplies"] = taxable

                # Match ITC Availed (Table 6A)
                match_itc = re.search(
                    r"Total amount of input tax credit availed.*?\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)",
                    full_text
                )
                if match_itc:
                    igst, cgst, sgst, cess = [clean_value(match_itc.group(i)) for i in range(1,5)]
                    data["Table 6A - Total ITC Availed"] = igst + cgst + sgst

                # Match ITC Reversed (Table 7I)
                match_rev = re.search(
                    r"Total ITC Reversed.*?\s+([\d,]+\.\d+)",
                    full_text
                )
                if match_rev:
                    data["Table 7I - Total ITC Reversed"] = clean_value(match_rev.group(1))

        except Exception as e:
            print("Error parsing GSTR-9:", e)

        records.append(data)

    return pd.DataFrame(records)

# ---------------- Consolidation Pipeline ----------------

def execute_unified_consolidation(output_path="GST_Consolidated_Audit_Report.xlsx"):
    output_sheets = {}

    # --- Parse each return type directly from PDFs ---
    gstr1_df = parse_gstr1_pdfs()
    gstr2b_df = parse_gstr2b_pdfs()
    gstr3b_df = parse_gstr3b_pdfs()
    gstr9_df = parse_gstr9_pdfs()

    output_sheets["GSTR-1 Summary (12M)"] = gstr1_df
    output_sheets["GSTR-2B Summary (12M)"] = gstr2b_df
    output_sheets["GSTR-3B Summary (12M)"] = gstr3b_df
    output_sheets["GSTR-9 Summary (2 Year)"] = gstr9_df

    # --- Run reconciliation / audit algorithms ---
    # 1. ITC Supplier Recon
    recon_rows = []
    for m in MONTHS:
        g3b_val = gstr3b_df.loc[(gstr3b_df["HEADING"]=="4A ITC Availed") & (gstr3b_df["Details / Field"]=="IGST"), m] if not gstr3b_df.empty else []
        g3b = float(g3b_val.values[0]) if len(g3b_val)>0 else 0.0
        g2b_val = gstr2b_df.loc[(gstr2b_df["Table Section"].str.contains("4\\(A\\)\\(5\\)", na=False)) & (gstr2b_df["Details / Field"]=="IGST"), m] if not gstr2b_df.empty else []
        g2b = float(g2b_val.values[0]) if len(g2b_val)>0 else 0.0
        deviation = round(max(0.0, g3b-g2b),2)
        recon_rows.append({
            "Period": m,
            "GSTR-3B ITC Availed": g3b,
            "GSTR-2B ITC Available": g2b,
            "Deviation": deviation,
            "Risk Status": "Excess Availment Risk" if deviation>100 else "Satisfactory"
        })
    output_sheets["1. ITC Supplier Recon"] = pd.DataFrame(recon_rows)

    # Placeholders for algorithms 2–9
    for name in [
        "2. ITC Reclaim Audit","3. RCM ITC vs Liability","4. RCM Historical Trail",
        "5. Negative ITC Discharge","6. Annual Liability Discharge",
        "7. B2C Credit Notes Audit","8. Unregistered Remittance","9. E-Way Bill vs Liability"
    ]:
        output_sheets[name] = pd.DataFrame([{"Message":"Placeholder"}])

    # 10. Delayed Interest Audit
    interest_rows = []
    for m in MONTHS:
        days_late_val = gstr3b_df.loc[gstr3b_df["Details / Field"]=="No. of days late", m] if not gstr3b_df.empty else []
        days_late = float(days_late_val.values[0]) if len(days_late_val)>0 else 0.0
        cgst_val = gstr3b_df.loc[(gstr3b_df["HEADING"]=="3.1 Tax Liability") & (gstr3b_df["Details / Field"]=="CGST"), m] if not gstr3b_df.empty else []
        sgst_val = gstr3b_df.loc[(gstr3b_df["HEADING"]=="3.1 Tax Liability") & (gstr3b_df["Details / Field"]=="SGST"), m] if not gstr3b_df.empty else []
        cash_tax = (float(cgst_val.values[0]) if len(cgst_val)>0 else 0.0) + (float(sgst_val.values[0]) if len(sgst_val)>0 else 0.0)
        calc_interest = round(cash_tax*0.18*(days_late/365.0),2) if days_late>0 else 0.0
        interest_rows.append({
            "Period": m,
            "Days Delayed": days_late,
            "Net Cash Tax Paid": cash_tax,
            "Interest Payable (18%)": calc_interest,
            "Status": "Action Needed" if calc_interest>0 else "Satisfactory"
        })
    output_sheets["10. Delayed Interest Audit"] = pd.DataFrame(interest_rows)

    # --- Safeguard: ensure at least one sheet ---
    if not output_sheets:
        output_sheets["Info"] = pd.DataFrame([{"Message":"No GST data files found"}])

    # --- Save workbook ---
    print("\n📁 Saving unified sheets to Excel...")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df_data in output_sheets.items():
            if isinstance(df_data, pd.DataFrame) and not df_data.empty:
                df_data.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"   -> Included Sheet: '{sheet_name}'")
            else:
                pd.DataFrame([{"Message":"No data"}]).to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\n🎉 SUCCESS! Unified Consolidated Audit workbook created at: {output_path}")
    return output_path
