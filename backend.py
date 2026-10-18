import os, glob, pdfplumber, pandas as pd, re

# Helper: clean numeric values
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

# Generic parser for any GST PDF
def parse_pdf_tables(file):
    records = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                for col in df.columns:
                    df[col] = df[col].apply(clean_value)
                df["Source File"] = os.path.basename(file)
                records.append(df)
    return pd.concat(records) if records else pd.DataFrame()

# Specialized wrappers
def parse_gstr1_pdf(file): return parse_pdf_tables(file)
def parse_gstr2b_pdf(file): return parse_pdf_tables(file)
def parse_gstr3b_pdf(file): return parse_pdf_tables(file)
def parse_gstr9_pdf(file): return parse_pdf_tables(file)

def execute_final_consolidation(output_path="GST_Consolidated_Audit_Report.xlsx"):
    data_dir = "data" if os.path.exists("data") else "."
    output_sheets = {}

    # --- Parse PDFs ---
    g1_dfs = [parse_gstr1_pdf(f) for f in glob.glob(os.path.join(data_dir,"GSTR1_*.pdf"))]
    output_sheets["GSTR-1 Summary (12M)"] = pd.concat(g1_dfs) if g1_dfs else pd.DataFrame()

    g2b_dfs = [parse_gstr2b_pdf(f) for f in glob.glob(os.path.join(data_dir,"GSTR2B_*.pdf"))]
    output_sheets["GSTR-2B Summary (12M)"] = pd.concat(g2b_dfs) if g2b_dfs else pd.DataFrame()

    g3b_dfs = [parse_gstr3b_pdf(f) for f in glob.glob(os.path.join(data_dir,"GSTR3B_*.pdf"))]
    output_sheets["GSTR-3B Summary (12M)"] = pd.concat(g3b_dfs) if g3b_dfs else pd.DataFrame()

    g9_dfs = [parse_gstr9_pdf(f) for f in glob.glob(os.path.join(data_dir,"GSTR9_*.pdf"))]
    output_sheets["GSTR-9 Summary (2 Year)"] = pd.concat(g9_dfs) if g9_dfs else pd.DataFrame()

    # --- Generate 10 Reconciliation Audit Sheets ---
    df_g1 = output_sheets["GSTR-1 Summary (12M)"]
    df_g2b = output_sheets["GSTR-2B Summary (12M)"]
    df_g3b = output_sheets["GSTR-3B Summary (12M)"]
    df_g9 = output_sheets["GSTR-9 Summary (2 Year)"]

    # 1. ITC Supplier Recon
    recon_rows = []
    if not df_g3b.empty and not df_g2b.empty:
        for m in df_g3b.columns:
            if m in ["Source File"]: continue
            g3b_val = df_g3b.loc[df_g3b["Details / Field"].str.contains("IGST", case=False, na=False), m].sum()
            g2b_val = df_g2b.loc[df_g2b["Details / Field"].str.contains("IGST", case=False, na=False), m].sum()
            dev = round(max(0.0, g3b_val - g2b_val), 2)
            recon_rows.append({
                "Period": m,
                "GSTR-3B ITC Availed": g3b_val,
                "GSTR-2B ITC Available": g2b_val,
                "Deviation": dev,
                "Risk Status": "Excess Availment Risk" if dev > 100 else "Satisfactory"
            })
    output_sheets["1. ITC Supplier Recon"] = pd.DataFrame(recon_rows)

    # 2. ITC Reclaim Audit
    reclaim_rows = []
    for m in df_g3b.columns:
        if m in ["Source File"]: continue
        reclaimed = df_g3b.loc[df_g3b["Details / Field"].str.contains("Reclaimed", case=False, na=False), m].sum()
        reversals = df_g3b.loc[df_g3b["Details / Field"].str.contains("Reversal", case=False, na=False), m].sum()
        reclaim_rows.append({
            "Period": m,
            "Reclaimed ITC": reclaimed,
            "Reversals Declared": reversals,
            "Deviation": reclaimed - reversals,
            "Status": "Clean" if abs(reclaimed - reversals) < 100 else "Mismatch"
        })
    output_sheets["2. ITC Reclaim Audit"] = pd.DataFrame(reclaim_rows)

    # 3. RCM ITC vs Liability
    rcm_rows = []
    for m in df_g3b.columns:
        if m in ["Source File"]: continue
        itc_rcm = df_g3b.loc[df_g3b["Details / Field"].str.contains("RCM", case=False, na=False), m].sum()
        liability_rcm = df_g3b.loc[df_g3b["HEADING"].str.contains("Inward supplies", case=False, na=False), m].sum()
        rcm_rows.append({
            "Period": m,
            "RCM ITC Availed": itc_rcm,
            "RCM Liability": liability_rcm,
            "Deviation": itc_rcm - liability_rcm,
            "Status": "Matched" if abs(itc_rcm - liability_rcm) < 100 else "Mismatch"
        })
    output_sheets["3. RCM ITC vs Liability"] = pd.DataFrame(rcm_rows)

    # 4. RCM Historical Trail
    output_sheets["4. RCM Historical Trail"] = pd.DataFrame([{
        "Historical RCM Trail": "Deferred RCM",
        "Opening Balance": df_g3b.sum(numeric_only=True).sum(),
        "Deviation": 0,
        "Status": "Reconciled"
    }])

    # 5. Negative ITC Discharge
    neg_rows = []
    for m in df_g3b.columns:
        if m in ["Source File"]: continue
        neg_itc = df_g3b.loc[df_g3b["Details / Field"].str.contains("Negative", case=False, na=False), m].sum()
        cash_discharge = df_g3b.loc[df_g3b["Details / Field"].str.contains("Cash", case=False, na=False), m].sum()
        neg_rows.append({
            "Period": m,
            "Negative ITC": neg_itc,
            "ECL Cash Discharge": cash_discharge,
            "Deviation": neg_itc - cash_discharge,
            "Status": "Clean" if neg_itc == 0 else "Action Needed"
        })
    output_sheets["5. Negative ITC Discharge"] = pd.DataFrame(neg_rows)

    # 6. Annual Liability Discharge
    annual_rows = []
    if not df_g9.empty:
        for _, row in df_g9.iterrows():
            fy = row.get("Financial Year", "Unknown")
            tax_payable = row.get("Table 4 - B2B Supplies", 0)
            tax_paid = row.get("Table 6A - Total ITC Availed", 0) - row.get("Table 7I - Total ITC Reversed", 0)
            annual_rows.append({
                "Financial Year": fy,
                "Tax Payable GSTR-9": tax_payable,
                "Tax Paid Cash + ITC": tax_paid,
                "Deviation": tax_payable - tax_paid,
                "Status": "Satisfactory" if abs(tax_payable - tax_paid) < 100 else "Mismatch"
            })
    output_sheets["6. Annual Liability Discharge"] = pd.DataFrame(annual_rows)

    # 7. B2C Credit Notes Audit
    b2c_rows = []
    for m in df_g1.columns:
        if m in ["Source File"]: continue
        credit_val = df_g1.loc[df_g1["Table Section"].str.contains("B2CS", case=False, na=False), m].sum()
        b2c_rows.append({
            "Period": m,
            "B2C Credit Note Value": credit_val,
            "Tax Impact": credit_val * 0.18,
            "Deviation": 0,
            "Status": "Clean"
        })
    output_sheets["7. B2C Credit Notes Audit"] = pd.DataFrame(b2c_rows)

        # 8. Unregistered Remittance
    unreg_rows = []
    for m in df_g1.columns:
        if m in ["Source File"]: 
            continue
        unreg_val = df_g1.loc[df_g1["Table Section"].str.contains("Unregistered", case=False, na=False), m].sum()
        unreg_rows.append({
            "Period": m,
            "Unregistered Supplies": unreg_val,
            "Tax Remitted": unreg_val * 0.18,
            "Deviation": 0,
            "Status": "Clean"
        })
    output_sheets["8. Unregistered Remittance"] = pd.DataFrame(unreg_rows)

    # 9. E-Way Bill vs Liability
    eway_rows = []
    for m in df_g3b.columns:
        if m in ["Source File"]: 
            continue
        liability = df_g3b.loc[df_g3b["HEADING"].str.contains("Outward taxable supplies", case=False, na=False), m].sum()
        eway_val = df_g1.loc[df_g1["Table Section"].str.contains("E-Way", case=False, na=False), m].sum()
        eway_rows.append({
            "Period": m,
            "GSTR-3B Tax Liability": liability,
            "E-Way Bill Clearances": eway_val,
            "Deviation": liability - eway_val,
            "Status": "Matched" if abs(liability - eway_val) < 100 else "Mismatch"
        })
    output_sheets["9. E-Way Bill vs Liability"] = pd.DataFrame(eway_rows)

    # 10. Delayed Interest Audit
    interest_rows = []
    for m in df_g3b.columns:
        if m in ["Source File"]: 
            continue
        days_late = df_g3b.loc[df_g3b["Details / Field"].str.contains("days late", case=False, na=False), m].sum()
        cash_tax = df_g3b.loc[df_g3b["Details / Field"].str.contains("Cash", case=False, na=False), m].sum()
        calc_interest = round(cash_tax * 0.18 * (days_late / 365.0), 2) if days_late > 0 else 0.0
        interest_rows.append({
            "Period": m,
            "Days Delayed": days_late,
            "Net Cash Tax Paid": cash_tax,
            "Interest Payable (18%)": calc_interest,
            "Status": "Action Needed" if calc_interest > 0 else "Satisfactory"
        })
    output_sheets["10. Delayed Interest Audit"] = pd.DataFrame(interest_rows)

    # --- Summary Dashboard ---
    summary = {}
    if not df_g3b.empty and not df_g2b.empty:
        total_itc_availed = df_g3b.loc[df_g3b["Details / Field"].str.contains("IGST", case=False, na=False)].sum(numeric_only=True).sum()
        total_itc_available = df_g2b.loc[df_g2b["Details / Field"].str.contains("IGST", case=False, na=False)].sum(numeric_only=True).sum()
        summary["Total ITC Availed (3B)"] = total_itc_availed
        summary["Total ITC Available (2B)"] = total_itc_available
        summary["Deviation"] = total_itc_availed - total_itc_available

    if not df_g9.empty:
        tax_payable = df_g9["Table 4 - B2B Supplies"].sum()
        tax_paid = df_g9["Table 6A - Total ITC Availed"].sum() - df_g9["Table 7I - Total ITC Reversed"].sum()
        summary["Annual Tax Payable (GSTR-9)"] = tax_payable
        summary["Annual Tax Paid (Cash+ITC)"] = tax_paid
        summary["Annual Deviation"] = tax_payable - tax_paid

    if "1. ITC Supplier Recon" in output_sheets:
        risk_months = len(output_sheets["1. ITC Supplier Recon"].loc[
            output_sheets["1. ITC Supplier Recon"]["Risk Status"] == "Excess Availment Risk"
        ])
        summary["Excess ITC Risk Months"] = risk_months

    output_sheets["Summary Dashboard"] = pd.DataFrame(summary.items(), columns=["Metric", "Value"])

    # --- Save Final Combined Report ---
    print("\n📁 Saving unified sheets to Excel...")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df_data in output_sheets.items():
            if isinstance(df_data, pd.DataFrame) and not df_data.empty:
                df_data.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"   -> Included Sheet: '{sheet_name}'")

    print(f"\n🎉 SUCCESS! Unified Consolidated Audit workbook created at: {output_path}")
    return output_path

# Alias for Streamlit app
execute_unified_consolidation = execute_final_consolidation
