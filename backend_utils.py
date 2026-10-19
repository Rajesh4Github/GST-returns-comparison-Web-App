import re
import pandas as pd

def clean_val(val):
    if val is None or pd.isna(val):
        return 0.0
    val_str = str(val).replace(",", "").replace(" ", "").replace("\n", "").strip()
    if val_str in ["", "-", ".", "0.00"]:
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def get_month_key_from_filename(filename):
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
