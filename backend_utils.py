import os, re

def clean_value(val):
    """Convert raw cell values into floats safely."""
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

def get_month_key(filename):
    """Extract month key from filename like _042024 → APR."""
    match = re.search(r"_(\d{2})\d{4}", filename)
    if match:
        m_num = match.group(1)
        months_map = {
            "04": "APR","05": "MAY","06": "JUN","07": "JUL",
            "08": "AUG","09": "SEP","10": "OCT","11": "NOV",
            "12": "DEC","01": "JAN","02": "FEB","03": "MAR"
        }
        return months_map.get(m_num, None)
    return None
