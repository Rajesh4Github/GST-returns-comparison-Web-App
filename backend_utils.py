import re

def clean_value(val):
    """Convert raw cell values into floats safely."""
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
    """Extract month key from filename or PDF text."""
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
    fname_lower = filename.lower()
    for k,v in months_map.items():
        if k in fname_lower:
            return v
    text_lower = pdf_text.lower()
    for k,v in months_map.items():
        if k in text_lower:
            return v
    return "UNKNOWN"
