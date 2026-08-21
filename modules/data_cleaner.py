import re
import pandas as pd
from typing import Any, Dict

def clean_cell_value(val: Any) -> str:
    """
    Clean individual cell values to sanitized strings.
    Handles floats (e.g. 277735.0 -> '277735'), NaNs, datetimes, extra spaces.
    """
    if pd.isna(val) or val is None:
        return ""
        
    s_val = str(val).strip()
    
    if s_val.lower() in ("nan", "none", "blank", "<na>", "null"):
        return ""
        
    # Handle pandas Timestamp or datetime objects
    if isinstance(val, pd.Timestamp):
        return val.strftime("%d.%m.%Y")
        
    # Handle float representations like 277735.0 or 9176750670.0
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
        
    # Regex fix for numeric strings ending in .0
    if re.match(r'^\d+\.0$', s_val):
        s_val = s_val[:-2]
        
    return s_val

def combine_address(add1: str, add2: str, address: str = "") -> str:
    """
    Combines ADD1 and ADD2 or uses single ADDRESS field.
    """
    a1 = clean_cell_value(add1)
    a2 = clean_cell_value(add2)
    addr = clean_cell_value(address)
    
    parts = []
    if a1:
        parts.append(a1.rstrip(','))
    if a2:
        parts.append(a2)
        
    if parts:
        combined = ", ".join(parts)
        # Clean up double commas if any
        return re.sub(r',\s*,', ',', combined).strip()
        
    return addr

def format_display_value(val: Any) -> str:
    """
    Returns sanitized value or '-' if empty/blank for Word document & UI preview.
    """
    cleaned = clean_cell_value(val)
    return cleaned if cleaned else "-"
