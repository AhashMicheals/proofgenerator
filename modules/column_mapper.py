from typing import Dict, List, Optional, Any
import re

PHOTO_ALIASES = ["photo", "photo_name", "photo name", "image", "image name", "picture", "photo path", "photo_filename", "pic", "img"]
ID_ALIASES = ["rf id no", "rf_id_no", "rf id", "rf_id", "rfid", "id no", "id_no", "id", "id number", "reg no", "emp id", "employee id", "staff id", "student id", "roll no", "code"]
TITLE_ALIASES = ["name", "staff name", "staff_name", "student name", "person name", "full name", "employee name", "title"]

def detect_key_columns(columns: List[str]) -> Dict[str, Optional[str]]:
    """
    Detects potential key columns from arbitrary Excel headers:
    - photo_col: Column containing photo filenames or images
    - id_col: Primary ID/registration number column
    - title_col: Primary Name/Title column
    """
    normalized_cols = {str(col).strip(): str(col).strip().lower() for col in columns}
    
    photo_col = None
    id_col = None
    title_col = None
    
    # 1. Detect Photo column
    for col, norm_col in normalized_cols.items():
        clean_norm = re.sub(r'[^a-z0-9]', '', norm_col)
        for alias in PHOTO_ALIASES:
            clean_alias = re.sub(r'[^a-z0-9]', '', alias)
            if clean_alias == clean_norm or clean_alias in clean_norm:
                photo_col = col
                break
        if photo_col:
            break
            
    # 2. Detect ID column
    for col, norm_col in normalized_cols.items():
        clean_norm = re.sub(r'[^a-z0-9]', '', norm_col)
        for alias in ID_ALIASES:
            clean_alias = re.sub(r'[^a-z0-9]', '', alias)
            if clean_alias == clean_norm or clean_alias in clean_norm:
                id_col = col
                break
        if id_col:
            break
            
    # 3. Detect Title/Name column
    for col, norm_col in normalized_cols.items():
        clean_norm = re.sub(r'[^a-z0-9]', '', norm_col)
        for alias in TITLE_ALIASES:
            clean_alias = re.sub(r'[^a-z0-9]', '', alias)
            if clean_alias == clean_norm or clean_alias in clean_norm:
                title_col = col
                break
        if title_col:
            break
            
    return {
        "photo_col": photo_col,
        "id_col": id_col,
        "title_col": title_col
    }

