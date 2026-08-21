import pandas as pd
from typing import Dict, List, Any, Tuple, Optional
from modules.photo_matcher import match_photo_for_staff, process_and_crop_photo, create_placeholder_photo
from modules.data_cleaner import format_display_value

def validate_staff_records(
    df: pd.DataFrame, 
    photo_registry: Dict[str, Any],
    key_cols: Optional[Dict[str, Optional[str]]] = None
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Validates dynamic Excel records, matches photos, and compiles summary metrics.
    Does NOT depend on fixed field names.
    """
    processed_records = []
    
    key_cols = key_cols or {}
    id_col = key_cols.get("id_col")
    photo_col = key_cols.get("photo_col")
    
    photos_found_count = 0
    photos_missing_count = 0
    missing_id_count = 0
    
    id_counts: Dict[str, int] = {}
    
    # Track ID duplicates if an ID column exists
    if id_col and id_col in df.columns:
        for idx, row in df.iterrows():
            id_val = str(row.get(id_col, "")).strip()
            if id_val and id_val != "-":
                id_counts[id_val] = id_counts.get(id_val, 0) + 1
                
    duplicate_ids = [id_v for id_v, count in id_counts.items() if count > 1]
    
    # Process each record
    for idx, row in df.iterrows():
        # Store original data dict
        record = {}
        for col in df.columns:
            record[col] = str(row[col]).strip() if pd.notna(row[col]) else ""
            
        record["_row_index"] = idx + 1
        
        # Check missing ID if ID column exists
        if id_col and id_col in df.columns:
            if not record.get(id_col):
                missing_id_count += 1
                
        # Match photo
        photo_match = match_photo_for_staff(record, photo_registry, key_cols)
        
        if photo_match:
            photo_status = "✓ Found"
            matched_filename = photo_match["filename"]
            photo_bytes = process_and_crop_photo(photo_match["bytes"])
            is_photo_found = True
            photos_found_count += 1
        else:
            photo_status = "✗ Missing"
            matched_filename = None
            photo_bytes = create_placeholder_photo()
            is_photo_found = False
            photos_missing_count += 1
            
        record["_photo_status"] = photo_status
        record["_matched_filename"] = matched_filename
        record["_photo_bytes"] = photo_bytes
        record["_is_photo_found"] = is_photo_found
        record["_key_cols"] = key_cols
        
        # Also store backward compatible display fields for legacy standard keys if present
        for col in df.columns:
            record[f"display_{col}"] = format_display_value(record[col])
            
        processed_records.append(record)

    # Build warning messages
    warnings_list = []
    if missing_id_count > 0 and id_col:
        warnings_list.append(f"⚠ {missing_id_count} record(s) have missing values in ID column '{id_col}'")
    if photos_missing_count > 0:
        warnings_list.append(f"⚠ {photos_missing_count} record(s) have missing photos (will use 'PHOTO NOT FOUND' placeholder)")
    if duplicate_ids and id_col:
        warnings_list.append(f"⚠ Duplicate IDs found in '{id_col}': {', '.join(duplicate_ids[:5])}{'...' if len(duplicate_ids) > 5 else ''}")
        
    summary = {
        "total_records": len(processed_records),
        "photos_found": photos_found_count,
        "photos_missing": photos_missing_count,
        "records_missing_rfid": missing_id_count,
        "duplicate_rfids": duplicate_ids,
        "warnings": warnings_list
    }
    
    return processed_records, summary

