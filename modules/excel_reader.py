import pandas as pd
from typing import Dict, List, Tuple, Any, Optional
from modules.column_mapper import detect_key_columns
from modules.data_cleaner import clean_cell_value, combine_address

def get_sheet_names(file_input: Any) -> List[str]:
    """
    Returns list of sheet names from uploaded Excel file.
    """
    try:
        if hasattr(file_input, "seek"):
            file_input.seek(0)
            
        filename = getattr(file_input, "name", "")
        if filename.endswith(".xls"):
            xl = pd.ExcelFile(file_input, engine="xlrd")
        else:
            xl = pd.ExcelFile(file_input, engine="openpyxl")
            
        return xl.sheet_names
    except Exception:
        try:
            if hasattr(file_input, "seek"):
                file_input.seek(0)
            xl = pd.ExcelFile(file_input)
            return xl.sheet_names
        except Exception:
            return []

def read_excel_data(
    file_input: Any, 
    sheet_name: Optional[str] = None, 
    key_overrides: Optional[Dict[str, Optional[str]]] = None
) -> Tuple[pd.DataFrame, Dict[str, Optional[str]], List[str], str]:
    """
    Reads Excel sheet, dynamically cleans all columns, and returns clean DataFrame & key column role mappings.
    Does NOT restrict data to hardcoded columns.
    """
    if hasattr(file_input, "seek"):
        file_input.seek(0)
        
    sheets = get_sheet_names(file_input)
    selected_sheet = sheet_name if sheet_name in sheets else (sheets[0] if sheets else "Sheet1")
    
    try:
        if hasattr(file_input, "seek"):
            file_input.seek(0)
            
        filename = getattr(file_input, "name", "")
        if filename.endswith(".xls"):
            df = pd.read_excel(file_input, sheet_name=selected_sheet, engine="xlrd")
        else:
            df = pd.read_excel(file_input, sheet_name=selected_sheet, engine="openpyxl")
    except Exception as e:
        try:
            if hasattr(file_input, "seek"):
                file_input.seek(0)
            df = pd.read_excel(file_input, sheet_name=selected_sheet)
        except Exception as ex:
            return pd.DataFrame(), {}, sheets, selected_sheet

    # Drop completely blank rows
    df = df.dropna(how="all").reset_index(drop=True)
    
    if df.empty:
        return pd.DataFrame(), {}, sheets, selected_sheet

    # Clean header names (strip spaces)
    df.columns = [str(col).strip() for col in df.columns]

    # Dynamically clean cell values across all columns
    clean_df = pd.DataFrame()
    for col in df.columns:
        clean_df[col] = df[col].apply(clean_cell_value)
        
    # Auto-combine ADD1 + ADD2 into ADDRESS if ADD1 & ADD2 exist in Excel but combined ADDRESS doesn't
    cols_upper = [c.upper() for c in clean_df.columns]
    if "ADD1" in cols_upper and "ADD2" in cols_upper and "ADDRESS" not in cols_upper:
        add1_col = clean_df.columns[cols_upper.index("ADD1")]
        add2_col = clean_df.columns[cols_upper.index("ADD2")]
        combined = []
        for idx, row in clean_df.iterrows():
            combined.append(combine_address(row[add1_col], row[add2_col], ""))
        clean_df["ADDRESS"] = combined

    # Detect key columns (photo, id, title)
    key_cols = detect_key_columns(clean_df.columns.tolist())
    
    if key_overrides:
        for k, v in key_overrides.items():
            if v and v != "-- Auto-Detect --" and v != "-- Select --":
                key_cols[k] = v

    return clean_df, key_cols, sheets, selected_sheet

