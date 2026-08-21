import os
import pandas as pd
from modules.excel_reader import read_excel_data, get_sheet_names
from modules.photo_matcher import build_photo_registry
from modules.validator import validate_staff_records
from modules.word_generator import generate_word_document

def test_full_pipeline():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_dir, "sample_data", "staff_data.xlsx")
    photos_dir = os.path.join(base_dir, "sample_data", "photos")
    output_dir = os.path.join(base_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"Reading Excel sheets from {excel_path}...")
    sheets = get_sheet_names(excel_path)
    print(f"Detected sheets: {sheets}")
    
    clean_df, mapping, _, selected_sheet = read_excel_data(excel_path, sheet_name="STAFF")
    print(f"Sheet '{selected_sheet}' read successfully. Total rows: {len(clean_df)}")
    print(f"Detected column mapping: {mapping}")
    
    photo_files = []
    if os.path.exists(photos_dir):
        for f in os.listdir(photos_dir):
            full_p = os.path.join(photos_dir, f)
            if os.path.isfile(full_p):
                photo_files.append(full_p)
                
    print(f"Loaded {len(photo_files)} photo files.")
    registry = build_photo_registry(photo_files)
    
    processed_records, summary = validate_staff_records(clean_df, registry)
    print(f"Total Records: {summary['total_records']}, Photos Found: {summary['photos_found']}, Photos Missing: {summary['photos_missing']}")
    for w in summary['warnings']:
        print("Warning:", w.encode('ascii', 'replace').decode('ascii'))
    
    print("Generating Word Document...")
    docx_bytes = generate_word_document(processed_records)
    
    out_docx_path = os.path.join(output_dir, "Staff_ID_Proof.docx")
    with open(out_docx_path, "wb") as f:
        f.write(docx_bytes)
        
    print(f"SUCCESS: Word Document generated at: {out_docx_path}")
    print(f"File size: {len(docx_bytes)} bytes")

if __name__ == "__main__":
    test_full_pipeline()
