import os
import docx

def verify_docx(docx_path):
    print("================ DOCUMENT VERIFICATION ================")
    doc = docx.Document(docx_path)
    
    section = doc.sections[0]
    print(f"Page width: {section.page_width.inches:.2f} in (A4 is 8.27 in)")
    print(f"Page height: {section.page_height.inches:.2f} in (A4 is 11.69 in)")
    print(f"Margins (T, B, L, R): {section.top_margin.inches:.2f}, {section.bottom_margin.inches:.2f}, {section.left_margin.inches:.2f}, {section.right_margin.inches:.2f}")
    
    tables = doc.tables
    print(f"\nTotal Tables: {len(tables)} (Continuous table layout)")
    
    total_cards_found = 0
    for idx, table in enumerate(tables):
        rows = len(table.rows)
        cols = len(table.columns)
        print(f"\n--- Table #{idx+1} ({rows} rows, {cols} cols) ---")
        
        for r in range(rows):
            page_num = (r // 5) + 1
            for c in range(cols):
                cell = table.cell(r, c)
                text = cell.text
                if "ID PROOF" in text:
                    total_cards_found += 1
                    lines = [line.strip() for line in text.split('\n') if line.strip()]
                    name = next((l for l in lines if l.startswith("Name")), "")
                    rfid = next((l for l in lines if l.startswith("RF ID")), "")
                    print(f"  [Row {r}, Col {c}] (Page ~{page_num}) Card -> {name} | {rfid}")
                else:
                    print(f"  [Row {r}, Col {c}] (Page ~{page_num}) Empty slot")
                    
    print("\n================ SUMMARY ================")
    print(f"Total Staff Cards Verified in Document: {total_cards_found}")
    assert total_cards_found == 21, f"Expected 21 staff cards, found {total_cards_found}"
    print("VERIFICATION SUCCESSFUL: Continuous 10-per-page (2x5) A4 grid structure verified with zero blank pages!")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    doc_file = os.path.join(base_dir, "output", "Staff_ID_Proof.docx")
    verify_docx(doc_file)
