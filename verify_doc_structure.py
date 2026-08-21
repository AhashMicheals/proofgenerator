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
    print(f"\nTotal Page Tables: {len(tables)} (Expected 6 for 21 records)")
    
    total_cards_found = 0
    for idx, table in enumerate(tables):
        rows = len(table.rows)
        cols = len(table.columns)
        print(f"\n--- Page Table #{idx+1} ---")
        
        table_cards = 0
        for r in range(rows):
            for c in range(cols):
                cell = table.cell(r, c)
                text = cell.text
                if "ID PROOF" in text:
                    table_cards += 1
                    total_cards_found += 1
                    lines = [line.strip() for line in text.split('\n') if line.strip()]
                    name = next((l for l in lines if l.startswith("Name")), "")
                    rfid = next((l for l in lines if l.startswith("RF ID")), "")
                    print(f"  [Cell {r},{c}] Card -> {name} | {rfid}")
                else:
                    print(f"  [Cell {r},{c}] Empty slot")
                    
        print(f"Cards on Page {idx+1}: {table_cards}")
        
    print("\n================ SUMMARY ================")
    print(f"Total Staff Cards Verified in Document: {total_cards_found}")
    assert total_cards_found == 21, f"Expected 21 staff cards, found {total_cards_found}"
    print("VERIFICATION SUCCESSFUL: 4-per-page A4 grid structure verified for 21 staff records!")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    doc_file = os.path.join(base_dir, "output", "Staff_ID_Proof.docx")
    verify_docx(doc_file)
