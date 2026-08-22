import sys
import os
import io

# Set Streamlit 2GB max upload size programmatically
os.environ["STREAMLIT_SERVER_MAX_UPLOAD_SIZE"] = "2000"
os.environ["STREAMLIT_SERVER_MAX_MESSAGE_SIZE"] = "2000"

# Ensure local workspace root is prioritized in sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import streamlit as st

import pandas as pd

from modules.excel_reader import read_excel_data, get_sheet_names
from modules.column_mapper import detect_key_columns
from modules.photo_matcher import build_photo_registry
from modules.validator import validate_staff_records
from modules.word_generator import generate_word_document

# Page Configuration
st.set_page_config(
    page_title="iVEEem ID Proof Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1F497D;
        margin-bottom: 0.1rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #555555;
        margin-bottom: 1.5rem;
    }
    .info-card {
        background-color: #EBF3FA;
        border-radius: 8px;
        padding: 12px 18px;
        border-left: 5px solid #1F497D;
        margin-bottom: 15px;
    }
    .stButton button {
        background-color: #1F497D;
        color: white;
        font-weight: 600;
        font-size: 1rem;
        padding: 0.6rem 2rem;
        border-radius: 6px;
        border: none;
        width: 100%;
    }
    .stButton button:hover {
        background-color: #153256;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

def main():
    st.markdown('<div class="main-header">iVEEem ID Proof Generator</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Upload ANY Excel file — process & generate 2×5 grid (10-per-page) Word proof documents automatically</div>', unsafe_allow_html=True)

    # Sidebar Data Sources
    st.sidebar.header("📁 Data Inputs")
    
    excel_file = st.sidebar.file_uploader(
        "Upload Any Excel File", 
        type=["xlsx", "xls"],
        help="Upload .xlsx or .xls file containing any records"
    )
    
    photo_files = st.sidebar.file_uploader(
        "Upload Photos (Images or ZIP)", 
        type=["jpg", "jpeg", "png", "webp", "bmp", "zip"],
        accept_multiple_files=True,
        help="Upload multiple images or ZIP archives (up to 2GB total)"
    )
    
    server_photo_folder = st.sidebar.text_input(
        "Or Server Photo Folder Path",
        value="",
        placeholder="e.g. C:/Photos or ./photos",
        help="Enter local or server folder path containing photos"
    )
    
    items_to_process = list(photo_files or [])
    if server_photo_folder:
        if os.path.exists(server_photo_folder):
            items_to_process.append(server_photo_folder)
        else:
            st.sidebar.warning(f"Folder not found: `{server_photo_folder}`")
        
    photo_registry = {}
    if items_to_process:
        photo_registry = build_photo_registry(items_to_process)
        unique_photos = len(set(v["filename"] for v in photo_registry.values()))
        st.sidebar.success(f"✓ Indexed {unique_photos} photos in registry!")

        
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 💡 Quick Workflow")
    st.sidebar.markdown("""
    1. Upload any Excel file & select sheet.
    2. Review detected columns (all columns included automatically).
    3. Upload staff photos or a **.ZIP archive** containing 1000s of photos.
    4. Review preview table & validation metrics.
    5. Click **Generate Word Proof Document**.
    6. Download `Staff_ID_Proof.docx`.
    """)

    if not excel_file:
        st.info("👋 **Welcome!** Upload any Excel file from the sidebar to generate ID proof cards automatically.")
        
        # Test Data Generator Helper
        if st.checkbox("Generate sample data files for testing"):
            try:
                from create_sample_data import generate_sample_data
                generate_sample_data()
                st.success("Sample data generated in `sample_data/`!")
                st.markdown("Sample Excel file: `sample_data/staff_data.xlsx` (Sheet: `STAFF`) with 21 records and photos in `sample_data/photos/`.")
            except Exception as e:
                st.error(f"Error generating sample data: {e}")
        return

    # 1. Sheet Selection
    sheets = get_sheet_names(excel_file)
    if not sheets:
        st.error("Could not read sheets from uploaded Excel file.")
        return
        
    selected_sheet = sheets[0]
    if len(sheets) > 1:
        selected_sheet = st.selectbox("Select Excel Sheet", options=sheets, index=0)
        
    # Read sheet raw data
    excel_file.seek(0)
    clean_df, detected_key_cols, _, _ = read_excel_data(excel_file, selected_sheet)
    
    if clean_df.empty:
        st.error("No valid record rows found in selected sheet.")
        return

    raw_columns = clean_df.columns.tolist()
    total_raw_records = len(clean_df)
    
    st.markdown(f"""
    <div class="info-card">
        <b>File:</b> <code>{excel_file.name}</code> &nbsp;|&nbsp; 
        <b>Sheet:</b> <code>{selected_sheet}</code> &nbsp;|&nbsp; 
        <b>Records:</b> <code>{total_raw_records}</code> &nbsp;|&nbsp; 
        <b>Detected Columns:</b> <code>{", ".join(raw_columns)}</code>
    </div>
    """, unsafe_allow_html=True)

    # 2. Key Role Overrides (Optional)
    with st.expander("⚙️ Key Column Settings (Auto-detected)", expanded=False):
        st.write("The app automatically processes all columns. You can specify key columns below for enhanced photo matching:")
        col1, col2, col3 = st.columns(3)
        
        options = ["-- Auto-Detect --"] + raw_columns
        
        det_photo = detected_key_cols.get("photo_col")
        det_id = detected_key_cols.get("id_col")
        det_title = detected_key_cols.get("title_col")
        
        idx_photo = options.index(det_photo) if det_photo in options else 0
        idx_id = options.index(det_id) if det_id in options else 0
        idx_title = options.index(det_title) if det_title in options else 0
        
        with col1:
            sel_photo = st.selectbox("Photo Filename Column", options=options, index=idx_photo, key="sel_photo")
        with col2:
            sel_id = st.selectbox("Primary ID Column", options=options, index=idx_id, key="sel_id")
        with col3:
            sel_title = st.selectbox("Primary Title / Name Column", options=options, index=idx_title, key="sel_title")

        key_overrides = {
            "photo_col": sel_photo if sel_photo != "-- Auto-Detect --" else det_photo,
            "id_col": sel_id if sel_id != "-- Auto-Detect --" else det_id,
            "title_col": sel_title if sel_title != "-- Auto-Detect --" else det_title
        }

    # 3. Validate Records & Match Photos
    if not photo_registry and items_to_process:
        photo_registry = build_photo_registry(items_to_process)

        
    processed_records, summary = validate_staff_records(clean_df, photo_registry, key_overrides)

    
    # 4. Display Validation Report & Metrics
    st.subheader("1. Data Summary & Metrics")
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Records", summary["total_records"])
    m2.metric("Photos Found", summary["photos_found"])
    m3.metric("Photos Missing", summary["photos_missing"], delta=None if summary["photos_missing"] == 0 else f"-{summary['photos_missing']}", delta_color="inverse")

    if summary["warnings"]:
        with st.expander("⚠️ Validation & Matching Warnings", expanded=True):
            for warn in summary["warnings"]:
                st.warning(warn)

    # 5. Dynamic Data Preview Table
    st.subheader("2. Data Preview")
    
    preview_df = clean_df.copy()
    preview_df.insert(0, "PHOTO STATUS", [r["_photo_status"] for r in processed_records])
    st.dataframe(preview_df, height=290)


    # 6. Word Document Generation & Download
    st.subheader("3. Word Document Generation")
    
    gen_btn = st.button("Generate Word Proof Document", key="gen_btn")
    
    if gen_btn or st.session_state.get("docx_generated"):
        # Generate a Word document with 10 proof cards per A4 page in a 2×5 grid
        with st.spinner("Generating 2×5 grid (10 cards per A4 page) Word Proof Document..."):
            docx_bytes = generate_word_document(processed_records)
            st.session_state["docx_bytes"] = docx_bytes
            st.session_state["docx_generated"] = True

        pages_count = (len(processed_records) + 9) // 10
        st.success("✓ Word document generated successfully!")
        st.info(f"**Total Records:** {len(processed_records)} | **Pages Created:** {pages_count} (10 proof cards per A4 page in 2×5 grid)")

        st.download_button(
            label="Download Staff_ID_Proof.docx",
            data=st.session_state["docx_bytes"],
            file_name="Staff_ID_Proof.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_btn"
        )

if __name__ == "__main__":
    main()
