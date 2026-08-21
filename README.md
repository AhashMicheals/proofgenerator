# Staff ID Proof Generator

A Python Streamlit web application to convert staff/student Excel data and photos into a professional Word document (`Staff_ID_Proof.docx`) formatted with **exactly 4 proof cards per A4 page** in a 2×2 grid layout.

## Features

- **Multi-Sheet Excel Support**: Supports `.xlsx` and `.xls` files, allowing sheet selection (e.g. `STAFF`).
- **Auto-Detection & Column Mapping**: Maps standard fields (`PHOTO`, `NAME`, `RF ID NO`, `DESIGNATION`, `DOB`, `FATHER`, `MOBILE`, `ADD1`, `ADD2`, `ADDRESS`, `BG`) with manual UI override selectors.
- **Address & Data Cleaning**: Combines `ADD1` + `ADD2` into unified address strings, strips float `.0` suffixes (e.g. `277735.0` -> `277735`), formats dates, and replaces empty/NaN cells with `-`.
- **4-Tier Photo Matching**: Matches photo files (`.jpg`, `.jpeg`, `.png`, `.webp`) by PHOTO name, RF ID NO, Staff NAME, or normalized filename slug.
- **Fallback Placeholder**: Renders a clear "PHOTO NOT FOUND" card for missing photos.
- **Validation Report**: Summary metrics (Total Records, Photos Found, Photos Missing, Missing RF IDs, Missing DOB, Duplicate RF IDs) and student preview table.
- **A4 2×2 Print Layout**: Uses `python-docx` to format cards inside a 2×2 page grid with exact margins and non-splitting row rules.

## Project Structure

```text
c:/Users/Mike/Desktop/EXCEL to DOC/
├── app.py
├── create_sample_data.py
├── test_generation.py
├── verify_doc_structure.py
├── requirements.txt
├── README.md
├── modules/
│   ├── column_mapper.py
│   ├── data_cleaner.py
│   ├── excel_reader.py
│   ├── photo_matcher.py
│   ├── validator.py
│   └── word_generator.py
├── sample_data/
└── output/
```

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Generate Sample Data

```bash
python create_sample_data.py
```

### 3. Run Streamlit App

```bash
python -m streamlit run app.py
```

Open your browser at `http://localhost:8501`.
