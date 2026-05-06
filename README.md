# PDF Table Extractor

A standalone GUI application for extracting table data and text from PDF files and exporting to **Excel**, **Word**, or **plain-text** format.

---

## Features

| Feature | Details |
|---|---|
| Upload any PDF | Browse to select a `.pdf` file |
| Choose page range | Pick the exact "from" and "to" page numbers |
| Three output formats | Excel (`.xlsx`), Word (`.docx`), or plain text (`.txt`) |
| Progress bar | Shows per-page progress during extraction |
| Parallel processing | Up to 4 worker threads extract pages simultaneously |
| Same-folder output | The output file is saved next to the original PDF |
| Standalone `.exe` | No Python installation needed on the end-user machine |

---

## Quick Start (run from source)

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Launch the application
python pdf_extractor.py
```

---

## Build a Standalone Executable

> **Requirements:** Python 3.10+, PyInstaller (installed automatically by the build script)

```bash
# On Windows / macOS / Linux
python build_exe.py
```

The finished executable will be in the `dist/` folder:

```
dist/
  PDFTableExtractor.exe   ← Windows
  PDFTableExtractor       ← macOS / Linux
```

No Python installation is required on the machine that runs the executable.

---

## Usage

1. **Select PDF File** – click *Browse…* and choose your PDF.  
   The total page count is displayed next to the file path.
2. **Page Range** – set *From page* and *To page* (e.g. `2` → `39`).
3. **Output Format** – pick Excel, Word, or Text.
4. **Extract** – click the *Extract* button.  
   The progress bar updates as each page is processed.
5. The output file is saved in the **same folder as the PDF**, named:  
   `<original_name>_p<from>-<to>.<ext>`  
   e.g. `report_p2-39.xlsx`

---

## Running Tests

```bash
pip install pytest
pytest test_pdf_extractor.py -v
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `pdfplumber` | PDF table and text extraction |
| `openpyxl` | Excel file writing |
| `python-docx` | Word document writing |
| `pyinstaller` | Build standalone executable |
