# PDF Table Extraction and Excel Export

This Python script extracts tables from a PDF file using PyMuPDF (fitz) and saves each table into separate sheets in an Excel workbook.

## Requirements
- Python 3.x
- PyMuPDF (`fitz`)
- pandas

Install dependencies using:
```bash
pip install PyMuPDF pandas openpyxl
```
To run the script:
```bash
python main.py <path_to_pdf_file> <output_excel_file.xlsx>
