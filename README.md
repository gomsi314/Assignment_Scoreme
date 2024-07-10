# PDF Table Extraction and Excel Export

## Objective

The objective of this project is to develop a tool that can detect and accurately extract tables from system-generated PDFs. The extracted tables should be stored in Excel sheets while adhering to the structure and content integrity of the original tables. Notably, the solution must achieve this without relying on Tabula, Camelot, or converting PDFs into images.

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
