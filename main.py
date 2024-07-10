import fitz  # PyMuPDF
import pandas as pd
import argparse

def extract_tables_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    tables = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("dict")

        blocks = text['blocks']
        for block in blocks:
            if 'lines' in block:
                table_data = []
                skip_first_line = True  # Flag to skip the first line of each new table
                for line_idx, line in enumerate(block['lines']):
                    if skip_first_line:
                        skip_first_line = False
                        continue
                    
                    row_data = []
                    for span in line['spans']:
                        row_data.extend(span['text'].split())
                    
                    # Append row data to table data
                    table_data.append(row_data)

                if table_data:  # Only append non-empty tables
                    tables.append(table_data)

    # Process tables to add even row content as a new column
    processed_tables = []
    for table in tables:
        processed_table = []
        for i in range(len(table)):
            if i % 2 == 0 and i + 1 < len(table):
                processed_table.append(table[i] + table[i + 1])
        processed_tables.append(processed_table)

    return processed_tables

def save_tables_to_excel(tables, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for idx, table in enumerate(tables):
            df = pd.DataFrame(table)
            df.to_excel(writer, sheet_name=f'Table_{idx+1}', index=False, header=False)

def main(pdf_file, excel_file):
    tables = extract_tables_from_pdf(pdf_file)
    save_tables_to_excel(tables, excel_file)
    print(f'Extracted tables from {pdf_file} and saved to {excel_file}')

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract tables from a PDF and save them to an Excel file.")
    parser.add_argument("pdf_file", help="PDF file name")
    parser.add_argument("excel_file", help="Excel file name")

    args = parser.parse_args()

    main(args.pdf_file, args.excel_file)
