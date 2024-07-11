import pandas as pd
from tabula import read_pdf
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

# Function to extract tables from a PDF
def extract_tables_from_pdf(pdf_path, password=None):
    tables = read_pdf(pdf_path, pages='all', password=password, multiple_tables=True)
    return tables

# Function to combine all extracted tables into a single DataFrame
def combine_tables(tables):
    combined_df = pd.concat(tables, ignore_index=True)
    return combined_df

# Function to write DataFrame to Excel with formatting
def write_to_excel(df, excel_path):
    wb = Workbook()
    ws = wb.active
    
    # Write DataFrame to worksheet
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    
    # Formatting
    for cell in ws["1:1"]:  # First row (header)
        cell.font = Font(bold=True)
    
    wb.save(excel_path)

# Main function
def pdf_to_excel(pdf_path, excel_path, password=None):
    tables = extract_tables_from_pdf(pdf_path, password)
    combined_df = combine_tables(tables)
    write_to_excel(combined_df, excel_path)

# Example usage
pdf_path = 'path_to_your_pdf_file.pdf'
excel_path = 'output_excel_file.xlsx'
password = 'your_pdf_password'
pdf_to_excel(pdf_path, excel_path, password)
