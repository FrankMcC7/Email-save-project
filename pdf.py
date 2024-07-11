import fitz  # PyMuPDF
from PIL import Image
import easyocr
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
import io

# Function to convert PDF page to image
def pdf_page_to_image(doc, page_num, zoom=2):
    page = doc.load_page(page_num)
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img = Image.open(io.BytesIO(pix.tobytes()))
    return img

# Function to extract text from an image using EasyOCR
def extract_text_from_image(image, reader):
    image_bytes = io.BytesIO()
    image.save(image_bytes, format='PNG')
    image_bytes.seek(0)
    text = reader.readtext(image_bytes, detail=0, paragraph=True)
    return ' '.join(text)

# Function to extract text from PDF pages
def extract_text_from_pdf(pdf_path, password=None):
    reader = easyocr.Reader(['en'])  # Initialize EasyOCR reader with English language
    doc = fitz.open(pdf_path)
    
    if doc.needs_pass:
        if not password:
            raise ValueError("Password is required for this PDF")
        doc.authenticate(password)
    
    data = []
    for page_num in range(len(doc)):
        img = pdf_page_to_image(doc, page_num)
        text = extract_text_from_image(img, reader)
        data.append(text)
    
    return data

# Function to parse the text into a structured format (table)
def parse_text_to_table(data):
    structured_data = []
    for page in data:
        lines = page.split('\n')
        for line in lines:
            # Assuming each line is a separate row, further parsing might be required based on actual data
            structured_data.append(line.split())
    
    return structured_data

# Function to convert structured data to DataFrame
def create_dataframe(structured_data):
    df = pd.DataFrame(structured_data)
    return df

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
    raw_data = extract_text_from_pdf(pdf_path, password)
    structured_data = parse_text_to_table(raw_data)
    df = create_dataframe(structured_data)
    write_to_excel(df, excel_path)

# Example usage
pdf_path = 'path_to_your_pdf_file.pdf'
excel_path = 'output_excel_file.xlsx'
password = 'your_pdf_password'
pdf_to_excel(pdf_path, excel_path, password)
