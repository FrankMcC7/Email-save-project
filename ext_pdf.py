import os
import re
import pandas as pd
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import glob
from datetime import datetime

def extract_fund_name(pdf_path):
    """
    Extract the fund name from the third line of the first page.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) > 0:
                first_page_text = pdf.pages[0].extract_text()
                if first_page_text:
                    # Split text by newlines and get the third line if available
                    lines = first_page_text.split('\n')
                    if len(lines) >= 3:
                        return lines[2].strip()
                    elif len(lines) > 0:
                        # If third line isn't available, look for line with 'Fund' or 'SICAV' in first 5 lines
                        for i in range(min(5, len(lines))):
                            if 'Fund' in lines[i] or 'SICAV' in lines[i] or 'S.C.A' in lines[i]:
                                return lines[i].strip()
        
        # Fallback to extracting text from first few pages and looking for fund name patterns
        all_text = extract_text_from_pdf(pdf_path, max_pages=3)
        fund_patterns = [
            r"([A-Za-z0-9\s\-\.&]+(?:S\.C\.A\.|SICAV|Fund))",
            r"([A-Za-z]+\s+[A-Za-z]+\s+[A-Za-z]+(?:\s+[IVX]+)?)"
        ]
        
        for pattern in fund_patterns:
            match = re.search(pattern, all_text)
            if match:
                return match.group(1).strip()
    except Exception as e:
        print(f"Error extracting fund name: {str(e)}")
    
    return "Unknown Fund"

def extract_text_from_pdf(pdf_path, max_pages=None):
    """
    Extract text from PDF, optionally limiting to max_pages.
    """
    try:
        all_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            pages_to_extract = pdf.pages if max_pages is None else pdf.pages[:max_pages]
            for page in pages_to_extract:
                page_text = page.extract_text() or ""
                all_text += page_text + "\n"
        return all_text
    except Exception as e:
        print(f"Error extracting text from PDF: {str(e)}")
        return ""

def extract_nav_values(text):
    """
    Find 'Net Asset Value' and extract values from the surrounding table structure.
    """
    # First, try to find the exact "Net Asset Value" term
    nav_patterns = [
        r"Net\s+Asset\s+Value",
        r"Net\s+asset\s+value",
        r"NAV"
    ]
    
    # Search for the term using different patterns
    nav_section = None
    for pattern in nav_patterns:
        nav_match = re.search(pattern, text, re.IGNORECASE)
        if nav_match:
            # Extract a section of text around the match
            start_pos = max(0, nav_match.start() - 100)
            end_pos = min(len(text), nav_match.end() + 300)  # Look 300 chars after the term
            nav_section = text[start_pos:end_pos]
            break
    
    if not nav_section:
        return None, None
    
    # Now extract values from the NAV section
    # Look for patterns of numbers after the NAV term
    value_patterns = [
        # Pattern for values with apostrophes like 566'445'652
        r"Net\s+[Aa]sset\s+[Vv]alue.*?(\d[\d\s',\.]*)\s+(\d[\d\s',\.]*)",
        
        # Pattern for values in standard format
        r"Net\s+[Aa]sset\s+[Vv]alue.*?(\d[\d\s,\.]*)\s+(\d[\d\s,\.]*)",
        
        # More generic pattern for any numbers near NAV
        r"Net\s+[Aa]sset\s+[Vv]alue.*?(\d[\d,\.']+).*?(\d[\d,\.']+)"
    ]
    
    for pattern in value_patterns:
        values_match = re.search(pattern, nav_section, re.DOTALL)
        if values_match:
            try:
                # Clean the values (remove commas, apostrophes, spaces)
                period1_value = values_match.group(1).replace(',', '').replace('\'', '').replace(' ', '')
                period2_value = values_match.group(2).replace(',', '').replace('\'', '').replace(' ', '')
                
                # Convert to float
                period1_float = float(period1_value)
                period2_float = float(period2_value)
                
                return period1_float, period2_float
            except (ValueError, IndexError):
                continue
    
    # If table-based patterns fail, try to find any numbers near "Net Asset Value"
    number_pattern = r"\d[\d,\.\'']*"
    numbers = re.findall(number_pattern, nav_section)
    
    if len(numbers) >= 2:
        try:
            # Clean the values
            period1_value = numbers[0].replace(',', '').replace('\'', '').replace(' ', '')
            period2_value = numbers[1].replace(',', '').replace('\'', '').replace(' ', '')
            
            # Convert to float
            period1_float = float(period1_value)
            period2_float = float(period2_value)
            
            return period1_float, period2_float
        except ValueError:
            pass
    
    return None, None

def extract_period_labels(text):
    """
    Extract period labels by identifying dates near the NAV values.
    """
    # Look for date patterns in the text
    date_patterns = [
        # DD.MM.YYYY format (e.g., 30.09.2024)
        r"(\d{2}\.\d{2}\.\d{4})",
        
        # MM/DD/YYYY or DD/MM/YYYY format
        r"(\d{1,2}/\d{1,2}/\d{4})",
        
        # Quarter format (e.g., Q4 2024)
        r"(Q[1-4]\s+\d{4})",
        
        # Month and year format
        r"([A-Za-z]+\s+\d{4})"
    ]
    
    for pattern in date_patterns:
        dates = re.findall(pattern, text)
        if len(dates) >= 2:
            return dates[0], dates[1]
    
    # If no specific dates found, look for a report period
    period_pattern = r"(?:period|report).*?(\d{4}).*?(?:to|-).*?(\d{4})"
    period_match = re.search(period_pattern, text, re.IGNORECASE)
    
    if period_match:
        return f"Period {period_match.group(1)}", f"Period {period_match.group(2)}"
    
    # Default labels
    return "Period 1", "Period 2"

def process_pdf(pdf_path):
    """
    Process a single PDF file to extract fund name and NAV values.
    """
    try:
        # Extract fund name from the third line of the first page
        fund_name = extract_fund_name(pdf_path)
        
        # Extract full text from the PDF
        all_text = extract_text_from_pdf(pdf_path)
        
        # Extract NAV values
        period1_nav, period2_nav = extract_nav_values(all_text)
        
        # Extract period labels
        period1_label, period2_label = extract_period_labels(all_text)
        
        return {
            'Fund Name': fund_name,
            'Period 1 Label': period1_label,
            'Period 1 NAV': period1_nav,
            'Period 2 Label': period2_label,
            'Period 2 NAV': period2_nav,
            'PDF Filename': os.path.basename(pdf_path)
        }
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
        return {
            'Fund Name': f"Error: {os.path.basename(pdf_path)}",
            'Period 1 Label': 'Period 1',
            'Period 1 NAV': None,
            'Period 2 Label': 'Period 2',
            'Period 2 NAV': None,
            'PDF Filename': os.path.basename(pdf_path)
        }

def format_excel(excel_path):
    """
    Apply formatting to the Excel file for better readability.
    """
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        
        # Apply header formatting
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
        
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = openpyxl.utils.get_column_letter(column[0].column)
            
            for cell in column:
                if cell.value:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
            
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Format numeric values
        for row in ws.iter_rows(min_row=2):
            for cell in row[2:]:  # Apply only to NAV columns
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right')
        
        wb.save(excel_path)
        return True
    except Exception as e:
        print(f"Error formatting Excel file: {str(e)}")
        return False

def main():
    # Folder containing PDF files
    pdf_folder = input("Enter the path to the folder containing PDF files: ")
    
    # Check if the folder exists
    if not os.path.isdir(pdf_folder):
        print(f"Error: Folder '{pdf_folder}' does not exist.")
        return
    
    # Find all PDF files in the folder
    pdf_files = glob.glob(os.path.join(pdf_folder, "*.pdf"))
    
    if not pdf_files:
        print(f"No PDF files found in folder '{pdf_folder}'.")
        return
    
    print(f"Found {len(pdf_files)} PDF files. Processing...")
    
    # Process each PDF file
    results = []
    for pdf_path in pdf_files:
        print(f"Processing {os.path.basename(pdf_path)}...")
        result = process_pdf(pdf_path)
        results.append(result)
    
    # Create DataFrame from results
    df = pd.DataFrame(results)
    
    # Create output Excel file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(pdf_folder, f"Fund_NAV_Summary_{timestamp}.xlsx")
    
    # Export to Excel
    df.to_excel(output_path, index=False)
    
    # Format the Excel file
    format_excel(output_path)
    
    print(f"\nProcessing complete. Results saved to: {output_path}")

if __name__ == "__main__":
    main()
