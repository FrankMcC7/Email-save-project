import os
import re
import pandas as pd
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import glob
from datetime import datetime

def extract_fund_name(text):
    """
    Extract the fund name from the PDF text.
    This function assumes the fund name typically appears early in the document.
    You may need to adjust the regex pattern based on the actual PDF structure.
    """
    # Pattern to match common fund name formats
    # This is a generic pattern and might need adjustment based on your specific PDFs
    patterns = [
        r"(?:Fund|FUND)[\s:]+([A-Za-z0-9\s\-\.&]+(?:Fund|FUND))",
        r"([A-Za-z0-9\s\-\.&]+(?:Fund|FUND))",
        r"Report for\s+([A-Za-z0-9\s\-\.&]+)",
        r"([A-Za-z0-9\s\-\.&]+)\s+Performance Report"
    ]
    
    for pattern in patterns:
        matches = re.search(pattern, text[:1000])  # Look only in first 1000 chars
        if matches:
            return matches.group(1).strip()
    
    return "Unknown Fund"  # Default if no match found

def extract_nav_values(text):
    """
    Extract Net Asset Value figures for both periods from the Key Figure table.
    This function looks for "Net Asset Value" and extracts the associated values.
    """
    # First find the Key Figure table section
    table_pattern = r"Key (?:Figure|Figures|FIGURE|FIGURES).*?(?:Table|TABLE)"
    table_match = re.search(table_pattern, text, re.DOTALL | re.IGNORECASE)
    
    if not table_match:
        # If we can't find a specific table header, search for Net Asset Value directly
        nav_pattern = r"Net\s+Asset\s+Value.*?(?:[\d,\.]+)"
        section_text = text
    else:
        # If we found the table section, limit our search to that area plus some context
        start_pos = max(0, table_match.start() - 100)
        end_pos = min(len(text), table_match.end() + 1000)  # Look 1000 chars after table header
        section_text = text[start_pos:end_pos]
    
    # Find the Net Asset Value line
    nav_pattern = r"Net\s+Asset\s+Value[^\n\d]*([\d,\.]+)[^\n\d]*([\d,\.]+)"
    nav_match = re.search(nav_pattern, section_text, re.DOTALL | re.IGNORECASE)
    
    if nav_match:
        # Extract and clean the values
        period1_value = nav_match.group(1).replace(',', '')
        period2_value = nav_match.group(2).replace(',', '')
        
        # Convert to numeric values
        try:
            period1_value = float(period1_value)
            period2_value = float(period2_value)
            return period1_value, period2_value
        except ValueError:
            pass
    
    # Alternative approach with more flexible pattern
    nav_pattern2 = r"Net\s+Asset\s+Value.*?(\d[\d,\.]+).*?(\d[\d,\.]+)"
    nav_match2 = re.search(nav_pattern2, section_text, re.DOTALL | re.IGNORECASE)
    
    if nav_match2:
        # Extract and clean the values
        period1_value = nav_match2.group(1).replace(',', '')
        period2_value = nav_match2.group(2).replace(',', '')
        
        # Convert to numeric values
        try:
            period1_value = float(period1_value)
            period2_value = float(period2_value)
            return period1_value, period2_value
        except ValueError:
            pass
    
    return None, None  # Return None if we couldn't find the values

def extract_period_labels(text):
    """
    Extract the period labels from the document.
    This function assumes period labels would be near the NAV values or in the header.
    """
    # Pattern to match dates in various formats
    date_pattern = r"(?:(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s,.-]+\d{1,2}[\s,.-]+\d{2,4}|\d{1,2}[\s,.-]+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s,.-]+\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2})"
    
    # Find dates in the document
    dates = re.findall(date_pattern, text, re.IGNORECASE)
    
    # If we found at least two dates, use those as period labels
    if len(dates) >= 2:
        return dates[0], dates[1]
    
    # Fallback to quarters or years if specific dates aren't found
    quarter_pattern = r"(?:Q[1-4]\s+\d{4}|[1-4](?:st|nd|rd|th)\s+Quarter\s+\d{4})"
    quarters = re.findall(quarter_pattern, text, re.IGNORECASE)
    
    if len(quarters) >= 2:
        return quarters[0], quarters[1]
    
    # Last resort: look for years
    year_pattern = r"\b(20\d{2})\b"
    years = re.findall(year_pattern, text)
    
    if len(years) >= 2:
        return years[0], years[1]
    
    # Default labels if we can't find anything specific
    return "Period 1", "Period 2"

def process_pdf(pdf_path):
    """
    Process a single PDF file to extract fund name and NAV values.
    """
    try:
        # Extract text from PDF
        all_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                all_text += page.extract_text() + "\n"
        
        # Extract fund name
        fund_name = extract_fund_name(all_text)
        
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
        
        # Center numeric values
        for row in ws.iter_rows(min_row=2):
            for cell in row[2:]:  # Apply only to NAV columns
                if isinstance(cell.value, (int, float)):
                    cell.alignment = Alignment(horizontal='center')
        
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
