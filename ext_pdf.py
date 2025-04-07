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

def extract_text_from_pdf(pdf_path, max_pages=None, start_page=0):
    """
    Extract text from PDF, optionally limiting to max_pages starting from start_page.
    """
    try:
        all_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            # Determine which pages to extract
            if max_pages is None:
                pages_to_extract = pdf.pages[start_page:]
            else:
                end_page = min(start_page + max_pages, len(pdf.pages))
                pages_to_extract = pdf.pages[start_page:end_page]
                
            for page in pages_to_extract:
                page_text = page.extract_text() or ""
                all_text += page_text + "\n"
        return all_text
    except Exception as e:
        print(f"Error extracting text from PDF: {str(e)}")
        return ""

def extract_tables_from_pdf(pdf_path, start_page=1):
    """
    Extract tables from PDF starting from page 2 (index 1).
    Returns a list of tables, where each table is a list of rows.
    """
    tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num in range(start_page, len(pdf.pages)):
                page = pdf.pages[page_num]
                page_tables = page.extract_tables()
                if page_tables:
                    tables.extend(page_tables)
                    
                    # If we found tables, also try to extract the page text for context
                    page_text = page.extract_text() or ""
                    tables.append([["PAGE_TEXT"], [page_text]])
    except Exception as e:
        print(f"Error extracting tables from PDF: {str(e)}")
    
    return tables

def find_nav_in_tables(tables):
    """
    Find the Net Asset Value in the extracted tables.
    Returns tuple of (period1_label, period1_value, period2_label, period2_value).
    """
    nav_row = None
    header_row = None
    
    # First, find the table containing "Net Asset Value" or similar
    for table in tables:
        if not table:
            continue
            
        # Skip tables that are just page text containers
        if len(table) > 0 and len(table[0]) > 0 and table[0][0] == "PAGE_TEXT":
            continue
            
        for row_idx, row in enumerate(table):
            # Convert all items in row to strings for searching
            str_row = [str(item).strip() if item is not None else "" for item in row]
            row_text = " ".join(str_row).lower()
            
            # Check if this is a header row with dates or period labels
            if (any(date_indicator in row_text for date_indicator in ['.20', '/20', 'q1', 'q2', 'q3', 'q4']) or
                any(month in row_text for month in ['january', 'february', 'march', 'april', 'may', 'june', 'july', 
                                                   'august', 'september', 'october', 'november', 'december'])):
                header_row = row
            
            # Look for Net Asset Value in this row
            if 'net asset value' in row_text or 'net asset' in row_text or ('nav' in row_text and len(row_text) < 20):
                nav_row = row
                
                # If we found NAV row but no header yet, look at rows above
                if header_row is None and row_idx > 0:
                    for i in range(row_idx-1, -1, -1):
                        potential_header = table[i]
                        header_text = " ".join([str(item).strip() if item is not None else "" for item in potential_header]).lower()
                        if (any(date_indicator in header_text for date_indicator in ['.20', '/20', 'q1', 'q2', 'q3', 'q4']) or
                            any(month in header_text for month in ['january', 'february', 'march', 'april', 'may', 'june', 'july', 
                                                                  'august', 'september', 'october', 'november', 'december'])):
                            header_row = potential_header
                            break
                
                if header_row is not None:
                    # We found both NAV row and header - extract values
                    return extract_values_from_rows(header_row, nav_row)
    
    # If we couldn't find a clear table structure, try a different approach
    for table in tables:
        if not table:
            continue
            
        # Skip tables that are just page text containers
        if len(table) > 0 and len(table[0]) > 0 and table[0][0] == "PAGE_TEXT":
            page_text = table[1][0]
            # Use regex to find NAV values in the page text
            nav_match = re.search(r"[Nn]et\s+[Aa]sset\s+[Vv]alue.*?(\d[\d\s',\.]*)\s+(\d[\d\s',\.]*)", page_text)
            if nav_match:
                # Also look for column headers/dates in various formats
                
                # Standard numeric date format (DD.MM.YYYY or DD/MM/YYYY)
                date_match_numeric = re.findall(r"(\d{2}[\.\/]\d{2}[\.\/]\d{4})", page_text)
                
                # Month name format (DD. Month YYYY or DD Month YYYY)
                date_match_text = re.findall(r"(\d{1,2}\.?\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})", page_text, re.IGNORECASE)
                
                # Combine all found dates
                all_dates = date_match_numeric + date_match_text
                
                if len(all_dates) >= 2:
                    return all_dates[0], clean_and_convert_value(nav_match.group(1)), all_dates[1], clean_and_convert_value(nav_match.group(2))
                else:
                    return "Period 1", clean_and_convert_value(nav_match.group(1)), "Period 2", clean_and_convert_value(nav_match.group(2))
    
    return None, None, None, None

def extract_values_from_rows(header_row, nav_row):
    """
    Extract values from the NAV row and corresponding headers.
    """
    # Clean and prepare the header row
    header_values = []
    for item in header_row:
        if item is not None:
            header_values.append(str(item).strip())
        else:
            header_values.append("")
    
    # Clean and prepare the NAV row
    nav_values = []
    for item in nav_row:
        if item is not None:
            nav_values.append(str(item).strip())
        else:
            nav_values.append("")
    
    # Find the NAV label column
    nav_label_idx = None
    for idx, value in enumerate(nav_values):
        if 'net asset value' in value.lower() or 'nav' == value.lower():
            nav_label_idx = idx
            break
    
    if nav_label_idx is None:
        # Try alternative approach if we didn't find an exact match
        for idx, value in enumerate(nav_values):
            if 'net' in value.lower() and 'asset' in value.lower():
                nav_label_idx = idx
                break
    
    if nav_label_idx is None:
        return None, None, None, None
    
    # Extract the numeric values from columns to the right of the label
    numeric_values = []
    numeric_indices = []
    
    for idx in range(nav_label_idx + 1, len(nav_values)):
        value = nav_values[idx]
        # Skip empty cells
        if not value:
            continue
            
        # Try to clean and convert the value
        try:
            numeric_value = clean_and_convert_value(value)
            if numeric_value is not None:
                numeric_values.append(numeric_value)
                numeric_indices.append(idx)
        except:
            # If conversion fails, it's probably not a numeric value
            pass
    
    # If we found at least two numeric values
    if len(numeric_values) >= 2:
        # Extract corresponding headers
        period1_label = header_values[numeric_indices[0]] if numeric_indices[0] < len(header_values) else "Period 1"
        period2_label = header_values[numeric_indices[1]] if numeric_indices[1] < len(header_values) else "Period 2"
        
        # Use default labels if headers are empty
        if not period1_label:
            period1_label = "Period 1"
        if not period2_label:
            period2_label = "Period 2"
            
        return period1_label, numeric_values[0], period2_label, numeric_values[1]
    
    return None, None, None, None

def clean_and_convert_value(value_str):
    """
    Clean and convert a string value to a float.
    Handles various number formats including apostrophes and commas.
    """
    try:
        # Remove apostrophes, commas, and spaces
        cleaned_value = value_str.replace("'", "").replace(",", "").replace(" ", "")
        
        # Handle percentage values
        if "%" in cleaned_value:
            cleaned_value = cleaned_value.replace("%", "")
            return float(cleaned_value) / 100
            
        return float(cleaned_value)
    except (ValueError, TypeError):
        return None

def process_pdf(pdf_path):
    """
    Process a single PDF file to extract fund name and NAV values.
    """
    try:
        # Extract fund name from the third line of the first page
        fund_name = extract_fund_name(pdf_path)
        
        # Extract tables starting from page 2
        tables = extract_tables_from_pdf(pdf_path, start_page=1)
        
        # Find NAV values in the tables
        period1_label, period1_nav, period2_label, period2_nav = find_nav_in_tables(tables)
        
        # If we couldn't find NAV values in tables, try text-based approach as fallback
        if period1_nav is None or period2_nav is None:
            # Extract text from page 2 onwards
            text_from_page2 = extract_text_from_pdf(pdf_path, start_page=1)
            
            # Look for NAV values in the text
            nav_match = re.search(r"[Nn]et\s+[Aa]sset\s+[Vv]alue.*?(\d[\d\s',\.]*)\s+(\d[\d\s',\.]*)", text_from_page2)
            if nav_match:
                period1_nav = clean_and_convert_value(nav_match.group(1))
                period2_nav = clean_and_convert_value(nav_match.group(2))
                
                # Try to find period labels in various formats
                
                # Standard numeric date format
                date_match_numeric = re.findall(r"(\d{2}[\.\/]\d{2}[\.\/]\d{4})", text_from_page2[:1000])
                
                # Month name format (30. September 2024 or 30 December 2024)
                date_match_text = re.findall(r"(\d{1,2}\.?\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})", text_from_page2[:1000], re.IGNORECASE)
                
                # Combine all found dates
                all_dates = date_match_numeric + date_match_text
                
                if len(all_dates) >= 2:
                    period1_label = all_dates[0]
                    period2_label = all_dates[1]
                else:
                    period1_label = "Period 1"
                    period2_label = "Period 2"
        
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
            'Fund Name': fund_name if 'fund_name' in locals() else f"Error: {os.path.basename(pdf_path)}",
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
