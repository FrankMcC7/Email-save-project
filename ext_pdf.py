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
                    
                    # First try getting the third line if it's not empty
                    if len(lines) >= 3 and lines[2].strip():
                        return lines[2].strip()
                    
                    # Try looking for fund name patterns in the first 10 lines
                    for i in range(min(10, len(lines))):
                        line = lines[i].strip()
                        # Skip empty lines or very short lines
                        if len(line) < 3:
                            continue
                            
                        # Check for fund name indicators
                        if any(indicator in line.lower() for indicator in ['fund', 'sicav', 's.c.a', 'l.p.', 'partners group']):
                            return line
        
        # Fallback to extracting text from first few pages and looking for fund name patterns
        all_text = extract_text_from_pdf(pdf_path, max_pages=3)
        
        # Try to find title lines that might contain fund names
        lines = all_text.split('\n')
        for line in lines[:20]:  # Look only in first 20 lines
            line = line.strip()
            if len(line) > 10 and any(indicator in line.lower() for indicator in ['fund', 'sicav', 's.c.a', 'l.p.', 'partners group']):
                return line
        
        # If all else fails, try to match common fund name patterns
        fund_patterns = [
            r"([A-Za-z0-9\s\-\.&]+(?:S\.C\.A\.|SICAV|Fund|L\.P\.))",
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
    # First, process the table data
    for table in tables:
        if not table or len(table) < 2:
            continue
            
        # Skip tables that are just page text containers
        if len(table) > 0 and len(table[0]) > 0 and table[0][0] == "PAGE_TEXT":
            continue
        
        # First, try to find the header row with any date-like content
        header_row = None
        header_row_idx = -1
        for row_idx, row in enumerate(table):
            # Skip empty rows or rows with only one cell
            if len(row) < 2:
                continue
                
            # Convert all items in row to strings for searching
            str_row = [str(item).strip() if item is not None else "" for item in row]
            row_text = " ".join(str_row).lower()
            
            # Check for various date formats or indicators
            if (any(month.lower() in row_text for month in [
                'january', 'february', 'march', 'april', 'may', 'june', 'july', 
                'august', 'september', 'october', 'november', 'december',
                'jan', 'feb', 'mar', 'apr', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'
            ])):
                header_row = row
                header_row_idx = row_idx
                break
            
            # Check for numeric date formats or quarters
            if any(date_indicator in row_text for date_indicator in ['.20', '/20', '-20']) or any(q in row_text for q in ['q1', 'q2', 'q3', 'q4']):
                header_row = row
                header_row_idx = row_idx
                break
                
            # Check for year formats (assuming recent years)
            years = re.findall(r'\b(20\d\d)\b', row_text)
            if len(years) >= 2:
                header_row = row
                header_row_idx = row_idx
                break
        
        # If we found a header row, now look for NAV row
        if header_row is not None and header_row_idx >= 0:
            for row_idx in range(header_row_idx + 1, len(table)):
                row = table[row_idx]
                if len(row) < 2:  # Skip rows with only one cell
                    continue
                    
                # Convert all items in row to strings for searching
                str_row = [str(item).strip() if item is not None else "" for item in row]
                first_cell = str_row[0].lower() if str_row else ""
                row_text = " ".join(str_row).lower()
                
                # Look for NAV indicators with broader matches
                nav_indicators = [
                    'net asset value', 
                    'net asset', 
                    'nav',
                    'asset value'
                ]
                
                if any(indicator in first_cell for indicator in nav_indicators) or any(indicator in row_text for indicator in nav_indicators):
                    # Found the NAV row - get column data
                    header_values = []
                    for item in header_row:
                        header_values.append(str(item).strip() if item is not None else "")
                    
                    # Get the NAV values - we need to extract all numeric values
                    nav_values = []
                    for idx, item in enumerate(row):
                        if item is not None and idx > 0:  # Skip the first column which contains the label
                            try:
                                value_str = str(item).strip()
                                if value_str and not value_str.isalpha():
                                    cleaned_value = clean_and_convert_value(value_str)
                                    if cleaned_value is not None:
                                        nav_values.append((idx, cleaned_value))
                            except:
                                pass
                    
                    # If we found at least two values
                    if len(nav_values) >= 2:
                        idx1, value1 = nav_values[0]
                        idx2, value2 = nav_values[1]
                        
                        # Get the corresponding headers
                        period1_label = header_values[idx1] if idx1 < len(header_values) else "Period 1"
                        period2_label = header_values[idx2] if idx2 < len(header_values) else "Period 2"
                        
                        # If headers are empty or just column indicators like "1" or "2", use default labels
                        if not period1_label or period1_label.isdigit():
                            period1_label = "Period 1"
                        if not period2_label or period2_label.isdigit():
                            period2_label = "Period 2"
                        
                        return period1_label, value1, period2_label, value2
    
    # If we still haven't found NAV values, try a different approach - look for any row with NAV first
    for table in tables:
        if not table or len(table) < 2:
            continue
        
        # Skip tables that are just page text containers
        if len(table) > 0 and len(table[0]) > 0 and table[0][0] == "PAGE_TEXT":
            continue
            
        # Find any row that might contain NAV
        for row_idx, row in enumerate(table):
            if len(row) < 2:  # Skip rows with only one cell
                continue
                
            # Check first column specifically
            first_cell = str(row[0]).strip().lower() if row[0] is not None else ""
            row_text = " ".join([str(item).strip() if item is not None else "" for item in row]).lower()
            
            nav_indicators = [
                'net asset value', 
                'net asset', 
                'nav',
                'asset value'
            ]
            
            if any(indicator in first_cell for indicator in nav_indicators) or any(indicator in row_text for indicator in nav_indicators):
                # Found potential NAV row - extract numeric values
                numeric_values = []
                for idx, item in enumerate(row):
                    if item is not None and idx > 0:  # Skip the first column
                        try:
                            value_str = str(item).strip()
                            if value_str and not value_str.isalpha():
                                cleaned_value = clean_and_convert_value(value_str)
                                if cleaned_value is not None:
                                    numeric_values.append((idx, cleaned_value))
                        except:
                            pass
                
                # If we found at least two numeric values
                if len(numeric_values) >= 2:
                    # Look for a header row above this row
                    header_row = None
                    for i in range(row_idx-1, max(0, row_idx-5), -1):  # Look up to 5 rows above
                        if i < 0 or i >= len(table):
                            continue
                            
                        potential_header = table[i]
                        if len(potential_header) < len(row):  # Skip if potential header has fewer columns
                            continue
                            
                        # Check if this row has date-like content
                        header_text = " ".join([str(item).strip() if item is not None else "" for item in potential_header]).lower()
                        
                        # Check for month names, dates, years, or quarters
                        if (any(month.lower() in header_text for month in [
                                'january', 'february', 'march', 'april', 'may', 'june', 'july', 
                                'august', 'september', 'october', 'november', 'december',
                                'jan', 'feb', 'mar', 'apr', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'
                            ]) or
                            any(date_indicator in header_text for date_indicator in ['.20', '/20', '-20']) or
                            any(q in header_text for q in ['q1', 'q2', 'q3', 'q4']) or
                            re.search(r'\b(20\d\d)\b', header_text)):
                            
                            header_row = potential_header
                            break
                    
                    idx1, value1 = numeric_values[0]
                    idx2, value2 = numeric_values[1]
                    
                    if header_row:
                        header_values = [str(item).strip() if item is not None else "" for item in header_row]
                        period1_label = header_values[idx1] if idx1 < len(header_values) else "Period 1"
                        period2_label = header_values[idx2] if idx2 < len(header_values) else "Period 2"
                        
                        # If headers are empty or just column indicators, use default labels
                        if not period1_label or period1_label.isdigit():
                            period1_label = "Period 1"
                        if not period2_label or period2_label.isdigit():
                            period2_label = "Period 2"
                    else:
                        # If no header row found, use default labels
                        period1_label = "Period 1"
                        period2_label = "Period 2"
                    
                    return period1_label, value1, period2_label, value2
    
    # If we reached here, check for page text
    for table in tables:
        if not table or len(table) < 2:
            continue
            
        # Look for page text
        if len(table) > 0 and len(table[0]) > 0 and table[0][0] == "PAGE_TEXT":
            page_text = table[1][0]
            
            # Use regex with more flexible patterns to find NAV values
            nav_patterns = [
                r"[Nn]et\s+[Aa]sset\s+[Vv]alue.*?(\d[\d\s',\.]*)\s+(\d[\d\s',\.]*)",
                r"[Nn]et\s+[Aa]sset.*?(\d[\d\s',\.]*)\s+(\d[\d\s',\.]*)",
                r"NAV.*?(\d[\d\s',\.]*)\s+(\d[\d\s',\.]*)"
            ]
            
            for pattern in nav_patterns:
                nav_match = re.search(pattern, page_text)
                if nav_match:
                    # Look for date patterns
                    date_patterns = [
                        # Standard numeric date format (DD.MM.YYYY or DD/MM/YYYY)
                        r"(\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4})",
                        
                        # Month name format (DD. Month YYYY or DD Month YYYY)
                        r"(\d{1,2}\.?\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})",
                        
                        # Short month format (DD. MMM YYYY)
                        r"(\d{1,2}\.?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4})",
                        
                        # Quarter format (Q1 YYYY)
                        r"(Q[1-4]\s+\d{4})"
                    ]
                    
                    all_dates = []
                    for date_pattern in date_patterns:
                        found_dates = re.findall(date_pattern, page_text, re.IGNORECASE)
                        all_dates.extend(found_dates)
                    
                    if len(all_dates) >= 2:
                        return all_dates[0], clean_and_convert_value(nav_match.group(1)), all_dates[1], clean_and_convert_value(nav_match.group(2))
                    else:
                        # If no dates found, look for years
                        years = re.findall(r'\b(20\d\d)\b', page_text)
                        if len(years) >= 2:
                            return f"Year {years[0]}", clean_and_convert_value(nav_match.group(1)), f"Year {years[1]}", clean_and_convert_value(nav_match.group(2))
                        else:
                            return "Period 1", clean_and_convert_value(nav_match.group(1)), "Period 2", clean_and_convert_value(nav_match.group(2))
    
    return None, None, None, None

def extract_nav_format1(tables, page_text):
    """
    Extract NAV based on format 1: Table with month names in header
    Like in the first sample image with "30 September 2024" and "31 December 2024"
    """
    # Check if the page contains month names in headers
    month_pattern = r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\b'
    if not re.search(month_pattern, page_text, re.IGNORECASE):
        return None, None, None, None
        
    for table in tables:
        if not table or len(table) < 4:  # Need enough rows
            continue
            
        # Look for header row with month names
        header_row = None
        header_row_idx = -1
        
        for row_idx, row in enumerate(table):
            if len(row) < 3:  # Need at least 3 columns
                continue
                
            row_text = " ".join([str(cell).strip() if cell is not None else "" for cell in row]).lower()
            
            # Check if this row contains month names
            if re.search(month_pattern, row_text, re.IGNORECASE):
                header_row = row
                header_row_idx = row_idx
                break
        
        if header_row is None:
            continue
            
        # Now look for "Net asset value" row below the header
        for row_idx in range(header_row_idx + 1, len(table)):
            row = table[row_idx]
            if len(row) < 3:  # Need at least 3 columns
                continue
                
            # Check first column for "Net asset value"
            first_col = str(row[0]).strip().lower() if row[0] is not None else ""
            
            if "net asset value" in first_col or "net asset" in first_col:
                # Found the NAV row - extract values from columns 1 and 2
                values = []
                
                for col_idx in range(1, min(len(row), 5)):  # Check first few columns
                    if row[col_idx] is not None:
                        cell_value = str(row[col_idx]).strip()
                        nav_value = clean_and_convert_value(cell_value)
                        if nav_value is not None:
                            values.append((col_idx, nav_value))
                
                if len(values) >= 2:
                    # Get the corresponding headers
                    idx1, val1 = values[0]
                    idx2, val2 = values[1]
                    
                    header1 = str(header_row[idx1]).strip() if idx1 < len(header_row) and header_row[idx1] is not None else "Period 1"
                    header2 = str(header_row[idx2]).strip() if idx2 < len(header_row) and header_row[idx2] is not None else "Period 2"
                    
                    return header1, val1, header2, val2
                    
    return None, None, None, None

def extract_nav_format2(tables, page_text):
    """
    Extract NAV based on format 2: Table with numeric dates
    Like in the second sample image with "30.09.2024" and "31.12.2024"
    """
    # Check if the page contains numeric dates
    numeric_date_pattern = r'\d{2}[\.\/]\d{2}[\.\/]\d{4}'
    if not re.search(numeric_date_pattern, page_text):
        return None, None, None, None
        
    for table in tables:
        if not table or len(table) < 4:  # Need enough rows
            continue
            
        # Look for header row with numeric dates
        header_row = None
        header_row_idx = -1
        
        for row_idx, row in enumerate(table):
            if len(row) < 3:  # Need at least 3 columns
                continue
                
            row_text = " ".join([str(cell).strip() if cell is not None else "" for cell in row])
            
            # Check if this row contains numeric dates (at least 2)
            date_matches = re.findall(numeric_date_pattern, row_text)
            if len(date_matches) >= 2:
                header_row = row
                header_row_idx = row_idx
                break
        
        if header_row is None:
            continue
            
        # Now look for "Net asset value" row below the header
        for row_idx in range(header_row_idx + 1, len(table)):
            row = table[row_idx]
            if len(row) < 3:  # Need at least 3 columns
                continue
                
            # Check first column for "Net asset value"
            first_col = str(row[0]).strip().lower() if row[0] is not None else ""
            
            if "net asset value" in first_col or "net asset" in first_col:
                # Found the NAV row - extract values from columns 1 and 2
                values = []
                
                for col_idx in range(1, min(len(row), 5)):  # Check first few columns
                    if row[col_idx] is not None:
                        cell_value = str(row[col_idx]).strip()
                        nav_value = clean_and_convert_value(cell_value)
                        if nav_value is not None:
                            values.append((col_idx, nav_value))
                
                if len(values) >= 2:
                    # Get the corresponding headers
                    idx1, val1 = values[0]
                    idx2, val2 = values[1]
                    
                    header1 = str(header_row[idx1]).strip() if idx1 < len(header_row) and header_row[idx1] is not None else "Period 1"
                    header2 = str(header_row[idx2]).strip() if idx2 < len(header_row) and header_row[idx2] is not None else "Period 2"
                    
                    return header1, val1, header2, val2
                    
    return None, None, None, None

def extract_nav_from_text(text):
    """
    Fallback method to extract NAV values from text when table approaches fail.
    """
    # Find period labels (dates)
    period_labels = []
    
    # Look for different date formats
    date_patterns = [
        r"(\d{2}\.\d{2}\.\d{4})",  # DD.MM.YYYY
        r"(\d{2}/\d{2}/\d{4})",    # DD/MM/YYYY
        r"(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})" # DD Month YYYY
    ]
    
    for pattern in date_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        period_labels.extend(matches)
    
    # Look for "Net asset value" and extract numbers
    nav_patterns = [
        r"[Nn]et\s+asset\s+value.*?(\d[\d\s',\.]*)\s+(\d[\d\s',\.]*)",
        r"[Nn]et\s+asset.*?(\d[\d\s',\.]*)\s+(\d[\d\s',\.]*)",
        r"NAV.*?(\d[\d\s',\.]*)\s+(\d[\d\s',\.]*)"
    ]
    
    for pattern in nav_patterns:
        match = re.search(pattern, text)
        if match:
            period1_nav = clean_and_convert_value(match.group(1))
            period2_nav = clean_and_convert_value(match.group(2))
            
            # Use period labels if found
            if len(period_labels) >= 2:
                return period_labels[0], period1_nav, period_labels[1], period2_nav
            else:
                return "Period 1", period1_nav, "Period 2", period2_nav
    
    return None, None, None, None

def clean_and_convert_value(value_str):
    """
    Clean and convert a string value to a float.
    Handles various number formats including apostrophes and commas.
    """
    try:
        if value_str is None:
            return None
            
        # Handle percentage values
        if isinstance(value_str, str) and "%" in value_str:
            cleaned_value = value_str.replace("%", "").replace(" ", "")
            return float(cleaned_value) / 100
        
        # If it's already a number, return it
        if isinstance(value_str, (int, float)):
            return float(value_str)
            
        if isinstance(value_str, str):
            # Remove apostrophes, commas, spaces
            cleaned_value = value_str.replace("'", "").replace(",", "").replace(" ", "").replace("x", "")
            
            # Handle cases where there's no digits
            if not any(c.isdigit() for c in cleaned_value):
                return None
                
            # Handle any remaining non-numeric characters (except decimal point)
            final_value = ""
            decimal_found = False
            for char in cleaned_value:
                if char.isdigit():
                    final_value += char
                elif char == "." and not decimal_found:
                    final_value += char
                    decimal_found = True
                    
            if final_value:
                return float(final_value)
    except (ValueError, TypeError):
        pass
        
    return None

def process_pdf(pdf_path):
    """
    Process a single PDF file to extract fund name and NAV values.
    Using the two specific format approaches shown in the sample images.
    """
    try:
        # Extract fund name from the third line of the first page
        fund_name = extract_fund_name(pdf_path)
        
        # Extract tables from PDF starting from page 2
        tables = extract_tables_from_pdf(pdf_path, start_page=1)
        
        # Get page text from page 2 for pattern matching
        page_text = extract_text_from_pdf(pdf_path, start_page=1, max_pages=1)
        
        # Try format 1: Table with month names in header (e.g., "30 September 2024")
        # As shown in the first sample image
        period1_label, period1_nav, period2_label, period2_nav = extract_nav_format1(tables, page_text)
        
        # If format 1 failed, try format 2: Table with numeric dates (e.g., "30.09.2024")
        # As shown in the second sample image
        if period1_nav is None or period2_nav is None:
            period1_label, period1_nav, period2_label, period2_nav = extract_nav_format2(tables, page_text)
            
        # If both specific formats failed, fall back to general extraction
        if period1_nav is None or period2_nav is None:
            # This will be our fallback for any other format
            period1_label, period1_nav, period2_label, period2_nav = find_nav_in_tables(tables)
        
        # If all table approaches failed, try text-based extraction as final fallback
        if period1_nav is None or period2_nav is None:
            period1_label, period1_nav, period2_label, period2_nav = extract_nav_from_text(page_text)
        
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
