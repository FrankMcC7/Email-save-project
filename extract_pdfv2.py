import os
import re
import pandas as pd
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import glob
from datetime import datetime
import traceback
import warnings
warnings.filterwarnings('ignore')

# Try to import additional libraries
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False
    print("PyMuPDF not available. For better results, install with: pip install PyMuPDF")

try:
    import tabula
    TABULA_AVAILABLE = True
except ImportError:
    TABULA_AVAILABLE = False
    print("Tabula not available. For better table extraction, install with: pip install tabula-py")

try:
    import subprocess
    POPPLER_AVAILABLE = True
except ImportError:
    POPPLER_AVAILABLE = False


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


def analyze_pdf_structure(pdf_path):
    """
    Analyze PDF to determine its structure and guide extraction approach.
    """
    pdf_info = {
        "has_tables": False,
        "is_image_based": False,
        "has_key_figures_section": False,
        "page_count": 0
    }
    
    try:
        # Try with pdfplumber first
        with pdfplumber.open(pdf_path) as pdf:
            pdf_info["page_count"] = len(pdf.pages)
            
            # Check first 3 pages
            for i in range(min(3, len(pdf.pages))):
                page = pdf.pages[i]
                page_text = page.extract_text() or ""
                
                # Check if this PDF has tables
                if page.find_tables():
                    pdf_info["has_tables"] = True
                
                # Check if this PDF has "Key Figures" section
                if "key figures" in page_text.lower():
                    pdf_info["has_key_figures_section"] = True
                    
                # If page has very little extractable text, might be image-based
                if len(page_text) < 100:
                    pdf_info["is_image_based"] = True
                    
            # If we haven't found tables yet, try a different approach
            if not pdf_info["has_tables"] and PYMUPDF_AVAILABLE:
                doc = fitz.open(pdf_path)
                for i in range(min(3, doc.page_count)):
                    page = doc[i]
                    if page.get_text("dict")["blocks"]:
                        # Check for structures that look like tables
                        blocks = page.get_text("dict")["blocks"]
                        # Look for grid-like text arrangement
                        if any(len(block.get("lines", [])) > 3 for block in blocks):
                            pdf_info["has_tables"] = True
        
        return pdf_info
    except Exception as e:
        print(f"Error analyzing PDF structure: {str(e)}")
        return pdf_info


def extract_text_multi_library(pdf_path):
    """
    Try multiple libraries to extract text from PDF.
    """
    text = ""
    
    # Try pdfplumber first
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text += page_text + "\n"
        
        if text.strip():
            return text
    except Exception as e:
        print(f"pdfplumber extraction failed: {str(e)}")
    
    # Try PyMuPDF if available
    if PYMUPDF_AVAILABLE:
        try:
            doc = fitz.open(pdf_path)
            text = ""
            for page in doc:
                text += page.get_text() + "\n"
            
            if text.strip():
                return text
        except Exception as e:
            print(f"PyMuPDF extraction failed: {str(e)}")
    
    # If all else fails, use the original extract_text_from_pdf function
    return extract_text_from_pdf(pdf_path)


def extract_tables_multi_library(pdf_path):
    """
    Try multiple libraries to extract tables from PDF.
    """
    all_tables = []
    
    # Try pdfplumber with different settings
    table_settings_list = [
        {'vertical_strategy': 'lines', 'horizontal_strategy': 'lines'},
        {'vertical_strategy': 'text', 'horizontal_strategy': 'text'},
        {'vertical_strategy': 'lines', 'horizontal_strategy': 'text'},
        {'vertical_strategy': 'text', 'horizontal_strategy': 'lines'}
    ]
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num in range(min(5, len(pdf.pages))):
                page = pdf.pages[page_num]
                
                for settings in table_settings_list:
                    try:
                        tables = page.extract_tables(settings)
                        if tables:
                            all_tables.extend([{'page': page_num, 'tables': tables, 'source': 'pdfplumber'}])
                            break  # If we found tables with these settings, move to next page
                    except:
                        continue
    except Exception as e:
        print(f"pdfplumber table extraction failed: {str(e)}")
    
    # Try tabula if available
    if TABULA_AVAILABLE and not all_tables:
        try:
            tabula_tables = tabula.read_pdf(pdf_path, pages='1-5', multiple_tables=True)
            if tabula_tables:
                # Convert tabula tables to our format
                for table in tabula_tables:
                    if not table.empty:
                        all_tables.append({'page': 0, 'tables': [table.values.tolist()], 'source': 'tabula'})
        except Exception as e:
            print(f"tabula table extraction failed: {str(e)}")
    
    return all_tables


def clean_number(value_str):
    """
    Clean and convert a string value to a float.
    Handles various number formats including apostrophes, commas, spaces.
    """
    try:
        if value_str is None:
            return None
            
        # If it's already a number, return it
        if isinstance(value_str, (int, float)):
            return float(value_str)
            
        if isinstance(value_str, str):
            # Handle percentage values - skip these
            if "%" in value_str:
                return None
            
            # Skip non-numeric text
            if len(value_str.strip()) == 0 or value_str.strip().isalpha():
                return None
                
            # Remove common separators and non-numeric characters
            cleaned = value_str.replace("'", "").replace(",", "").replace(" ", "").replace("x", "")
            
            # Handle cases where there's no digits
            if not any(c.isdigit() for c in cleaned):
                return None
                
            # Handle any remaining non-numeric characters (except decimal point)
            final_value = ""
            decimal_found = False
            for char in cleaned:
                if char.isdigit():
                    final_value += char
                elif char == "." and not decimal_found:
                    final_value += char
                    decimal_found = True
                    
            if final_value:
                # Convert to float with additional validation
                try:
                    value = float(final_value)
                    
                    # If we have a non-zero value and it's reasonably large (NAV values are typically large)
                    # This helps avoid misidentifying small numbers as NAV values
                    if value > 100:  # NAV is typically in thousands or millions
                        return value
                except:
                    pass
    except:
        pass
        
    return None


def extract_nav_from_tables(tables_data):
    """
    Extract NAV information from tables.
    """
    results = []
    
    for table_entry in tables_data:
        page = table_entry['page']
        tables = table_entry['tables']
        
        for table in tables:
            # Skip empty or tiny tables
            if not table or len(table) < 2:
                continue
            
            # Look for NAV values in this table
            for row_idx, row in enumerate(table):
                if not row or len(row) < 2:
                    continue
                
                row_str = [str(cell).strip() if cell is not None else "" for cell in row]
                row_text = " ".join(row_str).lower()
                
                # Check if this row contains NAV information
                if "net asset value" in row_text or "nav" in row_text.split():
                    # Get values
                    values = []
                    for cell_idx, cell in enumerate(row):
                        if cell_idx > 0:  # Skip first column (label)
                            value = clean_number(str(cell))
                            if value is not None and value > 100:  # NAV values are typically large
                                values.append((cell_idx, value))
                    
                    if len(values) >= 2:
                        # Look for header row (dates)
                        header_row = None
                        for i in range(row_idx):
                            if len(table[i]) >= len(row):
                                header_text = " ".join([str(cell).strip() if cell is not None else "" for cell in table[i]])
                                # Check for date patterns
                                if re.search(r'\b\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4}\b|\b\d{2}\.\d{2}\.\d{4}\b', header_text):
                                    header_row = table[i]
                                    break
                        
                        if header_row:
                            idx1, value1 = values[0]
                            idx2, value2 = values[1]
                            
                            header1 = str(header_row[idx1]).strip() if idx1 < len(header_row) and header_row[idx1] is not None else "Period 1"
                            header2 = str(header_row[idx2]).strip() if idx2 < len(header_row) and header_row[idx2] is not None else "Period 2"
                            
                            results.append({
                                'period1_label': header1,
                                'period1_nav': value1,
                                'period2_label': header2,
                                'period2_nav': value2,
                                'confidence': 0.9,  # High confidence for table with header
                                'source': f'table_p{page}'
                            })
                        else:
                            # No header row found, extract values anyway
                            results.append({
                                'period1_label': "Period 1",
                                'period1_nav': values[0][1],
                                'period2_label': "Period 2",
                                'period2_nav': values[1][1],
                                'confidence': 0.7,  # Medium confidence without header
                                'source': f'table_p{page}'
                            })
    
    # Sort by confidence
    return sorted(results, key=lambda x: x['confidence'], reverse=True)


def extract_nav_with_enhanced_patterns(text):
    """
    Use enhanced patterns to extract NAV values from text.
    """
    results = []
    
    # Enhanced NAV patterns
    nav_patterns = [
        # Pattern 1: "Net Asset Value: 12,345.67" or "NAV: 12,345.67"
        r"(?:Net\s+Asset\s+Value|NAV)[:\s]+([0-9\s,.\']+)",
        
        # Pattern 2: "Net Asset Value per Share: 12,345.67"
        r"(?:Net\s+Asset\s+Value|NAV)\s+per\s+(?:Share|Unit)[:\s]+([0-9\s,.\']+)",
        
        # Pattern 3: "Net Asset Value as of 31.12.2021: 12,345.67"
        r"(?:Net\s+Asset\s+Value|NAV)\s+(?:as\s+of|on)\s+[0-9.\/]+[:\s]+([0-9\s,.\']+)",
        
        # Pattern 4: Table-like format with labels in first column
        r"(?:Net\s+Asset\s+Value|NAV)[\s\|]+([0-9\s,.\']+)[\s\|]+([0-9\s,.\']+)"
    ]
    
    # Date patterns
    date_patterns = [
        r'\b(\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4})\b',
        r'\b(\d{2}\.\d{2}\.\d{4})\b',
        r'\b((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4})\b',
        r'\b((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},?\s+\d{4})\b'
    ]
    
    # Extract dates from the text
    dates = []
    for pattern in date_patterns:
        matches = re.findall(pattern, text)
        dates.extend(matches)
    
    # Ensure we have unique dates
    dates = list(set(dates))
    
    # Sort dates (assuming they are in standard formats)
    dates.sort()
    
    # Extract NAV values
    nav_values = []
    
    # Try all patterns
    for pattern in nav_patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            if isinstance(match, tuple):
                # Handle multiple groups
                for group in match:
                    value = clean_number(group)
                    if value is not None and value > 100:  # NAV values are typically large
                        nav_values.append(value)
            else:
                value = clean_number(match)
                if value is not None and value > 100:  # NAV values are typically large
                    nav_values.append(value)
    
    # Look for pairs of values that might be current and previous NAVs
    if len(nav_values) >= 2:
        # If we have dates and values
        if len(dates) >= 2:
            results.append({
                'period1_label': dates[-2],  # Second most recent date
                'period1_nav': nav_values[1],  # Second value
                'period2_label': dates[-1],  # Most recent date
                'period2_nav': nav_values[0],  # First value
                'confidence': 0.8,  # Good confidence with dates
                'source': 'pattern_with_dates'
            })
        else:
            # If we have values but no dates
            results.append({
                'period1_label': "Period 1",
                'period1_nav': nav_values[1],  # Second value
                'period2_label': "Period 2", 
                'period2_nav': nav_values[0],  # First value
                'confidence': 0.6,  # Lower confidence without dates
                'source': 'pattern_no_dates'
            })
    
    return results


def direct_table_extraction(pdf_path):
    """
    Directly target the Key Figures table and extract Net asset value.
    This approach focuses specifically on the exact table structure seen in examples.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Check pages 1-5 for the key figures table
            for page_num in range(min(5, len(pdf.pages))):
                page = pdf.pages[page_num]
                
                # Get the text to check for Key Figures section
                page_text = page.extract_text() or ""
                
                # Only process pages with "Key figures" or "key figures"
                if "Key figures" in page_text or "key figures" in page_text.lower():
                    # Extract all tables from this page
                    tables = page.extract_tables()
                    
                    # Process each table
                    for table in tables:
                        if not table or len(table) < 3:  # Skip small tables
                            continue
                        
                        # Look for NAV row in this table
                        for row_idx, row in enumerate(table):
                            # Skip rows with fewer than 2 columns
                            if len(row) < 3:
                                continue
                                
                            # Convert to string and check if this is the NAV row
                            row_str = [str(cell).strip() if cell is not None else "" for cell in row]
                            first_cell = row_str[0].lower()
                            
                            if "net asset value" in first_cell or "net asset" in first_cell:
                                # Found the NAV row - now get header and values
                                
                                # Look for header row (containing dates or periods)
                                header_row = None
                                for i in range(row_idx):
                                    potential_header = table[i]
                                    if len(potential_header) >= len(row):
                                        header_text = " ".join([str(cell).strip() if cell is not None else "" for cell in potential_header])
                                        # Check if it contains dates or period indicators
                                        if (re.search(r'\b\d{1,2}[\.\/\s]+(?:[A-Za-z]+|\d{1,2})[\.\/\s]+\d{4}\b', header_text) or
                                            re.search(r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b', header_text, re.IGNORECASE) or
                                            re.search(r'\b(?:Q[1-4]|20\d\d)\b', header_text)):
                                            header_row = potential_header
                                            break
                                
                                # If no proper header found, use the row above if available
                                if header_row is None and row_idx > 0:
                                    header_row = table[row_idx - 1]
                                
                                # Get values from NAV row (skip first column which is the label)
                                values = []
                                for cell_idx in range(1, len(row)):
                                    if row[cell_idx] is not None:
                                        # Clean and convert the value
                                        try:
                                            value_str = str(row[cell_idx]).strip()
                                            value = clean_number(value_str)
                                            if value is not None:
                                                values.append((cell_idx, value))
                                        except:
                                            pass
                                
                                # Need at least two values
                                if len(values) >= 2:
                                    # Get corresponding headers
                                    idx1, value1 = values[0]
                                    idx2, value2 = values[1]
                                    
                                    if header_row:
                                        header1 = str(header_row[idx1]).strip() if idx1 < len(header_row) and header_row[idx1] is not None else "Period 1"
                                        header2 = str(header_row[idx2]).strip() if idx2 < len(header_row) and header_row[idx2] is not None else "Period 2"
                                    else:
                                        # Extract dates from page text if headers not found
                                        date_matches = re.findall(r'\b\d{1,2}[\.\/\s]+(?:[A-Za-z]+|\d{1,2})[\.\/\s]+\d{4}\b|\b\d{2}\.\d{2}\.\d{4}\b|\b(?:Q[1-4]|20\d\d)\b', page_text)
                                        if len(date_matches) >= 2:
                                            header1 = date_matches[0]
                                            header2 = date_matches[1]
                                        else:
                                            header1 = "Period 1"
                                            header2 = "Period 2"
                                    
                                    return header1, value1, header2, value2
    except Exception as e:
        print(f"Error in direct table extraction: {str(e)}")
        
    return None, None, None, None


def scan_text_for_nav(pdf_path):
    """
    Scan the PDF text for Net Asset Value mentions and extract data.
    This is a fallback approach when table extraction fails.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Extract text from the first few pages
            all_text = ""
            for page_num in range(min(5, len(pdf.pages))):
                page_text = pdf.pages[page_num].extract_text() or ""
                all_text += page_text + "\n"
                
                # Check each page individually first
                lines = page_text.split('\n')
                
                # 1. First, look for "Net asset value" line specifically
                for i, line in enumerate(lines):
                    if "net asset value" in line.lower():
                        # Found NAV line - extract numbers
                        number_matches = re.findall(r"[\d',\.]+", line)
                        values = []
                        
                        for num in number_matches:
                            value = clean_number(num)
                            if value is not None and value > 1000:
                                values.append(value)
                        
                        if len(values) >= 2:
                            # Look for date headers in previous lines
                            date_matches = []
                            for j in range(max(0, i-5), i):
                                date_match = re.findall(r'\b\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4}\b|\b\d{2}\.\d{2}\.\d{4}\b', lines[j])
                                if date_match:
                                    date_matches.extend(date_match)
                            
                            if len(date_matches) >= 2:
                                return date_matches[0], values[0], date_matches[1], values[1]
                            else:
                                # Look for dates in the entire page
                                date_matches = re.findall(r'\b\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4}\b|\b\d{2}\.\d{2}\.\d{4}\b', page_text)
                                if len(date_matches) >= 2:
                                    return date_matches[0], values[0], date_matches[1], values[1]
                                else:
                                    return "Period 1", values[0], "Period 2", values[1]
            
            # If page-by-page approach failed, try with whole text
            nav_matches = re.finditer(r'(?:[Nn]et\s+[Aa]sset\s+[Vv]alue|NAV)', all_text)
            
            for match in nav_matches:
                # Extract text around the match
                start_pos = max(0, match.start() - 100)
                end_pos = min(len(all_text), match.end() + 300)
                context = all_text[start_pos:end_pos]
                
                # Find dates in the context
                dates = re.findall(r'\b\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4}\b|\b\d{2}\.\d{2}\.\d{4}\b', context)
                
                # Find large numbers that could be NAV values
                number_matches = re.findall(r"[\d',\.]+", context)
                values = []
                
                for num in number_matches:
                    value = clean_number(num)
                    if value is not None and value > 10000:  # NAV values are typically large
                        values.append(value)
                
                if len(dates) >= 2 and len(values) >= 2:
                    return dates[0], values[0], dates[1], values[1]
    except Exception as e:
        print(f"Error in text scanning: {str(e)}")
        
    return None, None, None, None


def last_resort_extraction(pdf_path):
    """
    Absolute last resort approach - specifically look for patterns like in the Excel screenshot.
    This approach focuses on known formats from the examples provided.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Check first 3 pages
            for page_num in range(min(3, len(pdf.pages))):
                page = pdf.pages[page_num]
                page_text = page.extract_text() or ""
                
                # Look specifically for "Net asset value" in a line
                lines = page_text.split('\n')
                
                # First approach: Look for the exact line format with Net asset value
                for line in lines:
                    if "net asset value" in line.lower():
                        # Try to extract exactly two numbers from this line
                        number_matches = re.findall(r"[\d',\.]+", line)
                        values = []
                        
                        for num in number_matches:
                            value = clean_number(num)
                            if value is not None and value > 10000:  # NAV values are typically large
                                values.append(value)
                        
                        if len(values) >= 2:
                            # Extract dates from the page text
                            date_matches = re.findall(r'\b\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4}\b|\b\d{2}\.\d{2}\.\d{4}\b', page_text)
                            
                            if len(date_matches) >= 2:
                                return date_matches[0], values[0], date_matches[1], values[1]
                            else:
                                # If no dates found, just use Period 1 and Period 2
                                return "Period 1", values[0], "Period 2", values[1]
                
                # Second approach: Look for numeric patterns in lines with dates
                date_line_idx = -1
                for i, line in enumerate(lines):
                    if re.search(r'\b\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4}\b|\b\d{2}\.\d{2}\.\d{4}\b', line):
                        date_line_idx = i
                        break
                
                if date_line_idx >= 0:
                    # Found a line with dates, now look for "Net asset value" within the next 15 lines
                    for i in range(date_line_idx + 1, min(date_line_idx + 15, len(lines))):
                        if "net asset value" in lines[i].lower():
                            # Found the NAV line
                            number_matches = re.findall(r"[\d',\.]+", lines[i])
                            values = []
                            
                            for num in number_matches:
                                value = clean_number(num)
                                if value is not None and value > 10000:
                                    values.append(value)
                            
                            if len(values) >= 2:
                                # Extract dates from the date line
                                date_matches = re.findall(r'\b\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4}\b|\b\d{2}\.\d{2}\.\d{4}\b', lines[date_line_idx])
                                
                                if len(date_matches) >= 2:
                                    return date_matches[0], values[0], date_matches[1], values[1]
            
            # Third approach: Process all lines looking for specific formats like in the Excel
            for page_num in range(min(5, len(pdf.pages))):
                page_text = pdf.pages[page_num].extract_text() or ""
                lines = page_text.split('\n')
                
                for line in lines:
                    # Look for lines that match the format "some text number1 number2"
                    if "asset" in line.lower() or "value" in line.lower():
                        number_matches = re.findall(r"[\d',\.]+", line)
                        values = []
                        
                        for num in number_matches:
                            value = clean_number(num)
                            if value is not None and value > 10000:
                                values.append(value)
                        
                        if len(values) >= 2:
                            # Extract dates from the page text
                            date_matches = re.findall(r'\b\d{1,2}[\.\/]\d{1,2}[\.\/]\d{4}\b|\b\d{2}\.\d{2}\.\d{4}\b', page_text)
                            
                            if len(date_matches) >= 2:
                                return date_matches[0], values[0], date_matches[1], values[1]
    except Exception as e:
        print(f"Error in last resort extraction: {str(e)}")
        
    return None, None, None, None


def process_pdf_comprehensive(pdf_path):
    """
    Process a PDF using multiple approaches in an intelligent sequence.
    """
    # Step 1: Extract fund name
    fund_name = extract_fund_name(pdf_path)
    
    # Step 2: Analyze PDF structure to determine best approach
    pdf_info = analyze_pdf_structure(pdf_path)
    
    results = []
    
    # Step 3: Multi-library text extraction
    text = extract_text_multi_library(pdf_path)
    
    # Step 4: Extract NAV using enhanced patterns from text
    pattern_results = extract_nav_with_enhanced_patterns(text)
    if pattern_results:
        results.extend(pattern_results)
    
    # Step 5: Table-based extraction if the PDF likely has tables
    if pdf_info["has_tables"]:
        tables_data = extract_tables_multi_library(pdf_path)
        table_results = extract_nav_from_tables(tables_data)
        if table_results:
            results.extend(table_results)
    
    # Step 6: If we have Key Figures section, use specialized approach
    if pdf_info["has_key_figures_section"]:
        # Use the original direct_table_extraction function
        period1_label, period1_nav, period2_label, period2_nav = direct_table_extraction(pdf_path)
        if period1_nav is not None and period2_nav is not None:
            results.append({
                'period1_label': period1_label,
                'period1_nav': period1_nav,
                'period2_label': period2_label,
                'period2_nav': period2_nav,
                'confidence': 0.85,  # High confidence for key figures
                'source': 'key_figures'
            })
    
    # Step 7: If all else fails, try the existing fallback methods
    if not results:
        # Use existing methods
        period1_label, period1_nav, period2_label, period2_nav = scan_text_for_nav(pdf_path)
        if period1_nav is not None and period2_nav is not None:
            results.append({
                'period1_label': period1_label,
                'period1_nav': period1_nav,
                'period2_label': period2_label,
                'period2_nav': period2_nav,
                'confidence': 0.6,
                'source': 'scan_text'
            })
        else:
            period1_label, period1_nav, period2_label, period2_nav = last_resort_extraction(pdf_path)
            if period1_nav is not None and period2_nav is not None:
                results.append({
                    'period1_label': period1_label,
                    'period1_nav': period1_nav,
                    'period2_label': period2_label,
                    'period2_nav': period2_nav,
                    'confidence': 0.5,
                    'source': 'last_resort'
                })
    
    # Step 8: Select best result based on confidence
    if results:
        # Sort by confidence
        results.sort(key=lambda x: x['confidence'], reverse=True)
        best_result = results[0]
        
        return {
            'Fund Name': fund_name,
            'Period 1 Label': best_result['period1_label'],
            'Period 1 NAV': best_result['period1_nav'],
            'Period 2 Label': best_result['period2_label'],
            'Period 2 NAV': best_result['period2_nav'],
            'PDF Filename': os.path.basename(pdf_path),
            'Extraction Method': best_result['source'],
            'Confidence': best_result['confidence']
        }
    
    # If all approaches failed
    return {
        'Fund Name': fund_name,
        'Period 1 Label': 'Period 1',
        'Period 1 NAV': None,
        'Period 2 Label': 'Period 2',
        'Period 2 NAV': None,
        'PDF Filename': os.path.basename(pdf_path),
        'Extraction Method': 'failed',
        'Confidence': 0.0
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
    """
    Main function to process a folder of PDF files.
    """
    print("=== Comprehensive PDF NAV Extraction Tool ===")
    print("This tool extracts fund names and NAV values from PDF documents using multiple approaches.")
    
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
    
    # Process each PDF file using the comprehensive approach
    results = []
    success_count = 0
    failure_count = 0
    
    for i, pdf_path in enumerate(pdf_files):
        print(f"Processing {i+1}/{len(pdf_files)}: {os.path.basename(pdf_path)}...")
        
        # Use the comprehensive approach
        result = process_pdf_comprehensive(pdf_path)
        results.append(result)
        
        if result['Period 1 NAV'] is not None and result['Period 2 NAV'] is not None:
            success_count += 1
            print(f"  ✓ Success! Method: {result['Extraction Method']} (Confidence: {result['Confidence']:.2f})")
            print(f"    {result['Period 1 Label']}: {result['Period 1 NAV']}")
            print(f"    {result['Period 2 Label']}: {result['Period 2 NAV']}")
        else:
            failure_count += 1
            print(f"  ✗ Failed to extract NAV values")
    
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
    print(f"Success: {success_count}/{len(pdf_files)} ({success_count/len(pdf_files)*100:.1f}%)")
    print(f"Failed: {failure_count}/{len(pdf_files)} ({failure_count/len(pdf_files)*100:.1f}%)")
    
    # Show extraction method statistics
    method_stats = df['Extraction Method'].value_counts()
    print("\nExtraction Method Statistics:")
    for method, count in method_stats.items():
        print(f"  {method}: {count} ({count/len(pdf_files)*100:.1f}%)")


if __name__ == "__main__":
    main()
