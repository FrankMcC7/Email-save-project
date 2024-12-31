#!/usr/bin/env python3
import xml.etree.ElementTree as ET
import pandas as pd
import os

# Adjust this to your GLEIF XML file
XML_FILE = "GLEIF_LEI_RECORDS.xml"
# Adjust this to your Excel file
EXCEL_FILE = "LEI.xlsx"
# Output file
OUTPUT_EXCEL = "updated_LEI.xlsx"

# Namespace in GLEIF XML (verify by inspecting the root element or your file's schema)
NS = {"lei": "http://www.gleif.org/data/schema/leidata/2016"}

def parse_gleif_xml_to_dict(xml_file):
    """
    Parse the entire GLEIF XML file in memory, returning a dictionary:
      { legalName.lower(): LEI }
    for exact matching of names.
    """
    print(f"Parsing XML file: {xml_file} (this could take a few minutes)...")
    
    # We can either do a one-shot parse with ET.parse() or iterparse().
    # ET.parse() loads the entire tree at once, which might be memory-intensive but simpler to write.
    # If the file is truly huge, iterparse() is sometimes better. 
    # However, since we're storing everything in memory eventually, 
    # parse() won't be much different in total memory usage.

    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    # We'll store legalName -> LEI in a dictionary
    lei_dict = {}

    # Find all <LEIRecord> elements. 
    # The GLEIF XML typically has many such elements under the root.
    # Check the correct tag name in your file's structure.
    for record_elem in root.findall(".//lei:LEIRecord", NS):
        lei_elem = record_elem.find(".//lei:LEI", NS)
        name_elem = record_elem.find(".//lei:Entity/lei:LegalName", NS)
        
        if lei_elem is not None and name_elem is not None:
            lei_code = lei_elem.text.strip()
            legal_name = name_elem.text.strip()
            lei_dict[legal_name.lower()] = lei_code

    print(f"Parsed {len(lei_dict)} LEI records into dictionary.")
    return lei_dict


def main():
    # 1. Parse the GLEIF XML into a dictionary in memory
    if not os.path.exists(XML_FILE):
        print(f"XML file '{XML_FILE}' not found.")
        return
    lei_dict = parse_gleif_xml_to_dict(XML_FILE)
    
    # 2. Read your Excel file with fund names
    if not os.path.exists(EXCEL_FILE):
        print(f"Excel file '{EXCEL_FILE}' not found.")
        return
    
    print(f"Reading Excel file: {EXCEL_FILE}")
    df = pd.read_excel(EXCEL_FILE)

    # 3. Check if a "Fund Name" column exists
    fund_name_column = "Fund Name"
    if fund_name_column not in df.columns:
        print(f"Error: Column '{fund_name_column}' not found in Excel.")
        return

    # 4. Create or ensure 'LEI' column is present
    if "LEI" not in df.columns:
        df["LEI"] = ""
    
    # 5. For each row, do an exact dictionary lookup
    total_rows = len(df)
    print(f"Processing {total_rows} rows...")
    
    matched_count = 0
    missing_count = 0
    
    for idx, row in df.iterrows():
        fund_name = row[fund_name_column]
        # Basic check if fund_name is valid
        if pd.isna(fund_name) or not isinstance(fund_name, str) or fund_name.strip() == "":
            df.at[idx, "LEI"] = "Empty fund name"
            missing_count += 1
            continue
        
        # Exact match in dictionary
        lei_result = lei_dict.get(fund_name.lower(), "No LEI found")
        df.at[idx, "LEI"] = lei_result
        if lei_result == "No LEI found":
            missing_count += 1
        else:
            matched_count += 1
    
    # 6. Save the updated Excel
    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"\nLookups complete!")
    print(f"Matched: {matched_count}, Not found: {missing_count}")
    print(f"Results saved to '{OUTPUT_EXCEL}'")


if __name__ == "__main__":
    main()
