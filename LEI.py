#!/usr/bin/env python3
import pandas as pd
import requests
import time
from datetime import datetime
import logging
import os

# Set up logging with more detailed format
logging.basicConfig(
    filename=f'lei_lookup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def validate_file_exists(filename):
    """Check if the Excel file exists."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Excel file '{filename}' not found in current directory")

def validate_excel_structure(df, fund_name_column):
    """Validate that the required column exists and the DataFrame isn't empty."""
    if fund_name_column not in df.columns:
        raise ValueError(
            f"Column '{fund_name_column}' not found in Excel file. "
            f"Available columns are: {', '.join(df.columns)}"
        )
    if df.empty:
        raise ValueError("Excel file is empty")

def search_lei(fund_name):
    """
    Search for LEI using GLEIF API with improved error handling.
    Returns:
        - The LEI (string) if found
        - 'No LEI found' if no LEI is found
        - 'Error: ...' if there's an exception
        - 'Empty fund name' if the fund name is blank
        - 'Invalid fund name format' if the fund name isn't a string
    """
    if not isinstance(fund_name, str):
        return "Invalid fund name format"
    if pd.isna(fund_name) or fund_name.strip() == "":
        return "Empty fund name"
        
    base_url = "https://api.gleif.org/api/v1/fuzzycompletions"
    params = {
        "field": "entity.legalName",
        "q": fund_name.strip()
    }

    try:
        # Rate limit to avoid hitting the API too quickly
        time.sleep(1)
        
        response = requests.get(base_url, params=params, timeout=10)  # 10-second timeout
        response.raise_for_status()  # Raise for 4XX/5XX responses
        
        data = response.json()
        # If data is returned and there's at least one match
        if data and 'data' in data and len(data['data']) > 0:
            return data['data'][0]['lei']
        return "No LEI found"
        
    except requests.exceptions.Timeout:
        error_msg = f"Timeout while searching for {fund_name}"
        logging.error(error_msg)
        return "Error: API timeout"
    except requests.exceptions.RequestException as e:
        # Handles all other requests exceptions
        error_msg = f"API error while searching for {fund_name}: {str(e)}"
        logging.error(error_msg)
        return "Error: API request failed"
    except Exception as e:
        # Catch any other unforeseen exceptions
        error_msg = f"Unexpected error while searching for {fund_name}: {str(e)}"
        logging.error(error_msg)
        return f"Error: {str(e)}"

def process_excel_file(input_file, sheet_name=0):
    """
    Read an Excel file, validate its structure, look up LEIs for each fund name,
    and save the updated file with a new or updated 'LEI' column.
    """
    try:
        # Validate the file
        validate_file_exists(input_file)
        
        # Read the Excel file
        print("Reading file:", input_file)
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # This should match the actual column name in your Excel
        fund_name_column = 'Fund Name'
        
        # Validate Excel structure
        validate_excel_structure(df, fund_name_column)
        
        # Create the LEI column if it doesn't exist
        if 'LEI' not in df.columns:
            df['LEI'] = ''
        
        total_funds = len(df)
        print(f"Processing {total_funds} funds...")
        logging.info(f"Starting processing of {total_funds} funds")
        
        successful_lookups = 0
        failed_lookups = 0
        
        for idx, row in df.iterrows():
            try:
                fund_name = row[fund_name_column]
                # Check if LEI is already filled
                if pd.isna(df.loc[idx, 'LEI']) or df.loc[idx, 'LEI'] == '':
                    print(f"Processing {idx + 1}/{total_funds}: {fund_name}")
                    lei_result = search_lei(fund_name)
                    df.loc[idx, 'LEI'] = lei_result
                    
                    if "Error" in lei_result or lei_result == "No LEI found":
                        failed_lookups += 1
                    else:
                        successful_lookups += 1
                    
                    # Save progress every 50 records
                    if (idx + 1) % 50 == 0:
                        output_file = "updated_" + input_file
                        df.to_excel(output_file, index=False)
                        print(f"Progress saved at record {idx + 1}")
                        logging.info(f"Progress saved at record {idx + 1}")
                        
            except Exception as e:
                error_msg = (
                    f"Error processing row {idx + 1} "
                    f"({fund_name}): {str(e)}"
                )
                logging.error(error_msg)
                print(error_msg)
                failed_lookups += 1
                continue
        
        # Final save of the Excel file
        output_file = "updated_" + input_file
        df.to_excel(output_file, index=False)
        
        summary = (
            f"\nProcessing complete:\n"
            f"- Total funds processed: {total_funds}\n"
            f"- Successful lookups: {successful_lookups}\n"
            f"- Failed lookups: {failed_lookups}\n"
            f"- Results saved to: {output_file}\n"
        )
        print(summary)
        logging.info(summary)
        
    except Exception as e:
        error_msg = f"Error processing file: {str(e)}"
        logging.error(error_msg)
        print(error_msg)
        # Re-raise to let the caller know
        raise

def main():
    """
    Main entry point for the script. 
    Updates 'input_file' below with your actual Excel file name.
    """
    try:
        # Update this with the actual name of your Excel file
        input_file = "LEI.xlsx"
        
        print("Starting LEI lookup process...")
        print("Current working directory:", os.getcwd())
        
        process_excel_file(input_file)
        
    except Exception as e:
        error_msg = f"Program failed: {str(e)}"
        print(error_msg)
        logging.error(error_msg)

if __name__ == "__main__":
    main()
