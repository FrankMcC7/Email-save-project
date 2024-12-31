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
    """Check if the Excel file exists"""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Excel file '{filename}' not found in current directory")

def validate_excel_structure(df, fund_name_column):
    """Validate Excel file structure"""
    if fund_name_column not in df.columns:
        raise ValueError(f"Column '{fund_name_column}' not found in Excel file. Available columns are: {', '.join(df.columns)}")
    
    # Check for empty dataframe
    if df.empty:
        raise ValueError("Excel file is empty")

def search_lei(fund_name):
    """Search for LEI using GLEIF API with improved error handling"""
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
        # Rate limiting
        time.sleep(1)
        
        response = requests.get(base_url, params=params, timeout=10)  # Added timeout
        response.raise_for_status()  # Will raise an exception for 4XX/5XX status codes
        
        data = response.json()
        if data and 'data' in data and len(data['data']) > 0:
            return data['data'][0]['lei']
        return "No LEI found"
        
    except requests.exceptions.Timeout:
        error_msg = f"Timeout while searching for {fund_name}"
        logging.error(error_msg)
        return "Error: API timeout"
    except requests.exceptions.RequestException as e:
        error_msg = f"API error while searching for {fund_name}: {str(e)}"
        logging.error(error_msg)
        return "Error: API request failed"
    except Exception as e:
        error_msg = f"Unexpected error while searching for {fund_name}: {str(e)}"
        logging.error(error_msg)
        return f"Error: {str(e)}"

def process_excel_file(input_file, sheet_name=0):
    """Process Excel file and add LEIs with improved error handling"""
    try:
        # Validate file exists
        validate_file_exists(input_file)
        
        # Read the Excel file
        print("Reading file:", input_file)
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Update this to match your column name
        fund_name_column = 'Fund Name'  
        
        # Validate Excel structure
        validate_excel_structure(df, fund_name_column)
        
        # Create new column for LEIs if it doesn't exist
        if 'LEI' not in df.columns:
            df['LEI'] = ''
        
        total_funds = len(df)
        print(f"Processing {total_funds} funds...")
        logging.info(f"Starting processing of {total_funds} funds")
        
        successful_lookups = 0
        failed_lookups = 0
        
        # Process each fund
        for idx, row in df.iterrows():
            try:
                fund_name = row[fund_name_column]
                if pd.isna(df.loc[idx, 'LEI']) or df.loc[idx, 'LEI'] == '':
                    print(f"Processing {idx + 1}/{total_funds}: {fund_name}")
                    lei = search_lei(fund_name)
                    df.loc[idx, 'LEI'] = lei
                    
                    if "Error" in lei or lei == "No LEI found":
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
                error_msg = f"Error processing row {idx + 1} ({fund_name}): {str(e)}"
                logging.error(error_msg)
                print(error_msg)
                failed_lookups += 1
                continue
        
        # Save final results
        output_file = "updated_" + input_file
        df.to_excel(output_file, index=False)
        
        # Log summary
        summary = f"""
        Processing complete:
        - Total funds processed: {total_funds}
        - Successful lookups: {successful_lookups}
        - Failed lookups: {failed_lookups}
        - Results saved to: {output_file}
        """
        print(summary)
        logging.info(summary)
        
    except Exception as e:
        error_msg = f"Error processing file: {str(e)}"
        logging.error(error_msg)
        print(error_msg)
        raise

def main():
    try:
        # Replace with your Excel file name
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
