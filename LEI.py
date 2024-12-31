import pandas as pd
import requests
import time
from datetime import datetime
import logging

# Set up logging
logging.basicConfig(filename=f'lei_lookup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
                   level=logging.INFO,
                   format='%(asctime)s - %(message)s')

def search_lei(fund_name):
    """Search for LEI using GLEIF API"""
    base_url = "https://api.gleif.org/api/v1/fuzzycompletions"
    params = {
        "field": "entity.legalName",
        "q": fund_name
    }
    
    try:
        # Rate limiting - be respectful to the API
        time.sleep(1)
        
        response = requests.get(base_url, params=params)
        if response.status_code == 200:
            data = response.json()
            # Check if we got any matches
            if data and 'data' in data and len(data['data']) > 0:
                # Return the first match's LEI
                return data['data'][0]['lei']
            return "No LEI found"
    except Exception as e:
        logging.error(f"Error searching for {fund_name}: {str(e)}")
        return f"Error: {str(e)}"

def process_excel_file(input_file, sheet_name=0):
    """Process Excel file and add LEIs"""
    try:
        # Read the Excel file
        print(f"Reading file: {input_file}")
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Assume the column with fund names is called 'Fund Name'
        # Change this to match your column name
        fund_name_column = 'Fund Name'  # Update this to your column name
        
        # Create new column for LEIs if it doesn't exist
        if 'LEI' not in df.columns:
            df['LEI'] = ''
        
        total_funds = len(df)
        print(f"Processing {total_funds} funds...")
        
        # Process each fund
        for idx, row in df.iterrows():
            fund_name = row[fund_name_column]
            if pd.isna(df.loc[idx, 'LEI']) or df.loc[idx, 'LEI'] == '':
                print(f"Processing {idx + 1}/{total_funds}: {fund_name}")
                lei = search_lei(fund_name)
                df.loc[idx, 'LEI'] = lei
                
                # Save progress every 50 records
                if (idx + 1) % 50 == 0:
                    output_file = f"updated_{input_file}"
                    df.to_excel(output_file, index=False)
                    print(f"Progress saved at record {idx + 1}")
        
        # Save final results
        output_file = f"updated_{input_file}"
        df.to_excel(output_file, index=False)
        print(f"Processing complete. Results saved to: {output_file}")
        
    except Exception as e:
        logging.error(f"Error processing file: {str(e)}")
        print(f"An error occurred: {str(e)}")
        raise

def main():
    # Replace with your Excel file name
    input_file = "your_file.xlsx"
    
    try:
        process_excel_file(input_file)
    except Exception as e:
        print(f"Program failed: {str(e)}")
        logging.error(f"Program failed: {str(e)}")

if __name__ == "__main__":
    main()
