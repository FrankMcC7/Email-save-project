import pandas as pd
import requests
import time
from datetime import datetime
import logging
import os

# Set up logging
logging.basicConfig(
    filename='lei_lookup.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def search_lei(fund_name):
    base_url = "https://api.gleif.org/api/v1/fuzzycompletions"
    params = {
        "field": "entity.legalName",
        "q": fund_name
    }
    
    try:
        time.sleep(1)
        response = requests.get(base_url, params=params, timeout=10)
        if response.status_code == 200:
            data = response.json()
            if data and 'data' in data and len(data['data']) > 0:
                return data['data'][0]['lei']
            return "No LEI found"
    except Exception as e:
        logging.error("Error searching for %s: %s", fund_name, str(e))
        return "Error: API request failed"

def process_excel_file(input_file):
    try:
        # Check if file exists
        if not os.path.exists(input_file):
            print("File not found:", input_file)
            return

        # Read Excel file
        print("Reading file:", input_file)
        df = pd.read_excel(input_file)
        
        # Get fund name column
        fund_name_column = 'Fund Name'  # Change this to match your column name
        
        if fund_name_column not in df.columns:
            print("Column not found:", fund_name_column)
            return
            
        # Add LEI column
        if 'LEI' not in df.columns:
            df['LEI'] = ''
        
        # Process funds
        total = len(df)
        for i, row in df.iterrows():
            print("Processing", i+1, "of", total)
            
            if pd.isna(df.loc[i, 'LEI']) or df.loc[i, 'LEI'] == '':
                lei = search_lei(row[fund_name_column])
                df.loc[i, 'LEI'] = lei
                
                # Save every 50 records
                if (i + 1) % 50 == 0:
                    df.to_excel("updated_" + input_file, index=False)
                    print("Progress saved")
        
        # Save final results
        df.to_excel("updated_" + input_file, index=False)
        print("Complete - saved to: updated_" + input_file)
        
    except Exception as e:
        print("Error:", str(e))
        logging.error(str(e))

def main():
    input_file = "LEI.xlsx"
    print("Starting process")
    print("Working directory:", os.getcwd())
    process_excel_file(input_file)

if __name__ == "__main__":
    main()
