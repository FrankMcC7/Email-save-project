import pandas as pd
import requests
import time
import os

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
        return "Error: API request failed"

def process_excel_file(input_file):
    try:
        print("Reading file...")
        df = pd.read_excel(input_file)
        
        fund_name_column = 'Fund Name'  # Change this to match your column name
        
        if fund_name_column not in df.columns:
            print("Column not found:", fund_name_column)
            return
            
        if 'LEI' not in df.columns:
            df['LEI'] = ''
        
        total = len(df)
        for i, row in df.iterrows():
            print("Processing {} of {}".format(i+1, total))
            
            if pd.isna(df.loc[i, 'LEI']) or df.loc[i, 'LEI'] == '':
                lei = search_lei(row[fund_name_column])
                df.loc[i, 'LEI'] = lei
                
                if (i + 1) % 50 == 0:
                    df.to_excel("updated_" + input_file, index=False)
                    print("Progress saved")
        
        df.to_excel("updated_" + input_file, index=False)
        print("Complete - saved to: updated_" + input_file)
        
    except Exception as e:
        print("Error:", str(e))

if __name__ == "__main__":
    input_file = "LEI.xlsx"
    print("Starting process")
    print("Working directory:", os.getcwd())
    process_excel_file(input_file)
