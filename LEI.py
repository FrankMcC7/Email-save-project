import pandas as pd
import requests
import time
import os

def search_lei(fund_name):
    base_url = "https://api.gleif.org/api/v1/fuzzycompletions"
    params = {"field": "entity.legalName", "q": fund_name}
    
    try:
        time.sleep(1)
        response = requests.get(base_url, params=params, timeout=10)
        if response.status_code == 200:
            data = response.json()
            if data and 'data' in data and len(data['data']) > 0:
                return data['data'][0]['lei']
            return "No LEI found"
    except:
        return "Error: API request failed"

def process_excel_file(filename):
    try:
        df = pd.read_excel(filename)
        fund_name_column = 'Fund Name'
        
        if 'LEI' not in df.columns:
            df['LEI'] = ''
        
        for i in range(len(df)):
            fund_name = df.iloc[i][fund_name_column]
            print("Processing %d of %d" % (i+1, len(df)))
            lei = search_lei(fund_name)
            df.iloc[i]['LEI'] = lei
            
            if (i + 1) % 50 == 0:
                df.to_excel("updated_LEI.xlsx", index=False)
                print("Saved progress")
        
        df.to_excel("updated_LEI.xlsx", index=False)
        print("Complete")
        
    except Exception as e:
        print("Error: %s" % str(e))

if __name__ == "__main__":
    print("Starting")
    process_excel_file("LEI.xlsx")
