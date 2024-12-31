#!/usr/bin/env python3
import pandas as pd
import requests
import time
from datetime import datetime
import logging
import os

# Set up logging
logging.basicConfig(
    filename=f'lei_lookup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# GraphQL query with minimal filtering: 
# - We remove "entityStatus_in: [ACTIVE]" 
# - We keep "includes: { legalName: $searchString }" to find matches by legal name.
GLEIF_GRAPHQL_QUERY = """
query ($searchString: String!) {
  leiRecords(
    filter: {
      includes: {
        entity: {
          legalName: $searchString
        }
      }
    }
    first: 5
  ) {
    totalCount
    records {
      lei
      entity {
        legalName
      }
    }
  }
}
"""

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

def search_lei_graphql(fund_name, max_retries=3):
    """
    Search for LEI using GLEIF's GraphQL endpoint (minimal filtering).
    
    Args:
        fund_name (str): The fund name to search.
        max_retries (int): Number of retries on transient errors.
    
    Returns:
        str: The LEI if found, or a descriptive error/no result message.
    """
    if not isinstance(fund_name, str):
        return "Invalid fund name format"
    if pd.isna(fund_name) or fund_name.strip() == "":
        return "Empty fund name"
    
    url = "https://api.gleif.org/api/graphql"
    headers = {"Content-Type": "application/json"}
    payload = {
        "query": GLEIF_GRAPHQL_QUERY,
        "variables": {
            "searchString": fund_name.strip()
        }
    }

    for attempt in range(1, max_retries + 1):
        try:
            # Small delay to reduce the chance of rate-limit issues
            time.sleep(0.5)
            
            response = requests.post(url, json=payload, headers=headers, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            # Check for GraphQL-level errors
            if "errors" in data:
                logging.error(f"GLEIF GraphQL responded with errors for '{fund_name}': {data['errors']}")
                return "Error: GraphQL error"

            # Extract matching records
            records = data["data"]["leiRecords"]["records"]
            if len(records) > 0:
                # Return the first matched LEI
                return records[0]["lei"]
            else:
                return "No LEI found"

        except requests.exceptions.RequestException as e:
            # Catch timeouts, connection errors, etc.
            logging.warning(f"Request exception on attempt {attempt} for '{fund_name}': {e}")
            if attempt == max_retries:
                return "Error: API request failed"
            # else, retry
        except Exception as e:
            logging.error(f"Unexpected error for '{fund_name}': {e}")
            return f"Error: {str(e)}"

def process_excel_file(input_file, fund_name_column='Fund Name'):
    """
    Read an Excel file, validate structure, look up LEIs (minimal GraphQL filter),
    and save to updated file with a new 'LEI' column.
    """
    try:
        # 1. Check file existence
        validate_file_exists(input_file)
        
        # 2. Read Excel
        print(f"Reading file: {input_file}")
        df = pd.read_excel(input_file)
        
        # 3. Validate structure
        validate_excel_structure(df, fund_name_column)
        
        # 4. Create/ensure LEI column
        if 'LEI' not in df.columns:
            df['LEI'] = ''
        
        total_funds = len(df)
        print(f"Processing {total_funds} funds...")
        logging.info(f"Starting processing of {total_funds} funds (minimal GraphQL filter)")
        
        successful_lookups = 0
        failed_lookups = 0
        
        # 5. Iterate through each fund
        for idx, row in df.iterrows():
            fund_name = row[fund_name_column]
            current_lei = row['LEI']
            
            # Only look up if LEI isn't already set or indicates an error
            if pd.isna(current_lei) or current_lei == '' or "Error" in current_lei:
                print(f"Processing {idx + 1}/{total_funds}: {fund_name}")
                lei_result = search_lei_graphql(fund_name)
                df.at[idx, 'LEI'] = lei_result
                
                if ("Error" in lei_result) or (lei_result == "No LEI found"):
                    failed_lookups += 1
                else:
                    successful_lookups += 1
                
                # Optionally save progress every 50 records
                if (idx + 1) % 50 == 0:
                    output_file = f"updated_{input_file}"
                    df.to_excel(output_file, index=False)
                    print(f"Progress saved at record {idx + 1}")
                    logging.info(f"Progress saved at record {idx + 1}")

        # 6. Final save
        output_file = f"updated_{input_file}"
        df.to_excel(output_file, index=False)
        
        # 7. Summary
        summary = (
            f"\nProcessing complete (minimal GraphQL filter):\n"
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
        raise

def main():
    """
    Main entry point. 
    Update 'input_file' for your Excel file name if needed.
    """
    try:
        input_file = "LEI.xlsx"  # Change to match your Excel filename
        print("Starting LEI lookup process (minimal GraphQL filter)...")
        print("Current working directory:", os.getcwd())
        
        process_excel_file(input_file)
        
    except Exception as e:
        error_msg = f"Program failed: {str(e)}"
        print(error_msg)
        logging.error(error_msg)

if __name__ == "__main__":
    main()
