import requests
from requests_ntlm import HttpNtlmAuth
import json
import csv
import getpass
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Prompt user for authentication details
username = input("Enter your username: ")
password = getpass.getpass("Enter your password: ")
site_url = 'abc'
child_list_name = 'abc'  # Replace with the actual name of the child list

# CSV file containing the data
csv_file_path = 'data.csv'  # Replace with the path to your CSV file

def get_request_digest(session, site_url):
    """Get the request digest token required for making POST requests."""
    digest_url = f"{site_url}/_api/contextinfo"
    response = session.post(digest_url, headers={"Accept": "application/json;odata=verbose"})
    
    if response.status_code == 200:
        return response.json()['d']['GetContextWebInformation']['FormDigestValue']
    else:
        logging.error(f"Failed to retrieve request digest token: {response.status_code} - {response.text}")
        raise Exception("Could not retrieve request digest token.")

def create_list_item(session, site_url, child_list_name, data, request_digest):
    """Create a new item in the specified SharePoint list."""
    add_item_url = f"{site_url}/_api/web/lists/GetByTitle('{child_list_name}')/items"
    
    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": request_digest
    }
    
    response = session.post(add_item_url, headers=headers, data=json.dumps(data))
    
    if response.status_code == 201:  # 201 Created
        logging.info(f"New entry successfully created: {data}")
    else:
        logging.error(f"Failed to create a new entry: {response.status_code} - {response.text}")

def process_csv_and_create_items(session, site_url, child_list_name, csv_file_path, request_digest):
    """Read the CSV file and create items in the SharePoint list based on the data."""
    try:
        with open(csv_file_path, mode='r') as csvfile:
            reader = csv.DictReader(csvfile)
            
            for row in reader:
                # Prepare the data for the new entry using the internal names of the fields
                new_entry_data = {
                    'ID': row.get('ID'),
                    'Name': row.get('Name'),
                    'rID': row.get('ID'),
                    'Name': row.get('Name'),
                    'Month': row.get('Month'),
                    'Level': row.get('Level'),
                    'Received': row.get('Received'),
                    'Action': row.get('Action'),
                    'ActionDate': row.get('ActionDate') if row.get('ActionDate') else None
                }
                
                # Log the data being sent for debugging purposes
                logging.debug(f"Creating item with data: {new_entry_data}")
                
                # Create the list item
                create_list_item(session, site_url, child_list_name, new_entry_data, request_digest)
    except FileNotFoundError:
        logging.error(f"CSV file not found: {csv_file_path}")
    except Exception as e:
        logging.error(f"Error processing CSV file: {e}")

def main():
    # Create a session and authenticate using NTLM
    session = requests.Session()
    session.auth = HttpNtlmAuth(username, password)
    
    try:
        # Get the request digest token
        request_digest = get_request_digest(session, site_url)
        
        # Process the CSV and create items in the SharePoint list
        process_csv_and_create_items(session, site_url, child_list_name, csv_file_path, request_digest)
        
    except Exception as e:
        logging.error(f"An error occurred during execution: {e}")

if __name__ == "__main__":
    main()
