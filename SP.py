import requests
from requests_ntlm import HttpNtlmAuth
import pandas as pd

# Authentication details
username = 'your_username'
password = 'your_password'
site_url = 'url'
parent_list_name = 'list'
child_list_name = 'Child List Name'  # Replace with the actual child list name
parent_key_field = 'ParentKey'  # Replace with the actual key field name in the parent list
lookup_field_name = 'ParentKey'  # Replace with the actual lookup field name in the child list
parent_key_value = '12345'  # Replace with the value of the parent key you are interested in

# Create a session
session = requests.Session()
session.auth = HttpNtlmAuth(username, password)

# Get site information
site_info_url = f"{site_url}/_api/web"
response = session.get(site_info_url, headers={"Accept": "application/json;odata=verbose"})
if response.status_code == 200:
    site_info = response.json()
    print("Site Title:", site_info['d']['Title'])
else:
    print("Failed to retrieve site information:", response.text)

# Get parent list information
parent_list_info_url = f"{site_url}/_api/web/lists/GetByTitle('{parent_list_name}')"
response = session.get(parent_list_info_url, headers={"Accept": "application/json;odata=verbose"})
if response.status_code == 200:
    parent_list_info = response.json()
    print("Parent List Title:", parent_list_info['d']['Title'])
    print("Parent List Description:", parent_list_info['d']['Description'])

    # Get the specified parent item using the key field
    parent_item_url = f"{parent_list_info_url}/items?$filter={parent_key_field} eq '{parent_key_value}'"
    response = session.get(parent_item_url, headers={"Accept": "application/json;odata=verbose"})
    if response.status_code == 200:
        parent_items = response.json()['d']['results']
        if parent_items:
            print(f"Parent Item '{parent_key_field}':", parent_items[0][parent_key_field])

            # Get child list items for the specified parent key
            child_list_info_url = f"{site_url}/_api/web/lists/GetByTitle('{child_list_name}')/items?$filter={lookup_field_name} eq '{parent_key_value}'"
            response = session.get(child_list_info_url, headers={"Accept": "application/json;odata=verbose"})
            if response.status_code == 200:
                child_items = response.json()['d']['results']
                print(f"Child records for Parent Key {parent_key_value}:")
                for child_item in child_items:
                    print(child_item)  # Modify to print specific child fields as needed
            else:
                print(f"Failed to retrieve child items for Parent Key {parent_key_value}:", response.text)
        else:
            print(f"No parent item found with key '{parent_key_field}' = '{parent_key_value}'")
    else:
        print(f"Failed to retrieve parent item with key '{parent_key_field}' = '{parent_key_value}':", response.text)
else:
    print("Failed to retrieve parent list information:", response.text)
