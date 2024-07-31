import requests
from requests_ntlm import HttpNtlmAuth
import pandas as pd

# Authentication details
username = 'your_username'
password = 'your_password'
site_url = 'https://sharepoint.bankofamerica.com/sites/WCQATeam/HFTransparency'
parent_list_name = 'Hedge Fund Client Risk Disclosure Data Parent'
child_list_name = 'Child List Name'  # Replace with the actual child list name
lookup_field_name = 'ParentLookupField'  # Replace with the actual lookup field name in the child list
parent_id = 1  # Replace with the ID of the parent item you are interested in

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

# Get child list items for the specified parent item
child_list_info_url = f"{site_url}/_api/web/lists/GetByTitle('{child_list_name}')/items?$filter={lookup_field_name}Id eq {parent_id}"
response = session.get(child_list_info_url, headers={"Accept": "application/json;odata=verbose"})
if response.status_code == 200:
    child_items = response.json()['d']['results']
    print(f"Child records for Parent ID {parent_id}:")
    for child_item in child_items:
        print(child_item)  # Modify to print specific child fields as needed
else:
    print(f"Failed to retrieve child items for Parent ID {parent_id}:", response.text)
