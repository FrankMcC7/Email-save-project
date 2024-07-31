pip install requests requests_ntlm

import requests
from requests_ntlm import HttpNtlmAuth
import pandas as pd

# Authentication details
username = 'your_username'
password = 'your_password'
site_url = 'url'
list_name = 'item name'

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

# Get list information
list_info_url = f"{site_url}/_api/web/lists/GetByTitle('{list_name}')"
response = session.get(list_info_url, headers={"Accept": "application/json;odata=verbose"})
if response.status_code == 200:
    list_info = response.json()
    print("List Title:", list_info['d']['Title'])
    print("List Description:", list_info['d']['Description'])

    # Get sample items from the list
    items_url = f"{list_info_url}/items?$top=5"
    response = session.get(items_url, headers={"Accept": "application/json;odata=verbose"})
    if response.status_code == 200:
        items = response.json()['d']['results']
        for item in items:
            print(item)
    else:
        print("Failed to retrieve list items:", response.text)
else:
    print("Failed to retrieve list information:", response.text)

# Verify list name
response = input(f"Is the SharePoint list '{list_name}' correct? (yes/no): ").strip().lower()
if response != 'yes':
    print("Please verify your SharePoint site URL and list name.")
else:
    print("Proceeding with the update...")
    # Your update code goes here
