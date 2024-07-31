pip install Office365-REST-Python-Client

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import pandas as pd

# Authentication
site_url = 'url'
client_id = 'your_client_id'  # Typically provided by your IT department
client_secret = 'your_client_secret'  # Typically provided by your IT department

# Connect to SharePoint
ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))

# Retrieve and display site information
web = ctx.web
ctx.load(web)
ctx.execute_query()
print("Site Title:", web.properties['Title'])
print("Site URL:", site_url)

# List name
base_list_name = 'list name'

# Retrieve and display list information
target_list = ctx.web.lists.get_by_title(base_list_name)
ctx.load(target_list)
ctx.execute_query()
print("List Title:", target_list.properties['Title'])

# Load sample items from the list
items = target_list.items.top(5).get().execute_query()
for item in items:
    print(item.properties)

# Verify list name
response = input(f"Is the SharePoint list '{base_list_name}' correct? (yes/no): ").strip().lower()
if response != 'yes':
    print("Please verify your SharePoint site URL and list name.")
else:
    print("Proceeding with the update...")
    # Your update code goes here
