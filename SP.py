from shareplum import Site
from shareplum import Office365
from shareplum.site import Version

# Authentication
username = 'your_username'
password = 'your_password'
site_url = 'https://your_sharepoint_site_url'
base_list_name = 'BaseDatabaseList'

# Connect to SharePoint
authcookie = Office365(site_url, username=username, password=password).GetCookies()
site = Site(site_url, version=Version.v365, authcookie=authcookie)

# Retrieve and display site information
print("Site Title:", site.info['Title'])
print("Site URL:", site_url)

# Retrieve and display list information
base_list = site.List(base_list_name)
list_info = base_list.GetListItems('All Items', rowlimit=5)  # Load a sample of 5 items
print("List Title:", base_list_name)
print("Sample Data from List:")
for item in list_info:
    print(item)

# Check if the base list name is correct
response = input(f"Is the SharePoint list '{base_list_name}' correct? (yes/no): ").strip().lower()
if response != 'yes':
    print("Please verify your SharePoint site URL and list name.")
else:
    print("Proceeding with the update...")
    # Your update code goes here
