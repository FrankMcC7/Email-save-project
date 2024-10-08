import pandas as pd

# Load the existing and new data files into pandas DataFrames
try:
    existing_df = pd.read_excel('existing_file.xlsx', engine='openpyxl')
    new_df = pd.read_excel('new_file.xlsx', engine='openpyxl')
    print("Files loaded successfully.")
except Exception as e:
    print(f"Error loading files: {e}")
    raise

# Merge the dataframes on
merged_df = pd.merge(existing_df, new_df, on='', suffixes=('_existing', '_new'))

# List to store results of changes
changes = []

# Iterate over each row in the merged dataframe
for index, row in merged_df.iterrows():
    hfad_fund_coperid = row['']
    row_changes = {'': }
    
    # Check each field if it has changed
    for col in ['', '', '', '', '']:
        existing_value = row[f'{col}_existing']
        new_value = row[f'{col}_new']
        
        if existing_value != new_value:
            row_changes[col] = {'existing': existing_value, 'new': new_value}
    
    # If any changes were found, add them to the list
    if len(row_changes) > 1:
        changes.append(row_changes)

# Convert the list of changes into a DataFrame for easy viewing or export
changes_df = pd.DataFrame(changes)

# Save the changes to a new Excel file
try:
    changes_df.to_excel('changes_report.xlsx', index=False)
    print("Changes report generated: 'changes_report.xlsx'")
except Exception as e:
    print(f"Error saving report: {e}")
    raise
