import pandas as pd

# Load the existing and new data files into pandas DataFrames
existing_df = pd.read_excel('existing_file.xlsx')
new_df = pd.read_excel('new_file.xlsx')

# Merge the dataframes
merged_df = pd.merge(existing_df, new_df, on='aa', suffixes=('_existing', '_new'))

# List to store results of changes
changes = []

# Iterate over each row in the merged dataframe
for index, row in merged_df.iterrows():
    hfad_fund_coperid = row['aa']
    row_changes = {'aa': hfad_fund_coperid}
    
    # Check each field if it has changed
    for col in ['aa', 'bb', 'cc', 'dd', 'ee']:
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
changes_df.to_excel('changes_report.xlsx', index=False)

print("Changes report generated: 'changes_report.xlsx'")
