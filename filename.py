import os
import pandas as pd

# Function to extract filenames from the provided directory
def extract_filenames_to_excel(directory_path, output_excel):
    # Check if the provided path is a valid directory
    if not os.path.isdir(directory_path):
        print(f"The path '{directory_path}' is not a valid directory.")
        return

    # Extract filenames
    filenames = os.listdir(directory_path)

    # Create a DataFrame
    df = pd.DataFrame(filenames, columns=["Filename"])

    # Save to Excel
    df.to_excel(output_excel, index=False)
    print(f"Filenames have been extracted to '{output_excel}' successfully.")

# User inputs for directory path and output Excel file
directory_path = input("Enter the directory path: ")
output_excel = input("Enter the output Excel filename (e.g., filenames.xlsx): ")

# Extract filenames and save to Excel
extract_filenames_to_excel(directory_path, output_excel)
