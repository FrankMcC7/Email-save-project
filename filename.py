import os
import pandas as pd
import sys

def main():
    # Check if a directory was provided as a command-line argument
    if len(sys.argv) > 1:
        directory = sys.argv[1]
    else:
        directory = '.'

    # Verify if the provided directory exists
    if not os.path.isdir(directory):
        print(f"The directory {directory} does not exist.")
        sys.exit(1)

    # List all files in the directory
    try:
        files = os.listdir(directory)
    except Exception as e:
        print(f"An error occurred while listing files in {directory}: {e}")
        sys.exit(1)

    # Filter out directories, only include files
    file_list = [f for f in files if os.path.isfile(os.path.join(directory, f))]

    # Create a DataFrame from the list of files
    df = pd.DataFrame(file_list, columns=['File Name'])

    # Specify the output Excel file path
    output_file = os.path.join(directory, 'filenames.xlsx')

    # Write the DataFrame to an Excel file
    try:
        df.to_excel(output_file, index=False)
        print(f"File names have been written to {output_file}")
    except Exception as e:
        print(f"An error occurred while writing to Excel: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
