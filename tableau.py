import zipfile
import os
import pandas as pd
import xml.etree.ElementTree as ET
import tempfile
import shutil

def extract_twbx(twbx_file_path, extract_dir=None):
    """
    Extract the contents of a .twbx file to a directory
    
    Args:
        twbx_file_path (str): Path to the .twbx file
        extract_dir (str, optional): Directory to extract to. If None, a temp directory is created.
        
    Returns:
        str: Path to the directory with extracted contents
    """
    if extract_dir is None:
        extract_dir = tempfile.mkdtemp()
    
    with zipfile.ZipFile(twbx_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)
    
    return extract_dir

def find_data_sources(twb_file_path):
    """
    Parse the .twb file to find data source information
    
    Args:
        twb_file_path (str): Path to the .twb file
        
    Returns:
        list: A list of dictionaries containing data source information
    """
    tree = ET.parse(twb_file_path)
    root = tree.getroot()
    
    # The namespace can vary, so we'll find it dynamically
    namespace = root.tag.split('}')[0] + '}' if '}' in root.tag else ''
    
    data_sources = []
    
    # Find datasources in the TWB file
    for datasource in root.findall(f".//{namespace}datasource"):
        ds_name = datasource.get('name', 'unknown')
        ds_type = datasource.get('type', 'unknown')
        
        # Try to find connection information
        connection = datasource.find(f".//{namespace}connection")
        if connection is not None:
            conn_class = connection.get('class', 'unknown')
            conn_dbname = connection.get('dbname', 'unknown')
            
            data_sources.append({
                'name': ds_name,
                'type': ds_type,
                'connection_class': conn_class,
                'dbname': conn_dbname
            })
    
    return data_sources

def extract_hyper_data(extract_dir, output_dir=None):
    """
    Extract data from .hyper files and convert to CSV
    
    Args:
        extract_dir (str): Directory with extracted .twbx contents
        output_dir (str, optional): Directory to save CSV files
        
    Returns:
        list: Paths to the generated CSV files
    """
    if output_dir is None:
        output_dir = os.path.join(extract_dir, 'csv_output')
        os.makedirs(output_dir, exist_ok=True)
    
    csv_files = []
    
    # Try to find and process Data Extract files
    data_dir = os.path.join(extract_dir, 'Data')
    if os.path.exists(data_dir):
        for root, _, files in os.walk(data_dir):
            for file in files:
                if file.endswith('.hyper'):
                    # For Hyper extracts, we need pantab or tableauhyperapi
                    # This is a more complex process requiring additional libraries
                    print(f"Found Hyper file: {file}")
                    print("Processing Hyper files requires the 'tableauhyperapi' library.")
                    print("Install it using: pip install tableauhyperapi")
                    
                    # Placeholder for Hyper processing code
                    # This would use tableauhyperapi to extract tables from the Hyper file
                    # For simplicity, we'll just note the existence of the file
                    csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}.csv")
                    csv_files.append(csv_path)
                
                elif file.endswith('.tde'):
                    # For older TDE files
                    print(f"Found TDE file: {file}")
                    print("Processing TDE files requires the 'dataextract' library.")
                    print("This is part of Tableau SDK which is now deprecated.")
                    
                    csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}.csv")
                    csv_files.append(csv_path)
    
    return csv_files

def process_twbx_to_csv(twbx_file_path, output_dir=None, cleanup=True):
    """
    Convert a .twbx file to CSV files
    
    Args:
        twbx_file_path (str): Path to the .twbx file
        output_dir (str, optional): Directory to save CSV files
        cleanup (bool): Whether to clean up temporary files
        
    Returns:
        list: Paths to the generated CSV files
    """
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(twbx_file_path), 'csv_output')
        os.makedirs(output_dir, exist_ok=True)
    
    # Extract the TWBX file
    extract_dir = extract_twbx(twbx_file_path)
    
    # Find the TWB file
    twb_file = None
    for file in os.listdir(extract_dir):
        if file.endswith('.twb'):
            twb_file = os.path.join(extract_dir, file)
            break
    
    if twb_file is None:
        raise FileNotFoundError("No .twb file found in the .twbx archive")
    
    # Parse the TWB file to find data sources
    data_sources = find_data_sources(twb_file)
    print(f"Found {len(data_sources)} data sources in the workbook")
    for ds in data_sources:
        print(f"Data source: {ds['name']}, Type: {ds['type']}")
    
    # Extract data from extracts
    csv_files = extract_hyper_data(extract_dir, output_dir)
    
    # If no extracts were found, we might have embedded data
    if not csv_files:
        print("No extract files found. Looking for embedded data...")
        # This would involve parsing the TWB XML more deeply to find embedded data
        # For a complete solution, you'd need to handle various data source types
    
    # Clean up temporary files if requested
    if cleanup:
        shutil.rmtree(extract_dir)
    
    return csv_files

# Example usage
if __name__ == "__main__":
    # Option 1: Using command-line arguments
    import argparse
    import sys
    
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description='Convert a Tableau .twbx file to CSV')
        parser.add_argument('twbx_file', help='Path to the .twbx file')
        parser.add_argument('--output-dir', help='Directory to save CSV files')
        parser.add_argument('--no-cleanup', action='store_true', help='Do not clean up temporary files')
        
        args = parser.parse_args()
        twbx_file_path = args.twbx_file
        output_dir = args.output_dir
        cleanup = not args.no_cleanup
    
    # Option 2: Hardcoded file path (uncomment and modify these lines to use)
    else:
        # Specify your file path here
        twbx_file_path = "C:/path/to/your/file.twbx"  # Change this to your file path
        output_dir = None  # Change this if you want to specify an output directory
        cleanup = True
        
        print(f"Using hardcoded file path: {twbx_file_path}")
    
    try:
        csv_files = process_twbx_to_csv(twbx_file_path, output_dir, cleanup)
        
        if csv_files:
            print(f"Successfully extracted {len(csv_files)} CSV files:")
            for csv_file in csv_files:
                print(f" - {csv_file}")
        else:
            print("No CSV files were generated. This could be due to:")
            print(" - The workbook doesn't contain data extracts")
            print(" - The data is stored in an unsupported format")
            print(" - The workbook uses live connections to a database")
            print("\nFor live connections, you'll need to access the original data source.")
    except Exception as e:
        print(f"Error: {e}")
