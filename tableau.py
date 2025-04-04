import zipfile
import os
import pandas as pd
import xml.etree.ElementTree as ET
import tempfile
import shutil
import csv
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

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
    
    logger.info(f"Extracting {twbx_file_path} to {extract_dir}")
    
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
    logger.info(f"Parsing {twb_file_path} to find data sources")
    
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

def find_embedded_data(twb_file_path, output_dir):
    """
    Find and extract embedded data from the .twb file
    
    Args:
        twb_file_path (str): Path to the .twb file
        output_dir (str): Directory to save CSV files
        
    Returns:
        list: Paths to the generated CSV files
    """
    logger.info("Looking for embedded data in the .twb file")
    
    csv_files = []
    tree = ET.parse(twb_file_path)
    root = tree.getroot()
    
    # The namespace can vary, so we'll find it dynamically
    namespace = root.tag.split('}')[0] + '}' if '}' in root.tag else ''
    
    # Look for embedded data in datasource elements
    for datasource in root.findall(f".//{namespace}datasource"):
        ds_name = datasource.get('name', 'unknown')
        
        # Look for inline data elements
        connection = datasource.find(f".//{namespace}connection")
        if connection is not None:
            inline_element = connection.find(f".//{namespace}inline")
            if inline_element is not None:
                logger.info(f"Found embedded data in datasource: {ds_name}")
                
                # Extract column names
                metadata_records = inline_element.findall(f".//{namespace}metadata-record")
                columns = []
                for record in metadata_records:
                    if record.get('class') == 'column':
                        local_name = None
                        for prop in record.findall(f".//{namespace}local-name"):
                            local_name = prop.text
                        if local_name:
                            columns.append(local_name)
                
                # Extract data rows
                rows = []
                for row in inline_element.findall(f".//{namespace}tuple"):
                    data_row = []
                    for val in row.findall(f".//{namespace}value"):
                        data_row.append(val.text if val.text else "")
                    if data_row:
                        rows.append(data_row)
                
                # Write to CSV file
                if columns and rows:
                    csv_path = os.path.join(output_dir, f"{ds_name}_embedded.csv")
                    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                        writer = csv.writer(csvfile)
                        writer.writerow(columns)
                        writer.writerows(rows)
                    
                    csv_files.append(csv_path)
                    logger.info(f"Created CSV file: {csv_path}")
    
    return csv_files

def extract_extract_files(extract_dir, output_dir):
    """
    Process extract files (.hyper, .tde) and convert to CSV
    
    Args:
        extract_dir (str): Directory with extracted .twbx contents
        output_dir (str): Directory to save CSV files
        
    Returns:
        list: Paths to the generated CSV files
    """
    logger.info("Looking for data extract files (.hyper, .tde)")
    
    csv_files = []
    
    # Try to find and process Data Extract files
    data_dir = os.path.join(extract_dir, 'Data')
    if os.path.exists(data_dir):
        for root, _, files in os.walk(data_dir):
            for file in files:
                file_path = os.path.join(root, file)
                
                if file.endswith('.hyper'):
                    logger.info(f"Found Hyper file: {file_path}")
                    csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}.csv")
                    
                    # Check if tableauhyperapi is available
                    try:
                        from tableauhyperapi import HyperProcess, Connection, Telemetry, CreateMode, TableName, SqlType, TableDefinition, Inserter
                        from tableauhyperapi import escape_string_literal
                        
                        # Process the hyper file
                        with HyperProcess(telemetry=Telemetry.SEND_USAGE_DATA_TO_TABLEAU) as hyper:
                            with Connection(endpoint=hyper.endpoint, database=file_path) as connection:
                                # Get all table names
                                table_names = connection.catalog.get_table_names()
                                logger.info(f"Tables in {file}: {table_names}")
                                
                                for table_name in table_names:
                                    table_csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}_{table_name.name}.csv")
                                    
                                    # Get table definition
                                    table_def = connection.catalog.get_table_definition(table_name)
                                    columns = [column.name for column in table_def.columns]
                                    
                                    # Query all rows
                                    query = f"SELECT * FROM {table_name}"
                                    with connection.execute_query(query) as result:
                                        with open(table_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                                            writer = csv.writer(csvfile)
                                            writer.writerow(columns)
                                            
                                            # Write data rows
                                            for row in result:
                                                writer.writerow(row)
                                    
                                    csv_files.append(table_csv_path)
                                    logger.info(f"Created CSV file: {table_csv_path}")
                    
                    except ImportError:
                        logger.warning("tableauhyperapi not installed. Install with 'pip install tableauhyperapi' to extract data from .hyper files.")
                        # Create a placeholder file with information
                        with open(csv_path, 'w') as f:
                            f.write(f"# Data from {file_path}\n")
                            f.write("# To extract data from .hyper files, install tableauhyperapi: pip install tableauhyperapi\n")
                        csv_files.append(csv_path)
                
                elif file.endswith('.tde'):
                    logger.info(f"Found TDE file: {file_path}")
                    csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}.csv")
                    
                    # Creating placeholder for TDE files (deprecated by Tableau)
                    with open(csv_path, 'w') as f:
                        f.write(f"# Data from {file_path}\n")
                        f.write("# TDE files are deprecated by Tableau and require the Tableau SDK\n")
                    
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
    if not os.path.exists(twbx_file_path):
        raise FileNotFoundError(f"File not found: {twbx_file_path}")
    
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(twbx_file_path), 'csv_output')
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    logger.info(f"Output directory: {output_dir}")
    
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
    logger.info(f"Found {len(data_sources)} data sources in the workbook")
    for ds in data_sources:
        logger.info(f"Data source: {ds['name']}, Type: {ds['type']}")
    
    # Extract data from the workbook
    csv_files = []
    
    # Look for embedded data in the TWB file
    embedded_csvs = find_embedded_data(twb_file, output_dir)
    csv_files.extend(embedded_csvs)
    
    # Look for extract files (.hyper, .tde)
    extract_csvs = extract_extract_files(extract_dir, output_dir)
    csv_files.extend(extract_csvs)
    
    # Clean up temporary files if requested
    if cleanup:
        logger.info(f"Cleaning up temporary directory: {extract_dir}")
        shutil.rmtree(extract_dir)
    
    return csv_files

# Example usage
if __name__ == "__main__":
    # Specify your file path here
    twbx_file_path = "C:/path/to/your/file.twbx"  # Change this to your file path
    
    # Specify output directory (optional)
    # If None, CSV files will be saved in a folder named 'csv_output' in the same directory as your .twbx file
    output_dir = None  
    
    # Alternatively, you can specify an exact output path:
    # output_dir = "C:/path/where/you/want/csv/files" 
    
    cleanup = True  # Set to False if you want to keep temporary extraction files
    
    print(f"Processing file: {twbx_file_path}")
    if output_dir:
        print(f"Output files will be saved to: {output_dir}")
    else:
        default_output = os.path.join(os.path.dirname(twbx_file_path), 'csv_output')
        print(f"Output files will be saved to: {default_output}")
    
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
    except FileNotFoundError as e:
        print(f"Error: {e}")
        print("Please check that the file path is correct and the file exists.")
    except Exception as e:
        print(f"Error: {e}")
