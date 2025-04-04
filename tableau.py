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
                columns = []
                
                # Try to get columns from metadata records
                metadata_records = datasource.findall(f".//{namespace}metadata-record[@class='column']")
                if metadata_records:
                    for record in metadata_records:
                        local_name_elem = record.find(f".//{namespace}local-name")
                        if local_name_elem is not None and local_name_elem.text:
                            columns.append(local_name_elem.text)
                
                # If no columns found in metadata, try to extract from the first tuple
                if not columns:
                    first_tuple = inline_element.find(f".//{namespace}tuple")
                    if first_tuple is not None:
                        for i, _ in enumerate(first_tuple.findall(f".//{namespace}value")):
                            columns.append(f"Column_{i+1}")
                
                # Extract data rows
                rows = []
                for row in inline_element.findall(f".//{namespace}tuple"):
                    data_row = []
                    for val in row.findall(f".//{namespace}value"):
                        data_row.append(val.text if val.text else "")
                    if data_row:
                        rows.append(data_row)
                
                # Write to CSV file
                if rows:  # Even if no columns, we can still write the data
                    csv_path = os.path.join(output_dir, f"{ds_name}_embedded.csv")
                    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                        writer = csv.writer(csvfile)
                        if columns:
                            writer.writerow(columns)
                        writer.writerows(rows)
                    
                    csv_files.append(csv_path)
                    logger.info(f"Created CSV file: {csv_path}")
    
    return csv_files

def extract_hyper_data(extract_dir, output_dir):
    """
    Extract data from .hyper files and convert to CSV
    
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
                    
                    # Check if tableauhyperapi is available
                    try:
                        from tableauhyperapi import HyperProcess, Connection, Telemetry, CreateMode, TableName, SqlType, TableDefinition, Inserter
                        
                        with HyperProcess(telemetry=Telemetry.SEND_USAGE_DATA_TO_TABLEAU) as hyper:
                            with Connection(endpoint=hyper.endpoint, database=file_path) as connection:
                                # Get all table names with proper schema handling
                                try:
                                    # First try to get schemas
                                    schemas = connection.catalog.get_schema_names()
                                    
                                    all_tables = []
                                    for schema in schemas:
                                        schema_tables = connection.catalog.get_table_names(schema)
                                        all_tables.extend(schema_tables)
                                        
                                    table_names = all_tables
                                except Exception as schema_error:
                                    logger.warning(f"Error getting schemas: {schema_error}")
                                    # Try with default "public" schema
                                    try:
                                        table_names = connection.catalog.get_table_names("public")
                                    except Exception as public_error:
                                        logger.error(f"Error getting tables from public schema: {public_error}")
                                        # Create a placeholder file with information
                                        csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}_extract_error.csv")
                                        with open(csv_path, 'w') as f:
                                            f.write(f"# Data from {file_path}\n")
                                            f.write("# Error extracting table names from hyper file\n")
                                            f.write(f"# Error: {schema_error}\n")
                                        csv_files.append(csv_path)
                                        continue
                                
                                logger.info(f"Tables in {file}: {table_names}")
                                
                                # Process each table
                                for table_name in table_names:
                                    table_csv_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(file))[0]}_{table_name.name}.csv")
                                    
                                    try:
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
                                                    formatted_row = []
                                                    for item in row:
                                                        if item is None:
                                                            formatted_row.append("")
                                                        else:
                                                            formatted_row.append(str(item))
                                                    writer.writerow(formatted_row)
                                        
                                        csv_files.append(table_csv_path)
                                        logger.info(f"Created CSV file: {table_csv_path}")
                                    except Exception as table_error:
                                        logger.error(f"Error processing table {table_name}: {table_error}")
                    
                    except ImportError:
                        logger.warning("tableauhyperapi not installed. Install with 'pip install tableauhyperapi' to extract data from .hyper files.")
                        # Create a placeholder file with information
                        csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}_missing_library.csv")
                        with open(csv_path, 'w') as f:
                            f.write(f"# Data from {file_path}\n")
                            f.write("# To extract data from .hyper files, install tableauhyperapi: pip install tableauhyperapi\n")
                        csv_files.append(csv_path)
                    except Exception as e:
                        logger.error(f"Error processing hyper file {file_path}: {e}")
                        # Create a placeholder file with error information
                        csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}_error.csv")
                        with open(csv_path, 'w') as f:
                            f.write(f"# Data from {file_path}\n")
                            f.write(f"# Error: {str(e)}\n")
                        csv_files.append(csv_path)
                
                elif file.endswith('.tde'):
                    logger.info(f"Found TDE file: {file_path}")
                    
                    # Try to use dataextract if available
                    try:
                        import dataextract as tde
                        
                        # Open the TDE file
                        extract = tde.Extract(file_path)
                        
                        # Get all table names
                        table_count = extract.getTableCount()
                        
                        for i in range(table_count):
                            table = extract.openTable(i)
                            
                            # Get table schema
                            schema = table.getTableDefinition()
                            column_count = schema.getColumnCount()
                            
                            # Get column names and types
                            columns = []
                            for j in range(column_count):
                                col_name = schema.getColumnName(j)
                                columns.append(col_name if col_name else f"Column_{j}")
                            
                            # Write to CSV
                            csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}{'_' + str(i) if table_count > 1 else ''}.csv")
                            
                            with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                                writer = csv.writer(csvfile)
                                writer.writerow(columns)
                                
                                # Write data rows
                                for row in table:
                                    data_row = []
                                    for j in range(column_count):
                                        value = row.getValue(j)
                                        data_row.append(str(value) if value is not None else "")
                                    writer.writerow(data_row)
                            
                            csv_files.append(csv_path)
                            logger.info(f"Created CSV file: {csv_path}")
                        
                        extract.close()
                    
                    except ImportError:
                        logger.warning("dataextract not installed. TDE files require the Tableau SDK.")
                        # Create a placeholder file
                        csv_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}_tde.csv")
                        with open(csv_path, 'w') as f:
                            f.write(f"# Data from {file_path}\n")
                            f.write("# TDE files require the Tableau SDK (dataextract)\n")
                            f.write("# This SDK is deprecated by Tableau\n")
                        csv_files.append(csv_path)
                    except Exception as e:
                        logger.error(f"Error processing TDE file {file_path}: {e}")
    
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
    extract_csvs = extract_hyper_data(extract_dir, output_dir)
    csv_files.extend(extract_csvs)
    
    # Clean up temporary files if requested
    if cleanup:
        logger.info(f"Cleaning up temporary directory: {extract_dir}")
        shutil.rmtree(extract_dir)
    
    return csv_files

# Example usage
if __name__ == "__main__":
    # Specify your file path here
    twbx_file_path = "C:/dhinka/RFAD Alerts.twbx"  # Your file path
    
    # Specify output directory
    output_dir = "C:/dhinka/csv_output"  # Your output directory
    
    # Whether to clean up temporary files
    cleanup = True
    
    print(f"Processing file: {twbx_file_path}")
    print(f"Output files will be saved to: {output_dir}")
    
    try:
        csv_files = process_twbx_to_csv(twbx_file_path, output_dir, cleanup)
        
        if csv_files:
            print(f"\nSuccessfully extracted {len(csv_files)} CSV files:")
            for csv_file in csv_files:
                print(f" - {csv_file}")
        else:
            print("\nNo CSV files were generated. This could be due to:")
            print(" - The workbook doesn't contain data extracts or embedded data")
            print(" - The data is stored in an unsupported format")
            print(" - The workbook uses live connections to a database")
            print("\nFor live connections, you'll need to access the original data source.")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()