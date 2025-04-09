#!/usr/bin/env python3
"""
File Sorter

This script sorts and groups files by their extension type and places them in 
respective folders. Existing folders within the target directory remain untouched.

Usage:
    python file_sorter.py /path/to/directory
"""

import os
import sys
import shutil
from pathlib import Path

def get_file_type_folder(file_extension):
    """
    Determine the appropriate folder name based on file extension.
    """
    extension = file_extension.lower().strip('.')
    
    # Define mappings of extensions to folder names
    extension_mappings = {
        # Documents
        'pdf': 'Documents',
        'doc': 'Documents',
        'docx': 'Documents',
        'txt': 'Documents',
        'rtf': 'Documents',
        'odt': 'Documents',
        'xls': 'Documents',
        'xlsx': 'Documents',
        'ppt': 'Documents',
        'pptx': 'Documents',
        'csv': 'Documents',
        
        # Images
        'jpg': 'Images',
        'jpeg': 'Images',
        'png': 'Images',
        'gif': 'Images',
        'bmp': 'Images',
        'svg': 'Images',
        'tiff': 'Images',
        'webp': 'Images',
        
        # Audio
        'mp3': 'Audio',
        'wav': 'Audio',
        'ogg': 'Audio',
        'flac': 'Audio',
        'aac': 'Audio',
        'm4a': 'Audio',
        
        # Video
        'mp4': 'Videos',
        'avi': 'Videos',
        'mkv': 'Videos',
        'mov': 'Videos',
        'wmv': 'Videos',
        'flv': 'Videos',
        
        # Archives
        'zip': 'Archives',
        'rar': 'Archives',
        '7z': 'Archives',
        'tar': 'Archives',
        'gz': 'Archives',
        
        # Code
        'py': 'Code',
        'js': 'Code',
        'html': 'Code',
        'css': 'Code',
        'java': 'Code',
        'c': 'Code',
        'cpp': 'Code',
        'h': 'Code',
        'php': 'Code',
        'rb': 'Code',
        'go': 'Code',
        'json': 'Code',
        'xml': 'Code',
        
        # Executables
        'exe': 'Executables',
        'msi': 'Executables',
        'app': 'Executables',
        'bat': 'Executables',
        'sh': 'Executables',
        
        # Fonts
        'ttf': 'Fonts',
        'otf': 'Fonts',
        'woff': 'Fonts',
        'woff2': 'Fonts',
    }
    
    # Return the mapped folder name or 'Other' if not found
    return extension_mappings.get(extension, 'Other')

def sort_files(directory_path):
    """
    Sort files in the given directory into folders based on file type.
    Existing folders remain untouched.
    """
    try:
        # Convert to Path object and resolve to absolute path
        directory = Path(directory_path).resolve()
        
        if not directory.exists():
            print(f"Error: Directory '{directory}' does not exist.")
            return False
        
        if not directory.is_dir():
            print(f"Error: '{directory}' is not a directory.")
            return False
            
        print(f"Sorting files in: {directory}")
        
        # Get all items in the directory
        items = list(directory.iterdir())
        
        # Track statistics
        stats = {
            'files_moved': 0,
            'folders_created': set(),
            'skipped_folders': 0,
            'errors': 0
        }
        
        # Process each item
        for item in items:
            # Skip directories
            if item.is_dir():
                stats['skipped_folders'] += 1
                continue
                
            # Get the file extension without the dot
            file_extension = item.suffix.lower()
            
            # Skip files with no extension
            if not file_extension:
                folder_name = "No_Extension"
            else:
                # Determine the folder based on file extension
                folder_name = get_file_type_folder(file_extension)
                
            # Create the destination folder if it doesn't exist
            destination_folder = directory / folder_name
            
            if not destination_folder.exists():
                destination_folder.mkdir()
                stats['folders_created'].add(folder_name)
                print(f"Created folder: {folder_name}")
                
            # Move the file to the destination folder
            destination_file = destination_folder / item.name
            
            # Handle name conflicts
            if destination_file.exists():
                base_name = destination_file.stem
                extension = destination_file.suffix
                counter = 1
                
                # Try adding numbers until we find a name that doesn't exist
                while True:
                    new_name = f"{base_name}_{counter}{extension}"
                    destination_file = destination_folder / new_name
                    if not destination_file.exists():
                        break
                    counter += 1
            
            try:
                # Move the file
                shutil.move(str(item), str(destination_file))
                stats['files_moved'] += 1
                print(f"Moved: {item.name} -> {folder_name}/{destination_file.name}")
            except Exception as e:
                stats['errors'] += 1
                print(f"Error moving {item.name}: {e}")
                
        # Print summary
        print("\nSummary:")
        print(f"Files moved: {stats['files_moved']}")
        print(f"Folders created: {len(stats['folders_created'])}")
        print(f"Existing folders skipped: {stats['skipped_folders']}")
        print(f"Errors: {stats['errors']}")
        
        return True
        
    except Exception as e:
        print(f"An error occurred: {e}")
        return False

def main():
    """
    Main function to handle command line arguments and execute the sorting.
    """
    # Check if a directory path was provided
    if len(sys.argv) != 2:
        print("Usage: python file_sorter.py /path/to/directory")
        return
    
    # Get the directory path from command line argument
    directory_path = sys.argv[1]
    
    # Sort files in the directory
    sort_files(directory_path)
    
if __name__ == "__main__":
    main()
