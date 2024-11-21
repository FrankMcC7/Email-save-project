import os
import sys

def remove_empty_dirs(path):
    """Recursively delete empty folders within the provided directory."""
    for dirpath, dirnames, filenames in os.walk(path, topdown=False):
        # If the directory is empty (no subdirectories and no files)
        if not dirnames and not filenames:
            try:
                os.rmdir(dirpath)
                print(f"Removed empty directory: {dirpath}")
            except OSError as e:
                print(f"Error removing {dirpath}: {e}")

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python script_name.py <directory>")
        sys.exit(1)

    root_dir = sys.argv[1]
    if not os.path.isdir(root_dir):
        print(f"Error: {root_dir} is not a valid directory.")
        sys.exit(1)

    remove_empty_dirs(root_dir)