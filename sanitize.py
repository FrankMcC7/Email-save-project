def sanitize_filename(filename):
    # Step 1: Normalize Unicode characters
    filename = unicodedata.normalize('NFKD', filename).encode('ASCII', 'ignore').decode('ASCII')

    # Step 2: Remove forbidden characters in Windows filenames
    forbidden_chars = r'[<>:"/\\|?*\x00-\x1F]'
    filename = re.sub(forbidden_chars, '', filename)

    # Step 3: Replace problematic characters with underscores
    filename = re.sub(r'[^a-zA-Z0-9.\-_() ]', '_', filename)

    # Step 4: Collapse multiple underscores
    filename = re.sub(r'_+', '_', filename)

    # Step 5: Remove leading/trailing spaces and dots
    filename = filename.strip(' .')

    # Step 6: Ensure the filename doesn't start with a dash
    filename = filename.lstrip('-')

    # Step 7: Handle Windows reserved names
    reserved_names = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9']
    name_part = filename.split('.')[0].upper()
    if name_part in reserved_names:
        filename = '_' + filename

    # Step 8: Ensure the filename isn't empty
    if not filename:
        filename = '_unnamed_file'

    # Step 9: Truncate to 255 characters (maximum length for NTFS)
    return filename[:255]
