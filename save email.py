def save_email(item, save_path, special_case):
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    
    valid_extensions = ('.xlsx', '.xls', '.csv', '.pdf', '.doc', '.docx')
    if special_case and special_case.lower() == 'yes' and item.Attachments.Count > 0:
        for attachment in item.Attachments:
            # Only consider attachments with specific file types
            if attachment.FileName.lower().endswith(valid_extensions):
                filename = sanitize_filename(os.path.splitext(attachment.Filename)[0])  # Remove extension
                filename = f"{filename}.msg"
                break
        else:
            # If no valid attachment is found, fallback to using the subject
            filename = f"{sanitize_filename(item.Subject)}.msg"
    else:
        filename = f"{sanitize_filename(item.Subject)}.msg"

    # Ensure the combined length of the save path and filename does not exceed 255 characters
    full_path = os.path.join(save_path, filename)
    if len(full_path) > 255:
        max_filename_length = 255 - len(save_path) - 1  # -1 for the path separator
        filename = filename[:max_filename_length]
        full_path = os.path.join(save_path, filename)

    item.SaveAs(full_path, 3)
    return filename
