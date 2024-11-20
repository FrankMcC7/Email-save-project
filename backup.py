import os
import datetime
import pythoncom
import win32com.client as win32
import openpyxl
from openpyxl import Workbook

METRICS_FILE = "backup_metrics.xlsx"  # File to store backup metrics

def sanitize_filename(filename):
    """
    Remove or replace characters that are invalid in Windows filenames.
    """
    invalid_chars = '<>:"/\\|?*'
    filename = ''.join(c if c not in invalid_chars else '_' for c in filename)
    # Remove control characters but keep Unicode characters
    filename = ''.join(c for c in filename if c.isprintable() and (c.isalnum() or c in ' ._-'))
    return filename.strip()

def truncate_or_fallback_filename(save_path, subject, max_path_length=255):
    """
    Truncate the subject to ensure the total path length is less than 255 characters.
    Use a fallback filename if the path still exceeds the limit.
    """
    base_path_length = len(os.path.dirname(save_path)) + 1  # Include the directory path and slash
    max_filename_length = max_path_length - base_path_length - 4  # Reserve 4 for ".msg"

    # Truncate the subject if necessary
    if len(subject) > max_filename_length:
        subject = subject[:max_filename_length]

    full_path = os.path.join(os.path.dirname(save_path), f"{subject}.msg")
    if len(full_path) > max_path_length:
        # Use a fallback filename if truncation is not enough
        counter = 1
        while True:
            fallback_filename = f"Email_Subject_Changed_{counter}.msg"
            fallback_path = os.path.join(os.path.dirname(save_path), fallback_filename)
            if len(fallback_path) <= max_path_length and not os.path.exists(fallback_path):
                return fallback_filename
            counter += 1

    return f"{subject}.msg"

def save_email(item, save_path):
    """
    Save an email as a .msg file.
    """
    try:
        # Save the email as a .msg file
        item.SaveAs(save_path, 3)  # 3 refers to the MSG format
        print(f"Saved email: {save_path}")
        return True
    except Exception as e:
        print(f"Failed to save email '{item.Subject}': {str(e)}")
        return False

def log_metrics(date, total_emails, saved_emails, fallback_emails, processing_errors):
    """
    Log the backup metrics in an Excel workbook.
    """
    if not os.path.exists(METRICS_FILE):
        # Create a new workbook if it doesn't exist
        wb = Workbook()
        ws = wb.active
        ws.title = "Backup Metrics"
        ws.append(["Date", "Total Emails", "Saved Emails", "Fallback Emails", "Processing Errors"])
        wb.save(METRICS_FILE)

    # Open the workbook and append the metrics
    wb = openpyxl.load_workbook(METRICS_FILE)
    ws = wb["Backup Metrics"]
    ws.append([date, total_emails, saved_emails, fallback_emails, processing_errors])
    wb.save(METRICS_FILE)

def backup_shared_mailbox(mailbox_name, backup_root_directory, backup_dates):
    """
    Backs up all emails from the specified shared mailbox received on the backup_dates to the specified directory.
    """
    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Locate the shared mailbox by name
        shared_mailbox = None
        for folder in outlook.Folders:
            if folder.Name.lower() == mailbox_name.lower():
                shared_mailbox = folder
                break

        if not shared_mailbox:
            print(f"Could not find the shared mailbox: {mailbox_name}")
            print("Available mailboxes are:")
            for folder in outlook.Folders:
                print(f"- {folder.Name}")
            return

        # Get the Inbox folder of the shared mailbox
        inbox = shared_mailbox.Folders["Inbox"]

        # Process each date in the backup_dates list
        for backup_date_str in backup_dates:
            try:
                # Convert backup_date string to datetime object
                backup_date = datetime.datetime.strptime(backup_date_str, '%Y-%m-%d')
            except ValueError:
                print(f"Invalid date format: {backup_date_str}. Please use YYYY-MM-DD format.")
                continue

            # Format the date components for the directory structure
            year_str = backup_date.strftime('%Y')
            month_str = backup_date.strftime('%m-%B')
            date_str = backup_date.strftime('%d-%m-%Y')

            # Construct the save directory path
            save_directory = os.path.join(backup_root_directory, year_str, month_str, date_str)

            # Create the save directory if it doesn't exist
            if not os.path.exists(save_directory):
                try:
                    os.makedirs(save_directory)
                except Exception as e:
                    print(f"Failed to create directory '{save_directory}': {str(e)}")
                    continue  # Skip processing this date if directory cannot be created

            # Filter emails received on the backup date
            messages = inbox.Items

            # Set the date range for the filter (from start to end of the backup date)
            start_date = backup_date.strftime('%m/%d/%Y %I:%M %p')
            end_date = (backup_date + datetime.timedelta(days=1)).strftime('%m/%d/%Y %I:%M %p')

            # Enclose dates in # symbols for proper Outlook formatting
            restriction = f"[ReceivedTime] >= #{start_date}# AND [ReceivedTime] < #{end_date}#"
            filtered_messages = messages.Restrict(restriction)

            total_messages = len(filtered_messages)
            print(f"Found {total_messages} messages received on {backup_date.strftime('%Y-%m-%d')} in '{mailbox_name}' inbox.")

            saved_emails = 0
            fallback_emails = 0
            processing_errors = 0

            for idx, message in enumerate(filtered_messages):
                try:
                    # Handle cases with no subject
                    subject = sanitize_filename(message.Subject or "No Subject")
                    filename = truncate_or_fallback_filename(save_directory, subject)

                    # Full path for the file
                    full_path = os.path.join(save_directory, filename)

                    # Save the email
                    if save_email(message, full_path):
                        if "Email_Subject_Changed_" in filename:
                            fallback_emails += 1
                        saved_emails += 1
                except Exception as e:
                    print(f"Error processing email '{message.Subject}': {str(e)}")
                    processing_errors += 1
                    continue

            # Log metrics for the date
            log_metrics(backup_date.strftime('%Y-%m-%d'), total_messages, saved_emails, fallback_emails, processing_errors)
            print(f"Backup for {backup_date.strftime('%Y-%m-%d')} completed successfully.\n")

    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    mailbox_name = "Mailbox"
    backup_root_directory = r"C:\EmailBackups"

    print("Enter start date (YYYY-MM-DD): ")
    start_date = input()
    print("Enter end date (YYYY-MM-DD): ")
    end_date = input()

    try:
        start_date_dt = datetime.datetime.strptime(start_date, '%Y-%m-%d')
        end_date_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d')
        if start_date_dt > end_date_dt:
            print("Start date cannot be after end date. Exiting.")
            exit()

        # Generate list of dates in range
        backup_dates = [
            (start_date_dt + datetime.timedelta(days=i)).strftime('%Y-%m-%d')
            for i in range((end_date_dt - start_date_dt).days + 1)
        ]
    except ValueError:
        print("Invalid date format. Please use YYYY-MM-DD.")
        exit()

    backup_shared_mailbox(mailbox_name, backup_root_directory, backup_dates)
