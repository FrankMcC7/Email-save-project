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
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    # Also remove non-printable characters
    filename = ''.join(c for c in filename if c.isprintable())
    return filename.strip()

def truncate_filename(full_path, filename, max_path_length=255):
    """
    Truncate the filename to ensure the total path length is less than 255 characters.
    """
    save_path_length = len(os.path.dirname(full_path)) + 1  # Include the directory path and slash
    max_filename_length = max_path_length - save_path_length

    # Ensure filename doesn't exceed the max length
    if len(filename) > max_filename_length:
        filename = filename[:max_filename_length - 4] + ".msg"  # Reserve space for ".msg"

    return filename

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

def log_metrics(date, total_emails, saved_emails, failed_emails):
    """
    Log the backup metrics in an Excel workbook.
    """
    if not os.path.exists(METRICS_FILE):
        # Create a new workbook if it doesn't exist
        wb = Workbook()
        ws = wb.active
        ws.title = "Backup Metrics"
        ws.append(["Date", "Total Emails", "Saved Emails", "Failed Emails"])
        wb.save(METRICS_FILE)

    # Open the workbook and append the metrics
    wb = openpyxl.load_workbook(METRICS_FILE)
    ws = wb["Backup Metrics"]
    ws.append([date, total_emails, saved_emails, failed_emails])
    wb.save(METRICS_FILE)

def generate_date_range(start_date, end_date):
    """
    Generate a list of dates from start_date to end_date inclusive.
    """
    date_range = []
    current_date = start_date
    while current_date <= end_date:
        date_range.append(current_date.strftime('%Y-%m-%d'))
        current_date += datetime.timedelta(days=1)
    return date_range

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
                os.makedirs(save_directory)

            # Filter emails received on the backup date
            messages = inbox.Items

            # Set the date range for the filter (from start to end of the backup date)
            start_date = backup_date.strftime('%m/%d/%Y 12:00 AM')
            end_date = (backup_date + datetime.timedelta(days=1)).strftime('%m/%d/%Y 12:00 AM')

            restriction = f"[ReceivedTime] >= '{start_date}' AND [ReceivedTime] < '{end_date}'"

            filtered_messages = messages.Restrict(restriction)

            total_messages = len(filtered_messages)
            print(f"Found {total_messages} messages received on {backup_date.strftime('%Y-%m-%d')} in '{mailbox_name}' inbox.")

            saved_emails = 0
            failed_emails = 0

            for idx, message in enumerate(filtered_messages):
                try:
                    # Handle cases with no sender or subject
                    subject = sanitize_filename(message.Subject or "No Subject")
                    sender = sanitize_filename(message.SenderName or "Unknown Sender")
                    filename = f"{subject}.msg"

                    # Truncate the filename if necessary
                    full_path = os.path.join(save_directory, filename)
                    filename = truncate_filename(full_path, filename)
                    full_path = os.path.join(save_directory, filename)

                    # Ensure uniqueness by appending a counter if the file already exists
                    counter = 1
                    while os.path.exists(full_path):
                        filename = f"{subject}_{counter}.msg"
                        filename = truncate_filename(os.path.join(save_directory, filename), filename)
                        full_path = os.path.join(save_directory, filename)
                        counter += 1

                    # Save the email
                    if save_email(message, full_path):
                        saved_emails += 1
                    else:
                        failed_emails += 1

                except Exception as e:
                    print(f"Error processing email '{message.Subject}': {str(e)}")
                    failed_emails += 1
                    continue

            # Log metrics for the date
            log_metrics(backup_date.strftime('%Y-%m-%d'), total_messages, saved_emails, failed_emails)
            print(f"Backup for {backup_date.strftime('%Y-%m-%d')} completed successfully.\n")

    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    # Hardcoded shared mailbox name
    mailbox_name = "Mailbox"  # Replace with the exact name of your shared mailbox

    # Hardcoded backup root directory
    backup_root_directory = r"C:\EmailBackups"  # Replace with your desired backup root directory

    # Prompt the user for dates
    print("Would you like to provide:")
    print("1. Specific dates (comma-separated, e.g., 2024-11-19,2024-11-20)")
    print("2. A date range (e.g., 2024-11-19 to 2024-11-25)")
    choice = input("Enter 1 or 2: ").strip()

    if choice == "1":
        dates_input = input("Enter date(s) (YYYY-MM-DD, comma-separated): ")
        backup_dates = [date.strip() for date in dates_input.split(",")]
    elif choice == "2":
        start_date_input = input("Enter start date (YYYY-MM-DD): ").strip()
        end_date_input = input("Enter end date (YYYY-MM-DD): ").strip()
        try:
            start_date = datetime.datetime.strptime(start_date_input, '%Y-%m-%d')
            end_date = datetime.datetime.strptime(end_date_input, '%Y-%m-%d')
            if start_date > end_date:
                print("Start date cannot be after end date. Exiting.")
                exit()
            backup_dates = generate_date_range(start_date, end_date)
        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD.")
            exit()
    else:
        print("Invalid choice. Exiting.")
        exit()

    if not backup_dates or backup_dates == ['']:
        print("No dates entered. Exiting the script.")
    else:
        backup_shared_mailbox(mailbox_name, backup_root_directory, backup_dates)
