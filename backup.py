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
    filename = ''.join(c for c in filename if c.isprintable() and (c.isalnum() or c in ' ._-'))
    return filename.strip()


def truncate_or_fallback_filename(save_directory, subject, max_path_length=255):
    """
    Generate a unique filename by appending a counter if necessary.
    """
    max_filename_length = max_path_length - len(save_directory) - len(os.sep) - len('.msg')
    if max_filename_length <= 0:
        raise Exception("Save directory path is too long.")

    base_subject = subject
    counter = 0
    while True:
        if counter == 0:
            filename = f"{subject}.msg"
        else:
            filename = f"{base_subject}_{counter}.msg"

        if len(filename) > max_filename_length:
            excess_length = len(filename) - max_filename_length
            base_subject = base_subject[:-excess_length]
            if not base_subject:
                base_subject = 'Email_Subject_Changed'
            if counter == 0:
                filename = f"{base_subject}.msg"
            else:
                filename = f"{base_subject}_{counter}.msg"

        full_path = os.path.join(save_directory, filename)
        if not os.path.exists(full_path):
            return filename
        counter += 1


def save_email(item, save_path):
    """
    Save an email as a .msg file.
    """
    try:
        if hasattr(item, "SaveAs"):
            item.SaveAs(save_path, 3)  # 3 refers to the MSG format
            print(f"Saved email: {save_path}")
            return True
        else:
            print(f"Item does not support SaveAs: {getattr(item, 'Subject', 'No Subject')}")
            return False
    except Exception as e:
        print(f"Failed to save email: {getattr(item, 'Subject', 'No Subject')} - Error: {str(e)}")
        return False


def log_metrics(date, total_emails, saved_emails, fallback_emails, processing_errors, failed_emails):
    """
    Log the backup metrics in an Excel workbook.
    Also, log the failed emails in a separate sheet.
    """
    if not os.path.exists(METRICS_FILE):
        # Create a new workbook if it doesn't exist
        wb = Workbook()
        ws = wb.active
        ws.title = "Backup Metrics"
        ws.append(["Date", "Total Emails", "Saved Emails", "Fallback Emails", "Processing Errors"])
        ws_failed = wb.create_sheet(title="Failed Emails")
        ws_failed.append(["Date", "Sender Address", "Subject", "Error Message"])
        wb.save(METRICS_FILE)

    # Open the workbook and append the metrics
    wb = openpyxl.load_workbook(METRICS_FILE)
    ws = wb["Backup Metrics"]
    ws.append([date, total_emails, saved_emails, fallback_emails, processing_errors])

    if failed_emails:
        ws_failed = wb["Failed Emails"]
        for failed_email in failed_emails:
            ws_failed.append(failed_email)

    wb.save(METRICS_FILE)


def backup_shared_mailbox(mailbox_name, backup_root_directory, backup_dates):
    """
    Back up all emails from the specified shared mailbox for the given dates.
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

        inbox = shared_mailbox.Folders["Inbox"]

        for backup_date_str in backup_dates:
            try:
                backup_date = datetime.datetime.strptime(backup_date_str, '%Y-%m-%d')
            except ValueError:
                print(f"Invalid date format: {backup_date_str}. Please use YYYY-MM-DD.")
                continue

            year_str = backup_date.strftime('%Y')
            month_str = backup_date.strftime('%m-%B')
            date_str = backup_date.strftime('%d-%m-%Y')
            save_directory = os.path.join(backup_root_directory, year_str, month_str, date_str)

            if not os.path.exists(save_directory):
                try:
                    os.makedirs(save_directory)
                except Exception as e:
                    print(f"Failed to create directory '{save_directory}': {str(e)}")
                    continue

            start_date = backup_date.strftime('%m/%d/%Y 00:00')
            end_date = (backup_date + datetime.timedelta(days=1)).strftime('%m/%d/%Y 00:00')
            restriction = f"[ReceivedTime] >= '{start_date}' AND [ReceivedTime] < '{end_date}'"
            filtered_messages = inbox.Items.Restrict(restriction)

            total_messages = len(filtered_messages)
            print(f"Found {total_messages} messages for {backup_date_str} in '{mailbox_name}'.")

            saved_emails = 0
            fallback_emails = 0
            processing_errors = 0
            failed_emails = []

            for message in filtered_messages:
                try:
                    if not hasattr(message, "SaveAs"):
                        failed_emails.append((
                            backup_date.strftime('%Y-%m-%d'),
                            getattr(message, 'SenderEmailAddress', 'Unknown'),
                            getattr(message, 'Subject', 'No Subject'),
                            "Unsupported item type"
                        ))
                        processing_errors += 1
                        continue

                    subject = sanitize_filename(message.Subject or "No Subject")
                    filename = truncate_or_fallback_filename(save_directory, subject)
                    full_path = os.path.join(save_directory, filename)

                    if save_email(message, full_path):
                        if filename != f"{subject}.msg":
                            fallback_emails += 1
                        saved_emails += 1
                    else:
                        processing_errors += 1
                        failed_emails.append((
                            backup_date.strftime('%Y-%m-%d'),
                            getattr(message, 'SenderEmailAddress', 'Unknown'),
                            subject,
                            "Save failed"
                        ))
                except Exception as e:
                    failed_emails.append((
                        backup_date.strftime('%Y-%m-%d'),
                        getattr(message, 'SenderEmailAddress', 'Unknown'),
                        getattr(message, 'Subject', 'No Subject'),
                        str(e)
                    ))
                    processing_errors += 1
                    continue

            log_metrics(
                backup_date.strftime('%Y-%m-%d'),
                total_messages,
                saved_emails,
                fallback_emails,
                processing_errors,
                failed_emails
            )
            print(f"Backup for {backup_date_str} completed.\n")

    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    mailbox_name = "GMailbox"
    backup_root_directory = r"C:\EmailBackups"

    print("Select backup option:")
    print("1. Backup emails for yesterday")
    print("2. Backup emails for a particular date or range of dates")
    choice = input("Enter 1 or 2: ")

    if choice == '1':
        yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
        backup_dates = [yesterday.strftime('%Y-%m-%d')]
    elif choice == '2':
        print("Do you want to backup emails for:")
        print("1. A single date")
        print("2. A range of dates")
        date_choice = input("Enter 1 or 2: ")

        if date_choice == '1':
            date_input = input("Enter the date (YYYY-MM-DD): ")
            try:
                backup_dates = [datetime.datetime.strptime(date_input, '%Y-%m-%d').strftime('%Y-%m-%d')]
            except ValueError:
                print("Invalid date format. Please use YYYY-MM-DD.")
                exit()
        elif date_choice == '2':
            start_date = input("Enter start date (YYYY-MM-DD): ")
            end_date = input("Enter end date (YYYY-MM-DD): ")
            try:
                start_dt = datetime.datetime.strptime(start_date, '%Y-%m-%d')
                end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d')
                if start_dt > end_dt:
                    print("Start date cannot be after end date.")
                    exit()
                backup_dates = [
                    (start_dt + datetime.timedelta(days=i)).strftime('%Y-%m-%d')
                    for i in range((end_dt - start_dt).days + 1)
                ]
            except ValueError:
                print("Invalid date format. Please use YYYY-MM-DD.")
                exit()
        else:
            print("Invalid choice.")
            exit()
    else:
        print("Invalid choice.")
        exit()

    backup_shared_mailbox(mailbox_name, backup_root_directory, backup_dates)
