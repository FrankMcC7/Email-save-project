import os
import sys
import datetime
import pythoncom
import win32com.client as win32
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

EMAIL_LOG_FILE = "backup_email_log.xlsx"  # File to log email details

def sanitize_filename(filename, max_length=100):
    """
    Remove or replace characters that are invalid in Windows filenames.
    """
    invalid_chars = '<>:"/\\|?*'
    filename = ''.join(c if c not in invalid_chars else '_' for c in filename)
    filename = ''.join(c for c in filename if c.isprintable() and (c.isalnum() or c in ' ._-'))
    filename = filename.strip()
    if len(filename) > max_length:
        filename = filename[:max_length]
    return filename

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

def log_email_details(backup_date, sender_email, subject, file_path):
    """
    Log email details into the Excel file for easy search and access.
    Converts the subject to a hyperlink pointing to the saved file.
    """
    try:
        if os.path.exists(EMAIL_LOG_FILE):
            wb = openpyxl.load_workbook(EMAIL_LOG_FILE)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Email Logs"
            ws.append(["Date", "Sender Email", "Subject"])

        # Get or create the worksheet
        if "Email Logs" in wb.sheetnames:
            ws = wb["Email Logs"]
        else:
            ws = wb.create_sheet(title="Email Logs")
            ws.append(["Date", "Sender Email", "Subject"])

        # Append the email details
        ws = wb["Email Logs"]
        new_row = ws.max_row + 1
        ws.cell(row=new_row, column=1, value=backup_date)  # Date
        ws.cell(row=new_row, column=2, value=sender_email)  # Sender Email
        
        # Add subject as a hyperlink
        subject_cell = ws.cell(row=new_row, column=3, value=subject)
        subject_cell.hyperlink = file_path
        subject_cell.font = Font(color="0000FF", underline="single")

        wb.save(EMAIL_LOG_FILE)
    except Exception as e:
        print(f"Failed to log email details: {str(e)}")

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
            messages = inbox.Items.Restrict(restriction)

            # Exclude 'Recall' messages
            messages = messages.Restrict("[MessageClass] <> 'IPM.Outlook.Recall'")

            total_messages = len(messages)
            print(f"Found {total_messages} messages for {backup_date_str} in '{mailbox_name}'.")

            for message in messages:
                try:
                    if not hasattr(message, "SaveAs"):
                        continue

                    subject = sanitize_filename(message.Subject or "No Subject")
                    sender_email = getattr(message, "SenderEmailAddress", "Unknown")
                    filename = truncate_or_fallback_filename(save_directory, subject)
                    full_path = os.path.join(save_directory, filename)

                    if save_email(message, full_path):
                        log_email_details(
                            backup_date.strftime('%Y-%m-%d'),
                            sender_email,
                            subject,
                            full_path
                        )
                except Exception as e:
                    print(f"Failed to process email: {str(e)}")
                    continue

        # Release Outlook COM objects
        inbox = None
        shared_mailbox = None
        outlook = None

    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    mailbox_name = "GMailbox"
    backup_root_directory = r"C:\EmailBackups"

    while True:
        print("Select backup option:")
        print("1. Backup emails for yesterday")
        print("2. Backup emails for a particular date or range of dates")
        choice = input("Enter 1 or 2: ")

        if choice == '1':
            yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
            backup_dates = [yesterday.strftime('%Y-%m-%d')]
            break
        elif choice == '2':
            while True:
                print("Do you want to backup emails for:")
                print("1. A single date")
                print("2. A range of dates")
                date_choice = input("Enter 1 or 2: ")

                if date_choice == '1':
                    date_input = input("Enter the date (YYYY-MM-DD): ")
                    try:
                        backup_date = datetime.datetime.strptime(date_input, '%Y-%m-%d')
                        backup_dates = [backup_date.strftime('%Y-%m-%d')]
                        break
                    except ValueError:
                        print("Invalid date format. Please use YYYY-MM-DD.")
                        continue
                elif date_choice == '2':
                    start_date = input("Enter start date (YYYY-MM-DD): ")
                    end_date = input("Enter end date (YYYY-MM-DD): ")
                    try:
                        start_dt = datetime.datetime.strptime(start_date, '%Y-%m-%d')
                        end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d')
                        if start_dt > end_dt:
                            print("Start date cannot be after end date.")
                            continue
                        backup_dates = [
                            (start_dt + datetime.timedelta(days=i)).strftime('%Y-%m-%d')
                            for i in range((end_dt - start_dt).days + 1)
                        ]
                        break
                    except ValueError:
                        print("Invalid date format. Please use YYYY-MM-DD.")
                        continue
                else:
                    print("Invalid choice. Please enter 1 or 2.")
            break
        else:
            print("Invalid choice. Please enter 1 or 2.")

    backup_shared_mailbox(mailbox_name, backup_root_directory, backup_dates)
