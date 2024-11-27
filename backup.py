import os
import datetime
import pythoncom
import win32com.client as win32
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

EMAIL_LOG_FILE = "backup_email_log.xlsm"  # Macro-enabled file to log email details
METRICS_FILE = "backup_metrics.xlsx"  # File to log backup metrics


def sanitize_filename(filename, max_length=100):
    """
    Remove or replace invalid characters from filenames.
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
    max_filename_length = max_path_length - len(save_directory) - len('.msg') - len(os.sep)
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
            return True
        return False
    except Exception:
        return False


def get_sender_email(message):
    """
    Retrieve the sender's email address using multiple fallback methods.
    """
    try:
        # Attempt to retrieve SMTP address directly from the PropertyAccessor
        PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        sender_email = message.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        if sender_email:
            return sender_email

        # Attempt to retrieve the email from the AddressEntry
        if message.Sender and message.Sender.AddressEntry:
            address_entry = message.Sender.AddressEntry
            if address_entry.Type == "EX":  # If it's an Exchange address
                exchange_user = address_entry.GetExchangeUser()
                if exchange_user:
                    sender_email = exchange_user.PrimarySmtpAddress
                    return sender_email
            else:
                # For non-Exchange addresses, try the Address property
                return address_entry.Address

        # Fallback to "Unknown" if no email is found
        return "Unknown"
    except Exception as e:
        print(f"Failed to retrieve sender email: {str(e)}")
        return "Unknown"


def log_email_details(backup_date, sender_name, sender_email, subject, file_path):
    """
    Log email details into the macro-enabled Excel file.
    """
    try:
        if os.path.exists(EMAIL_LOG_FILE):
            wb = openpyxl.load_workbook(EMAIL_LOG_FILE, keep_vba=True)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Email Logs"
            ws.append(["Date", "Sender Name", "Sender Email", "Subject"])

        if "Email Logs" in wb.sheetnames:
            ws = wb["Email Logs"]
        else:
            ws = wb.create_sheet(title="Email Logs")
            ws.append(["Date", "Sender Name", "Sender Email", "Subject"])

        ws = wb["Email Logs"]
        new_row = ws.max_row + 1
        ws.cell(row=new_row, column=1, value=backup_date)  # Date
        ws.cell(row=new_row, column=2, value=sender_name)  # Sender Name
        ws.cell(row=new_row, column=3, value=sender_email)  # Sender Email
        subject_cell = ws.cell(row=new_row, column=4, value=subject)  # Subject
        subject_cell.hyperlink = file_path
        subject_cell.font = Font(color="0000FF", underline="single")

        wb.save(EMAIL_LOG_FILE)
    except Exception as e:
        print(f"Failed to log email details: {str(e)}")


def log_metrics(backup_date, total_emails, saved_emails, fallback_emails, errors):
    """
    Log backup metrics in a separate Excel file.
    """
    try:
        if os.path.exists(METRICS_FILE):
            wb = openpyxl.load_workbook(METRICS_FILE)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Backup Metrics"
            ws.append(["Date", "Total Emails", "Saved Emails", "Fallback Emails", "Errors"])

        if "Backup Metrics" in wb.sheetnames:
            ws = wb["Backup Metrics"]
        else:
            ws = wb.create_sheet(title="Backup Metrics")
            ws.append(["Date", "Total Emails", "Saved Emails", "Fallback Emails", "Errors"])

        ws = wb["Backup Metrics"]
        ws.append([backup_date, total_emails, saved_emails, fallback_emails, errors])
        wb.save(METRICS_FILE)
    except Exception as e:
        print(f"Failed to log metrics: {str(e)}")


def backup_shared_mailbox(mailbox_name, backup_root_directory, backup_dates):
    """
    Back up all emails from the specified shared mailbox for the given dates.
    """
    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

        shared_mailbox = None
        for folder in outlook.Folders:
            if folder.Name.lower() == mailbox_name.lower():
                shared_mailbox = folder
                break

        if not shared_mailbox:
            print(f"Could not find the shared mailbox: {mailbox_name}")
            return

        inbox = shared_mailbox.Folders["Inbox"]

        for backup_date_str in backup_dates:
            backup_date = datetime.datetime.strptime(backup_date_str, '%Y-%m-%d')
            save_directory = os.path.join(
                backup_root_directory,
                backup_date.strftime('%Y'),
                backup_date.strftime('%m-%B'),
                backup_date.strftime('%d-%m-%Y')
            )

            os.makedirs(save_directory, exist_ok=True)

            start_date = backup_date.strftime('%m/%d/%Y 00:00')
            end_date = (backup_date + datetime.timedelta(days=1)).strftime('%m/%d/%Y 00:00')
            restriction = f"[ReceivedTime] >= '{start_date}' AND [ReceivedTime] < '{end_date}'"
            messages = inbox.Items.Restrict(restriction).Restrict("[MessageClass] <> 'IPM.Outlook.Recall'")

            total_messages = len(messages)
            saved_emails = 0
            fallback_emails = 0
            errors = 0

            for message in messages:
                try:
                    subject = sanitize_filename(message.Subject or "No Subject")
                    sender_name = getattr(message, "SenderName", "Unknown")
                    sender_email = get_sender_email(message)
                    filename = truncate_or_fallback_filename(save_directory, subject)
                    full_path = os.path.join(save_directory, filename)

                    if save_email(message, full_path):
                        if filename != f"{subject}.msg":
                            fallback_emails += 1
                        saved_emails += 1
                        log_email_details(
                            backup_date.strftime('%Y-%m-%d'),
                            sender_name,
                            sender_email,
                            subject,
                            full_path
                        )
                    else:
                        errors += 1
                except Exception as e:
                    print(f"Failed to process email: {str(e)}")
                    errors += 1

            log_metrics(backup_date.strftime('%Y-%m-%d'), total_messages, saved_emails, fallback_emails, errors)

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

    choice = input("Backup for 1) Yesterday or 2) Specific date/range? Enter 1 or 2: ")
    if choice == '1':
        yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
        backup_shared_mailbox(mailbox_name, backup_root_directory, [yesterday.strftime('%Y-%m-%d')])
    elif choice == '2':
        start_date = input("Enter start date (YYYY-MM-DD): ")
        end_date = input("Enter end date (YYYY-MM-DD): ")
        dates = [
            (datetime.datetime.strptime(start_date, '%Y-%m-%