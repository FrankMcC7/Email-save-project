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
    invalid_chars = '<>:"/\\|?*'
    filename = ''.join(c if c not in invalid_chars else '_' for c in filename)
    filename = ''.join(c for c in filename if c.isprintable() and (c.isalnum() or c in ' ._-'))
    filename = filename.strip()
    if len(filename) > max_length:
        filename = filename[:max_length]
    return filename


def truncate_or_fallback_filename(save_directory, subject, max_path_length=255):
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
    try:
        if hasattr(item, "SaveAs"):
            item.SaveAs(save_path, 3)  # 3 refers to the MSG format
            return True
        return False
    except Exception:
        return False


def get_sender_email(message):
    try:
        sender_email = getattr(message, "SenderEmailAddress", None)
        if sender_email:
            return sender_email
        if message.Sender and message.Sender.AddressEntry:
            sender_email = message.Sender.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            if sender_email:
                return sender_email
        PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        sender_email = message.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        return sender_email if sender_email else "Unknown"
    except Exception as e:
        print(f"Failed to retrieve sender email: {str(e)}")
        return "Unknown"


def log_email_details(backup_date, sender_name, sender_email, subject, file_path):
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
        ws.cell(row=new_row, column=1, value=backup_date)
        ws.cell(row=new_row, column=2, value=sender_name)
        ws.cell(row=new_row, column=3, value=sender_email)
        subject_cell = ws.cell(row=new_row, column=4, value=subject)
        subject_cell.hyperlink = file_path
        subject_cell.font = Font(color="0000FF", underline="single")

        wb.save(EMAIL_LOG_FILE)
    except Exception as e:
        print(f"Failed to log email details: {str(e)}")


# The rest of the script remains unchanged (backup_shared_mailbox, log_metrics, etc.)