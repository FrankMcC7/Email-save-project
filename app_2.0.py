import os
import datetime
import re
import pythoncom
import win32com.client
import pandas as pd
import nest_asyncio
import json
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_socketio import SocketIO
from threading import Thread
import openpyxl
from concurrent.futures import ThreadPoolExecutor

# Apply nest_asyncio to allow nested event loops
nest_asyncio.apply()

app = Flask(__name__)
app.secret_key = 'supersecretkey'
socketio = SocketIO(app)

# Load configurations from a JSON file
with open('config.json', 'r') as f:
    config = json.load(f)

DEFAULT_SAVE_PATH = config.get('DEFAULT_SAVE_PATH', 'path_to_default_folder')
LOG_FILE_PATH = config.get('LOG_FILE_PATH', 'logs.txt')
EXCEL_FILE_PATH = config.get('EXCEL_FILE_PATH', 'email_summary.xlsx')

# Compile regex pattern once to improve performance
date_pattern = re.compile(r"""
    (?i)
    \b(January|February|March|April|May|June|July|August|September|October|November|December)\b|
    \b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b|
    (\d{2})[./-](\d{4})|                       # Formats like 08.2024, 08-2024
    (\d{8})|                                  # Formats like 08312024 (MMDDYYYY)
    (\d{6})|                                  # Formats like 083124 (MMDDYY)
    (\d{2})[./-](\d{2})[./-](\d{4})|          # Formats like 08-31-2024, 08.31.2024
    (\d{4})-(\d{2})-(\d{2})|                  # Formats like 2024-08-31
    (Q[1-4])['\s]?(\d{2,4})|                  # Quarter formats like Q1 2024
    ([1-4]Q)['\s]?(\d{2,4})|                  # Quarter formats like 1Q 2024
    (\d{2})(\d{4})|                           # Formats like 082024 (MMYYYY)
    (\d{4})(\d{2})                            # Formats like 202408 (YYYYMM)
""", re.IGNORECASE | re.VERBOSE)

quarter_mappings = {'Q1': '03-March', 'Q2': '06-June', 'Q3': '09-September', 'Q4': '12-December'}

def sanitize_filename(filename):
    allowable_chars = re.compile(r'[^a-zA-Z0-9\s\-\_\.\+\%\(\)]')
    sanitized = allowable_chars.sub('_', filename).replace(' ', '_')
    return sanitized[:255]

def extract_date_from_text(text, default_year=None):
    match = date_pattern.search(text)
    if match:
        full_month = match.group(1)
        abbr_month = match.group(2)
        month_year_format = match.group(3)
        mmddyyyy_format = match.group(4)
        mmddyy_format = match.group(5)
        date_format = match.group(6) or match.group(7)
        quarter_1 = match.group(8)
        quarter_year_1 = match.group(9)
        quarter_2 = match.group(10)
        quarter_year_2 = match.group(11)
        month_year_2 = match.group(12)
        year_month_2 = match.group(13)
        year = default_year

        if full_month or abbr_month:
            month_num = datetime.datetime.strptime(full_month or abbr_month, "%B" if full_month else "%b").strftime("%m")
            month_name = datetime.datetime.strptime(full_month or abbr_month, "%B" if full_month else "%b").strftime("%B")
            return year, f"{month_num}-{month_name}"

        if month_year_format:  # Handles 08.2024 or 08-2024
            try:
                month_num, year = month_year_format.split(".") if "." in month_year_format else month_year_format.split("-")
                month_name = datetime.datetime.strptime(month_num, "%m").strftime("%B")
                return year, f"{month_num}-{month_name}"
            except ValueError:
                pass

        if mmddyyyy_format:  # Handles 08312024 (MMDDYYYY)
            try:
                parsed_date = datetime.datetime.strptime(mmddyyyy_format, "%m%d%Y")
                month_num = parsed_date.strftime('%m')
                month_name = parsed_date.strftime('%B')
                year = parsed_date.strftime('%Y')
                return year, f"{month_num}-{month_name}"
            except ValueError:
                pass

        if mmddyy_format:  # Handles 083124 (MMDDYY)
            try:
                parsed_date = datetime.datetime.strptime(mmddyy_format, "%m%d%y")
                month_num = parsed_date.strftime('%m')
                month_name = parsed_date.strftime('%B')
                year = parsed_date.strftime('%Y')
                return year, f"{month_num}-{month_name}"
            except ValueError:
                pass

        if date_format:  # Handles 08-31-2024, 08.31.2024, 2024-08-31
            try:
                parsed_date = datetime.datetime.strptime(date_format, "%m-%d-%Y" if "-" in date_format else "%m.%d.%Y")
                month_num = parsed_date.strftime('%m')
                month_name = parsed_date.strftime('%B')
                year = parsed_date.strftime('%Y')
                return year, f"{month_num}-{month_name}"
            except ValueError:
                pass

        if quarter_1 and quarter_year_1:
            quarter = quarter_mappings[quarter_1.upper()]
            year = f"20{quarter_year_1}" if len(quarter_year_1) == 2 else quarter_year_1
            return year, quarter

        if quarter_2 and quarter_year_2:
            quarter = quarter_mappings[f'Q{quarter_2[0]}']
            year = f"20{quarter_year_2}" if len(quarter_year_2) == 2 else quarter_year_2
            return year, quarter

        if month_year_2:  # Handles 082024 (MMYYYY)
            try:
                month_num = month_year_2[:2]
                year = month_year_2[2:]
                month_name = datetime.datetime.strptime(month_num, "%m").strftime("%B")
                return year, f"{month_num}-{month_name}"
            except ValueError:
                pass

        if year_month_2:  # Handles 202408 (YYYYMM)
            try:
                year = year_month_2[:4]
                month_num = year_month_2[4:]
                month_name = datetime.datetime.strptime(month_num, "%m").strftime("%B")
                return year, f"{month_num}-{month_name}"
            except ValueError:
                pass

    return default_year, None

def find_save_path(sender, subject, sender_path_table):
    sender_lower = sender.lower()
    subject_lower = subject.lower()

    rows = sender_path_table[sender_path_table['sender'].str.lower() == sender_lower]

    for _, row in rows.iterrows():
        keywords = set(str(row.get('keywords', '')).split(';'))
        if any(keyword.lower() in subject_lower for keyword in keywords):
            return row['keyword_path'], row['special_case'], True

    for _, row in rows.iterrows():
        if pd.notna(row['coper_name']) and row['coper_name'].lower() in subject_lower:
            return row['save_path'], row['special_case'], False

    if not rows.empty:
        return rows.iloc[0]['save_path'], rows.iloc[0]['special_case'], False

    return None, None, False

def update_excel_summary(date_str, total_emails, saved_default, saved_actual, not_saved, failed_emails):
    import openpyxl  # Lazy load to improve initial load time

    if os.path.exists(EXCEL_FILE_PATH):
        workbook = openpyxl.load_workbook(EXCEL_FILE_PATH)
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = 'Summary'
        sheet.append(['Date', 'Total Emails', 'Saved in Default', 'Saved in Actual Paths', 'Not Saved'])

    sheet = workbook.active
    sheet.append([date_str, total_emails, saved_default, saved_actual, not_saved])

    if 'Failed Emails' not in workbook.sheetnames:
        failed_sheet = workbook.create_sheet('Failed Emails')
        failed_sheet.append(['Date', 'Email Address', 'Subject'])
    else:
        failed_sheet = workbook['Failed Emails']

    for email in failed_emails:
        failed_sheet.append([date_str, email['email_address'], email['subject']])

    workbook.save(EXCEL_FILE_PATH)

def save_email(item, save_path, special_case):
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    
    valid_extensions = ('.xlsx', '.xls', '.csv', '.pdf', '.doc', '.docx')
    if special_case and item.Attachments.Count > 0:
        for attachment in item.Attachments:
            if attachment.FileName.lower().endswith(valid_extensions):
                filename_base = sanitize_filename(os.path.splitext(attachment.FileName)[0])
                break
        else:
            filename_base = sanitize_filename(item.Subject)
    else:
        filename_base = sanitize_filename(item.Subject)

    extension = ".msg"
    filename = f"{filename_base}{extension}"
    full_path = os.path.join(save_path, filename)

    # Check if file already exists, if yes, add a suffix to avoid overwriting
    counter = 1
    while os.path.exists(full_path):
        filename = f"{filename_base}_{counter}{extension}"
        full_path = os.path.join(save_path, filename)
        counter += 1

    item.SaveAs(full_path, 3)
    return filename

def process_email(item, sender_path_table, default_year, specific_date_str):
    logs = []
    failed_emails = []
    retries = 3
    processed = False

    sender_email = item.SenderEmailAddress.lower() if hasattr(item, 'SenderEmailAddress') else item.Sender.Address.lower()
    subject = item.Subject

    while retries > 0 and not processed:
        try:
            year, month = extract_date_from_text(subject, default_year)
            if not year or not month:
                for attachment in item.Attachments:
                    year, month = extract_date_from_text(attachment.FileName, default_year)
                    if year and month:
                        break
            year = year or default_year
            base_path, special_case, _ = find_save_path(sender_email, subject, sender_path_table)
            save_path = os.path.join(base_path or DEFAULT_SAVE_PATH, str(year), month or specific_date_str)
            filename = save_email(item, save_path, special_case)
            logs.append(f"Saved: {filename} to {save_path}")
            processed = True
        except pythoncom.com_error as com_err:
            retries -= 1
            logs.append(f"COM Error handling email '{subject}' (Code: {com_err.args})")
            if retries == 0:
                logs.append(f"Failed to save the email '{subject}' after 3 retries")
                failed_emails.append({'email_address': sender_email, 'subject': subject})
        except Exception as e:
            retries = 0
            logs.append(f"Error handling email '{subject}': {str(e)}")
            failed_emails.append({'email_address': sender_email, 'subject': subject})

    return logs, failed_emails

def save_emails_from_senders_on_date(email_address, specific_date_str, sender_path_table, default_year):
    logs = []
    pythoncom.CoInitialize()
    specific_date = datetime.datetime.strptime(specific_date_str, '%Y-%m-%d').date()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    inbox = None
    for store in outlook.Stores:
        if store.DisplayName.lower() == email_address.lower() or store.ExchangeStoreType == 3:
            try:
                root_folder = store.GetRootFolder()
                inbox = next((folder for folder in root_folder.Folders if folder.Name.lower() == "inbox"), None)
                if inbox:
                    break
            except AttributeError as e:
                logs.append(f"Error accessing inbox: {str(e)}")
                continue

    if not inbox:
        logs.append(f"No Inbox found for the account with email address: {email_address}")
        pythoncom.CoUninitialize()
        with open(LOG_FILE_PATH, 'w', encoding='utf-8') as f:
            f.writelines("\n".join(logs))
        return

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    items = items.Restrict(f"[ReceivedTime] >= '{specific_date.strftime('%m/%d/%Y')} 00:00 AM' AND [ReceivedTime] <= '{specific_date.strftime('%m/%d/%Y')} 11:59 PM'")

    def process_item(item):
        return process_email(item, sender_path_table, default_year, specific_date_str)

    total_emails, saved_default, saved_actual, not_saved = 0, 0, 0, 0
    failed_emails = []

    with ThreadPoolExecutor(max_workers=4) as executor:
        results = list(executor.map(process_item, items))

    for email_logs, email_failed_emails in results:
        total_emails += 1
        logs.extend(email_logs)
        failed_emails.extend(email_failed_emails)
        if any(DEFAULT_SAVE_PATH in log for log in email_logs):
            saved_default += 1
        else:
            saved_actual += 1

    pythoncom.CoUninitialize()
    with open(LOG_FILE_PATH, 'w', encoding='utf-8') as f:
        f.writelines("\n".join(logs))

    update_excel_summary(specific_date_str, total_emails, saved_default, saved_actual, not_saved, failed_emails)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        date_str = request.form['date']
        default_year = request.form['default_year']
        file = request.files['file']
        if file and date_str and default_year:
            try:
                datetime.datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                flash("Invalid date format. Please enter the date in YYYY-MM-DD format.", 'error')
                return redirect(url_for('index'))

            if not (default_year.isdigit() and len(default_year) == 4):
                flash("Invalid year format. Please enter the year in YYYY format.", 'error')
                return redirect(url_for('index'))

            filepath = os.path.join('uploads', sanitize_filename(file.filename))
            file.save(filepath)

            try:
                sender_path_table = pd.read_csv(filepath, encoding='utf-8')
            except UnicodeDecodeError:
                sender_path_table = pd.read_csv(filepath, encoding='latin1')

            account_email_address = "hf_data@bofa.com"
            socketio.start_background_task(save_emails_from_senders_on_date, account_email_address, date_str, sender_path_table, default_year)
            return redirect(url_for('results'))

    return render_template('index.html')

@app.route('/results')
def results():
    logs = []
    if os.path.exists(LOG_FILE_PATH):
        with open(LOG_FILE_PATH, 'r') as f:
            logs = f.readlines()

    return render_template('results.html', logs=logs)

def run_app():
    socketio.run(app, debug=True, use_reloader=False, allow_unsafe_werkzeug=True)

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    os.makedirs(DEFAULT_SAVE_PATH, exist_ok=True)

    thread = Thread(target=run_app)
    thread.start()
