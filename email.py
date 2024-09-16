import os
import datetime
import pythoncom
import win32com.client
import pandas as pd
import json
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_socketio import SocketIO
import unicodedata

app = Flask(__name__)
app.secret_key = 'supersecretkey'
socketio = SocketIO(app)

# Load configurations from a JSON file
try:
    with open('config.json', 'r') as f:
        config = json.load(f)
except FileNotFoundError:
    config = {}
    print("Configuration file 'config.json' not found. Using default settings.")

LOG_FILE_PATH = config.get('LOG_FILE_PATH', 'logs.txt')
NEW_SENDERS_CSV = config.get('NEW_SENDERS_CSV', 'new_senders.csv')

def sanitize_filename(filename):
    # Remove or replace problematic characters
    sanitized = unicodedata.normalize('NFKD', filename).encode('ASCII', 'ignore').decode('ASCII')
    sanitized = ''.join(c if c.isalnum() or c in (' ', '.', '_') else '_' for c in sanitized)
    sanitized = sanitized.strip()
    sanitized = sanitized[:255]
    return sanitized

def extract_emails_for_date(email_address, specific_date_str, sender_path_table):
    logs = []
    new_senders = []

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
    items = items.Restrict(f"[ReceivedTime] >= '{specific_date.strftime('%m/%d/%Y')} 12:00 AM' AND [ReceivedTime] <= '{specific_date.strftime('%m/%d/%Y')} 11:59 PM'")

    # Standardize column names to lowercase
    sender_path_table.columns = sender_path_table.columns.str.lower()

    extracted_emails = []

    for item in items:
        try:
            sender_email = item.SenderEmailAddress.lower() if hasattr(item, 'SenderEmailAddress') else item.Sender.Address.lower()
            subject = item.Subject
            extracted_emails.append({'sender_email': sender_email, 'subject': subject})
        except Exception as e:
            logs.append(f"Error processing email: {str(e)}")
            continue

    # Create a DataFrame from extracted emails
    extracted_df = pd.DataFrame(extracted_emails)

    # Remove duplicates
    extracted_df.drop_duplicates(subset=['sender_email', 'subject'], inplace=True)

    # Iterate over extracted emails
    for idx, row in extracted_df.iterrows():
        sender_email = row['sender_email']
        subject = row['subject']
        subject_lower = subject.lower() if subject else ''

        # Check if the sender exists in the database CSV
        matching_rows = sender_path_table[sender_path_table['sender'].str.lower() == sender_email]

        if matching_rows.empty:
            # Sender email not found in database; new sender
            new_senders.append({'sender': sender_email, 'subject': subject})
            continue  # Proceed to next email

        else:
            # Sender email exists in database
            is_known_sender = False

            # Iterate over all matching rows for the sender email
            for _, db_row in matching_rows.iterrows():
                coper_name = str(db_row.get('coper_name', '')).strip().lower()

                if not coper_name or pd.isna(coper_name):
                    # 'coper_name' is empty or NaN; known sender
                    is_known_sender = True
                    break  # No need to check further

                else:
                    # 'coper_name' is populated; check if it's in the subject
                    if coper_name in subject_lower:
                        # 'coper_name' is in the subject; known sender
                        is_known_sender = True
                        break  # No need to check further
                    else:
                        # 'coper_name' not in subject; continue checking other rows
                        continue

            if not is_known_sender:
                # Sender email exists but 'coper_name' not matched; new sender context
                new_senders.append({'sender': sender_email, 'subject': subject})

    # Save new senders to CSV
    if new_senders:
        new_senders_df = pd.DataFrame(new_senders)
        if os.path.exists(NEW_SENDERS_CSV):
            # Append to existing CSV
            new_senders_df.to_csv(NEW_SENDERS_CSV, mode='a', index=False, header=False)
        else:
            # Create new CSV
            new_senders_df.to_csv(NEW_SENDERS_CSV, index=False)

    # Save logs
    with open(LOG_FILE_PATH, 'w', encoding='utf-8') as f:
        f.writelines("\n".join(logs))

    pythoncom.CoUninitialize()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        date_str = request.form['date']
        file = request.files['file']
        if file and date_str:
            try:
                datetime.datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                flash("Invalid date format. Please enter the date in YYYY-MM-DD format.", 'error')
                return redirect(url_for('index'))

            filepath = os.path.join('uploads', sanitize_filename(file.filename))
            file.save(filepath)

            try:
                sender_path_table = pd.read_csv(filepath, encoding='utf-8')
            except UnicodeDecodeError:
                try:
                    sender_path_table = pd.read_csv(filepath, encoding='latin1')
                except Exception as e:
                    flash("Error reading the CSV file. Please ensure it's properly formatted.", 'error')
                    return redirect(url_for('index'))

            # Standardize column names to lowercase
            sender_path_table.columns = sender_path_table.columns.str.lower()

            account_email_address = "hf_data@bofa.com"
            socketio.start_background_task(extract_emails_for_date, account_email_address, date_str, sender_path_table)
            return redirect(url_for('results'))

    return render_template('index.html')

@app.route('/results')
def results():
    logs = []
    new_senders = []

    if os.path.exists(LOG_FILE_PATH):
        with open(LOG_FILE_PATH, 'r', encoding='utf-8') as f:
            logs = f.readlines()

    if os.path.exists(NEW_SENDERS_CSV):
        new_senders_df = pd.read_csv(NEW_SENDERS_CSV)
        new_senders = new_senders_df.to_dict('records')

    return render_template('results.html', logs=logs, new_senders=new_senders)

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    socketio.run(app, debug=True, use_reloader=False)
