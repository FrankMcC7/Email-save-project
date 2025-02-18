import os
import datetime
import pythoncom
import win32com.client
import pandas as pd
import unicodedata

# Hardcoded configuration for file paths and Outlook email account
CSV_FILE_PATH = "C:/path/to/your/sender_info.csv"    # <-- Update with your CSV file path
OUTLOOK_EMAIL = "your_email_address@example.com"     # <-- Update with your Outlook email address
NEW_SENDERS_CSV = "C:/path/to/output/new_senders.csv"  # <-- Update with your desired output path

LOG_FILE_PATH = "logs.txt"  # You can also hardcode the logs file path if desired

def sanitize_filename(filename):
    """Sanitize a filename by removing or replacing problematic characters."""
    sanitized = unicodedata.normalize('NFKD', filename).encode('ASCII', 'ignore').decode('ASCII')
    sanitized = ''.join(c if c.isalnum() or c in (' ', '.', '_') else '_' for c in sanitized)
    return sanitized.strip()[:255]

def extract_emails_for_range(email_address, start_date_str, end_date_str, sender_path_table):
    logs = []
    new_senders = []
    new_sender_emails = set()

    pythoncom.CoInitialize()
    try:
        start_date = datetime.datetime.strptime(start_date_str, '%Y-%m-%d').date()
        end_date = datetime.datetime.strptime(end_date_str, '%Y-%m-%d').date()
    except ValueError as ve:
        logs.append(f"Date conversion error: {ve}")
        with open(LOG_FILE_PATH, 'w', encoding='utf-8') as f:
            f.write("\n".join(logs))
        pythoncom.CoUninitialize()
        return

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Locate the Inbox folder for the given account
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
        with open(LOG_FILE_PATH, 'w', encoding='utf-8') as f:
            f.write("\n".join(logs))
        pythoncom.CoUninitialize()
        return

    # Locate the "NAV and Performance" folder within Inbox (if it exists)
    nav_perf_folder = None
    try:
        nav_perf_folder = next((folder for folder in inbox.Folders if folder.Name.lower() == "nav and performance"), None)
    except Exception as e:
        logs.append(f"Error accessing 'NAV and Performance' folder: {str(e)}")

    # Process emails from both the Inbox and the NAV and Performance folder (if available)
    folders_to_process = [inbox]
    if nav_perf_folder:
        folders_to_process.append(nav_perf_folder)

    # Build a date filter for the specified date range
    date_filter = (
        f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y')} 12:00 AM' AND "
        f"[ReceivedTime] <= '{end_date.strftime('%m/%d/%Y')} 11:59 PM'"
    )

    all_items = []
    for folder in folders_to_process:
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            restricted_items = items.Restrict(date_filter)
            for item in restricted_items:
                all_items.append(item)
        except Exception as e:
            logs.append(f"Error processing folder {folder.Name}: {str(e)}")

    # Standardize the CSV's column names to lowercase
    sender_path_table.columns = sender_path_table.columns.str.lower()

    # Normalize the 'sender' column in the CSV
    sender_path_table['sender'] = sender_path_table['sender'].astype(str).str.strip()
    sender_path_table['sender'] = sender_path_table['sender'].apply(
        lambda x: unicodedata.normalize('NFKD', x).encode('ASCII', 'ignore').decode('ASCII').lower()
    )

    extracted_emails = []
    for item in all_items:
        try:
            sender_email = (item.SenderEmailAddress if hasattr(item, 'SenderEmailAddress')
                            else item.Sender.Address)
            sender_email = sender_email.strip().lower()
            sender_email = unicodedata.normalize('NFKD', sender_email).encode('ASCII', 'ignore').decode('ASCII')
            subject = item.Subject
            extracted_emails.append({'sender_email': sender_email, 'subject': subject})
        except Exception as e:
            logs.append(f"Error processing email: {str(e)}")
            continue

    # Remove duplicate email records
    extracted_df = pd.DataFrame(extracted_emails)
    extracted_df.drop_duplicates(subset=['sender_email', 'subject'], inplace=True)

    # Compare extracted emails against the CSV data
    for idx, row in extracted_df.iterrows():
        sender_email = row['sender_email']
        subject = row['subject']
        subject_lower = subject.lower() if subject else ''

        matching_rows = sender_path_table[sender_path_table['sender'] == sender_email]
        if matching_rows.empty:
            if sender_email not in new_sender_emails:
                new_sender_emails.add(sender_email)
                new_senders.append({'sender': sender_email, 'subject': subject})
            continue
        else:
            is_known_sender = False
            for _, db_row in matching_rows.iterrows():
                coper_name = str(db_row.get('coper_name', '')).strip().lower()
                if not coper_name or pd.isna(coper_name) or coper_name == 'nan':
                    is_known_sender = True
                    break
                elif coper_name in subject_lower:
                    is_known_sender = True
                    break
            if not is_known_sender:
                if sender_email not in new_sender_emails:
                    new_sender_emails.add(sender_email)
                    new_senders.append({'sender': sender_email, 'subject': subject})

    # Save new senders to CSV at the hardcoded path
    if new_senders:
        new_senders_df = pd.DataFrame(new_senders)
        new_senders_df.drop_duplicates(subset=['sender'], inplace=True)
        if os.path.exists(NEW_SENDERS_CSV):
            existing_senders_df = pd.read_csv(NEW_SENDERS_CSV)
            combined_df = pd.concat([existing_senders_df, new_senders_df], ignore_index=True)
            combined_df.drop_duplicates(subset=['sender'], inplace=True)
            combined_df.to_csv(NEW_SENDERS_CSV, index=False)
        else:
            new_senders_df.to_csv(NEW_SENDERS_CSV, index=False)

    # Save logs to file
    with open(LOG_FILE_PATH, 'w', encoding='utf-8') as f:
        f.write("\n".join(logs))

    pythoncom.CoUninitialize()

def main():
    # Use the hardcoded CSV file path and Outlook email address
    csv_path = CSV_FILE_PATH
    email_address = OUTLOOK_EMAIL

    if not os.path.exists(csv_path):
        print(f"CSV file {csv_path} does not exist. Exiting.")
        return

    # Ask the user if they want to process a single date or a range of dates
    date_type = input("Do you want to run for a single date or a date range? (Enter 'S' for single, 'R' for range): ").strip().upper()
    if date_type == "S":
        date_str = input("Enter the date (YYYY-MM-DD): ").strip()
        start_date_str = date_str
        end_date_str = date_str
    elif date_type == "R":
        start_date_str = input("Enter the start date (YYYY-MM-DD): ").strip()
        end_date_str = input("Enter the end date (YYYY-MM-DD): ").strip()
    else:
        print("Invalid selection. Exiting.")
        return

    # Read the CSV file; try utf-8 first then latin1 if needed
    try:
        sender_path_table = pd.read_csv(csv_path, encoding='utf-8')
    except UnicodeDecodeError:
        try:
            sender_path_table = pd.read_csv(csv_path, encoding='latin1')
        except Exception as e:
            print("Error reading the CSV file. Please ensure it's properly formatted.")
            return

    # Execute the email extraction process
    extract_emails_for_range(email_address, start_date_str, end_date_str, sender_path_table)

    # Display logs
    if os.path.exists(LOG_FILE_PATH):
        print("\nLogs:")
        with open(LOG_FILE_PATH, 'r', encoding='utf-8') as f:
            print(f.read())
    else:
        print("No logs available.")

    # Display new senders
    if os.path.exists(NEW_SENDERS_CSV):
        print("\nNew Senders:")
        new_senders_df = pd.read_csv(NEW_SENDERS_CSV)
        if not new_senders_df.empty:
            print(new_senders_df)
        else:
            print("No new senders found.")
    else:
        print("No new senders file found.")

if __name__ == '__main__':
    main()
