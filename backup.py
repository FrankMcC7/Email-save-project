import os
import sys
import datetime
import pythoncom
import win32com.client as win32

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

def backup_inbox(email_address, backup_root_directory, backup_dates):
    """
    Backs up all emails from the specified Outlook inbox received on the backup_dates to the specified directory.
    """
    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

        inbox = None
        # Iterate through all accounts to find the matching email address
        for account in outlook.Folders:
            if account.Name.lower() == email_address.lower():
                inbox = account.Folders["Inbox"]
                break

        if inbox is None:
            print(f"Could not find inbox for email address: {email_address}")
            return

        # Process each date in the backup_dates list
        for backup_date_str in backup_dates:
            # Convert backup_date string to datetime object
            try:
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
            print(f"Found {total_messages} messages received on {backup_date.strftime('%Y-%m-%d')} in inbox '{email_address}'.")

            for idx, message in enumerate(filtered_messages):
                try:
                    # Construct a filename for the email
                    received_time = message.ReceivedTime.strftime('%Y-%m-%d_%H-%M-%S')
                    subject = sanitize_filename(message.Subject) or "No Subject"
                    sender = sanitize_filename(message.SenderName) or "Unknown Sender"
                    filename = f"{received_time} - {sender} - {subject}.msg"

                    # Ensure the filename is not too long
                    if len(filename) > 255:
                        filename = filename[:250] + ".msg"

                    # Define the full path to save the email
                    full_path = os.path.join(save_directory, filename)

                    # Check if file already exists
                    if os.path.exists(full_path):
                        # Append a counter to the filename
                        counter = 1
                        base_filename, ext = os.path.splitext(filename)
                        while os.path.exists(full_path):
                            filename = f"{base_filename}_{counter}{ext}"
                            full_path = os.path.join(save_directory, filename)
                            counter += 1

                    # Save the email as a .msg file
                    message.SaveAs(full_path, 3)  # 3 refers to the MSG format

                    print(f"[{idx+1}/{total_messages}] Saved email: {filename}")

                except Exception as e:
                    print(f"Failed to save email '{message.Subject}': {str(e)}")
                    continue

            print(f"Backup for {backup_date.strftime('%Y-%m-%d')} completed successfully.\n")

    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    # Hardcoded email address of the inbox
    email_address = "your_inbox_email@example.com"  # Replace with your inbox email address

    # Hardcoded backup root directory
    backup_root_directory = r"C:\EmailBackups"  # Replace with your desired backup root directory

    # Prompt the user for dates
    print("Enter the dates for which you want to backup emails.")
    print("You can enter multiple dates separated by commas.")
    print("Dates should be in YYYY-MM-DD format.")
    dates_input = input("Enter date(s): ")

    # Split the input into a list of dates
    backup_dates = [date.strip() for date in dates_input.split(",")]

    if not backup_dates or backup_dates == ['']:
        print("No dates entered. Exiting the script.")
    else:
        backup_inbox(email_address, backup_root_directory, backup_dates)
