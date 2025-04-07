import win32com.client
import os
import datetime
import sys
from pathlib import Path

def extract_attachments(account_name, inbox_name, sender_email, date_str):
    """
    Extract attachments from Outlook emails in the specified account and inbox,
    from the specified sender, on the specified date.
    
    Parameters:
    account_name (str): Name of the Outlook account
    inbox_name (str): Name of the Outlook inbox/folder
    sender_email (str): Email address of the sender
    date_str (str): Date in format 'YYYY-MM-DD'
    
    Returns:
    int: Number of attachments extracted
    """
    # Convert date string to datetime object
    try:
        target_date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
    except ValueError:
        print(f"Error: Invalid date format. Please use YYYY-MM-DD format.")
        return 0
    
    # Hardcoded save location - modify as needed
    save_location = r"C:\EmailAttachments"
    
    # Create save location if it doesn't exist
    if not os.path.exists(save_location):
        os.makedirs(save_location)
        
    # Connect to Outlook
    print("Connecting to Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"Error connecting to Outlook: {e}")
        return 0
    
    # Find the specified account
    account_found = False
    root_folder = None
    
    try:
        # List all accounts and find the requested one
        for i in range(1, namespace.Folders.Count + 1):
            current_account = namespace.Folders.Item(i)
            if current_account.Name.lower() == account_name.lower():
                root_folder = current_account
                account_found = True
                print(f"Found account: {current_account.Name}")
                break
        
        if not account_found:
            print(f"Error: Account '{account_name}' not found. Available accounts:")
            for i in range(1, namespace.Folders.Count + 1):
                print(f"  - {namespace.Folders.Item(i).Name}")
            return 0
            
        # Navigate to the specified inbox/folder
        folder = root_folder
        for folder_level in inbox_name.split('/'):
            found = False
            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_level.lower():
                    folder = subfolder
                    found = True
                    break
            if not found:
                print(f"Error: Folder '{folder_level}' not found in path '{inbox_name}'")
                print("Available folders:")
                for subfolder in folder.Folders:
                    print(f"  - {subfolder.Name}")
                return 0
    except Exception as e:
        print(f"Error accessing folder: {e}")
        return 0
    
    print(f"Searching for emails in account '{account_name}', folder '{inbox_name}' from '{sender_email}' on {date_str}...")
    
    # Filter emails
    attachment_count = 0
    try:
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)  # Sort by received time
        
        for message in messages:
            # Check if the message was received on the target date
            message_date = message.ReceivedTime.date()
            if message_date == target_date:
                # Check if the sender matches
                if sender_email.lower() in str(message.SenderEmailAddress).lower():
                    print(f"Found matching email: '{message.Subject}' received at {message.ReceivedTime}")
                    
                    # Process attachments
                    if message.Attachments.Count > 0:
                        for attachment in message.Attachments:
                            # Create a unique filename with timestamp
                            timestamp = message.ReceivedTime.strftime('%H%M%S')
                            safe_filename = f"{timestamp}_{attachment.FileName}"
                            file_path = os.path.join(save_location, safe_filename)
                            
                            # Save the attachment
                            try:
                                attachment.SaveAsFile(file_path)
                                print(f"  Saved attachment: {safe_filename}")
                                attachment_count += 1
                            except Exception as e:
                                print(f"  Error saving attachment '{attachment.FileName}': {e}")
                    else:
                        print("  No attachments found in this email.")
    except Exception as e:
        print(f"Error processing emails: {e}")
    
    print(f"\nExtraction complete. {attachment_count} attachment(s) saved to {save_location}")
    return attachment_count

def list_outlook_accounts():
    """List all available Outlook accounts"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("\nAvailable Outlook accounts:")
        for i in range(1, namespace.Folders.Count + 1):
            print(f"  - {namespace.Folders.Item(i).Name}")
        print()
    except Exception as e:
        print(f"Error listing Outlook accounts: {e}")

def main():
    """Main function to get user input and extract attachments"""
    if len(sys.argv) == 5:
        # Get parameters from command line arguments
        account_name = sys.argv[1]
        inbox_name = sys.argv[2]
        sender_email = sys.argv[3]
        date_str = sys.argv[4]
    else:
        # Get parameters from user input
        print("Outlook Email Attachment Extractor")
        print("==================================")
        
        # List available accounts
        list_outlook_accounts()
        
        account_name = input("Enter Outlook account name: ")
        inbox_name = input("Enter inbox/folder path (e.g., 'Inbox' or 'Inbox/Subfolder'): ")
        sender_email = input("Enter sender email address: ")
        date_str = input("Enter date (YYYY-MM-DD): ")
    
    # Extract attachments
    extract_attachments(account_name, inbox_name, sender_email, date_str)

if __name__ == "__main__":
    main()
