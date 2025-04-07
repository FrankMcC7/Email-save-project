import win32com.client
import os
import datetime
import sys
from pathlib import Path

def extract_attachments(inbox_name, sender_email, date_str):
    """
    Extract attachments from Outlook emails in the specified inbox,
    from the specified sender, on the specified date.
    
    Parameters:
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
    
    # Get the specified folder
    try:
        root_folder = namespace.Folders.Item(1)  # Default mailbox
        
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
                return 0
    except Exception as e:
        print(f"Error accessing folder '{inbox_name}': {e}")
        return 0
    
    print(f"Searching for emails in '{inbox_name}' from '{sender_email}' on {date_str}...")
    
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

def main():
    """Main function to get user input and extract attachments"""
    if len(sys.argv) == 4:
        # Get parameters from command line arguments
        inbox_name = sys.argv[1]
        sender_email = sys.argv[2]
        date_str = sys.argv[3]
    else:
        # Get parameters from user input
        print("Outlook Email Attachment Extractor")
        print("==================================")
        inbox_name = input("Enter inbox/folder path (e.g., 'Inbox' or 'Inbox/Subfolder'): ")
        sender_email = input("Enter sender email address: ")
        date_str = input("Enter date (YYYY-MM-DD): ")
    
    # Extract attachments
    extract_attachments(inbox_name, sender_email, date_str)

if __name__ == "__main__":
    main()
