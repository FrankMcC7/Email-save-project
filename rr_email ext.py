Let me update the script to meet these new requirements:

1. Allow selecting multiple accounts from the list
2. Include subject line in the output
3. Include the source account for each email
4. Save the output as an Excel file instead of CSV

For this, I'll need to:
- Modify the account selection to allow multiple selections
- Add code to extract and store the subject line along with the sender
- Include the account information for each email
- Use a library like `openpyxl` or `xlsxwriter` to create Excel files

Here's how I'll implement these changes:

# Enhanced Outlook Email Extractor

I'll create an improved script that lets you select multiple accounts, and generates an Excel file containing sender emails, subjects, and the account they came from.

```python
import win32com.client
import datetime
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def extract_sender_emails_from_outlook():
    """
    Connects to Outlook, lets user select multiple accounts/folders, 
    and extracts sender email addresses with subjects from the last 30 days.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Get all accounts
        accounts = namespace.Accounts
        account_list = []
        
        print("\nAvailable accounts:")
        print("-" * 50)
        
        # List all accounts
        for i, account in enumerate(accounts, 1):
            print(f"{i}. {account.DisplayName}")
            account_list.append(account)
        
        # Let user select multiple accounts
        print("\nSelect account numbers (comma-separated, e.g., 1,2,3 or 'all' for all accounts):")
        account_input = input("Account numbers: ").strip()
        
        selected_accounts = []
        if account_input.lower() == 'all':
            selected_accounts = account_list
        else:
            account_indices = [int(idx.strip()) - 1 for idx in account_input.split(',')]
            for idx in account_indices:
                if 0 <= idx < len(account_list):
                    selected_accounts.append(account_list[idx])
                else:
                    print(f"Warning: Account number {idx+1} is invalid and will be skipped.")
        
        if not selected_accounts:
            print("No valid accounts selected.")
            return
        
        print(f"\nSelected {len(selected_accounts)} account(s):")
        for account in selected_accounts:
            print(f"- {account.DisplayName}")
        
        # Calculate date 30 days ago
        thirty_days_ago = datetime.datetime.now() - datetime.timedelta(days=30)
        
        # Restriction for emails from the last 30 days
        restriction = "[ReceivedTime] >= '" + thirty_days_ago.strftime('%m/%d/%Y %H:%M %p') + "'"
        
        # List to store all email data
        email_data = []
        total_emails_processed = 0
        
        # Process each selected account
        for account in selected_accounts:
            print(f"\nProcessing account: {account.DisplayName}")
            
            # Get inbox for this account
            try:
                # Try to get the inbox for this specific account
                inbox = account.DeliveryStore.GetDefaultFolder(6)  # 6 represents the inbox folder
            except:
                # Fallback method if the above doesn't work
                inbox = namespace.GetDefaultFolder(6)
            
            # Get folders for selection
            folders = []
            
            # Add the main inbox
            folders.append(("Inbox", inbox))
            
            # Add subfolders in inbox
            for folder in inbox.Folders:
                folders.append((f"Inbox/{folder.Name}", folder))
            
            # List all folders
            print("\nAvailable folders:")
            print("-" * 50)
            for i, (folder_name, _) in enumerate(folders, 1):
                print(f"{i}. {folder_name}")
            
            # Let user select folders
            print("\nSelect folder numbers (comma-separated, e.g., 1,2,3 or 'all' for all folders):")
            folder_input = input("Folder numbers: ").strip()
            
            selected_folders = []
            if folder_input.lower() == 'all':
                selected_folders = folders
            else:
                folder_indices = [int(idx.strip()) - 1 for idx in folder_input.split(',')]
                for idx in folder_indices:
                    if 0 <= idx < len(folders):
                        selected_folders.append(folders[idx])
                    else:
                        print(f"Warning: Folder number {idx+1} is invalid and will be skipped.")
            
            if not selected_folders:
                print("No valid folders selected for this account. Skipping.")
                continue
            
            print(f"\nSelected {len(selected_folders)} folder(s) for {account.DisplayName}:")
            for folder_name, _ in selected_folders:
                print(f"- {folder_name}")
            
            # Process each folder
            for folder_name, folder in selected_folders:
                emails = folder.Items.Restrict(restriction)
                emails.Sort("[ReceivedTime]", True)  # Sort by received time in descending order
                
                print(f"\nProcessing folder: {folder_name}")
                folder_email_count = 0
                
                # Process each email
                for email in emails:
                    try:
                        sender_address = email.SenderEmailAddress
                        if sender_address and "@" in sender_address:
                            # Extract necessary information
                            received_time = email.ReceivedTime
                            subject = email.Subject
                            
                            # Add to our data list
                            email_data.append({
                                'account': account.DisplayName,
                                'folder': folder_name,
                                'sender': sender_address,
                                'subject': subject,
                                'received': received_time
                            })
                            
                            folder_email_count += 1
                            total_emails_processed += 1
                            
                            # Show progress every 25 emails
                            if folder_email_count % 25 == 0:
                                print(f"Processed {folder_email_count} emails in current folder...")
                    except Exception as e:
                        print(f"Error processing an email: {str(e)}")
                
                print(f"Processed {folder_email_count} emails in {folder_name}")
        
        if not email_data:
            print("\nNo emails found in the selected accounts and folders from the last 30 days.")
            return
        
        # Ask for output file name
        default_filename = f"outlook_emails_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_file = input(f"\nEnter output Excel file name (default: {default_filename}): ") or default_filename
        
        # Ensure file has .xlsx extension
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'
        
        # Create Excel file
        create_excel_file(email_data, output_file)
        
        print(f"\nExtracted data from {total_emails_processed} emails across {len(selected_accounts)} account(s)")
        print(f"Results saved to {os.path.abspath(output_file)}")
        
    except Exception as e:
        print(f"Error: {str(e)}")

def create_excel_file(email_data, output_file):
    """
    Creates an Excel file with the extracted email data.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Email Senders"
    
    # Define styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    
    # Create headers
    headers = ["Account", "Folder", "Sender Email", "Subject", "Received Date"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Add data
    for row_num, data in enumerate(email_data, 2):
        ws.cell(row=row_num, column=1).value = data['account']
        ws.cell(row=row_num, column=2).value = data['folder']
        ws.cell(row=row_num, column=3).value = data['sender']
        ws.cell(row=row_num, column=4).value = data['subject']
        ws.cell(row=row_num, column=5).value = data['received'].strftime('%Y-%m-%d %H:%M')
        
        # Apply borders to all cells
        for col_num in range(1, 6):
            ws.cell(row=row_num, column=col_num).border = thin_border
    
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
        
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = min(adjusted_width, 50)  # Cap width at 50
    
    # Freeze the header row
    ws.freeze_panes = "A2"
    
    # Create a summary sheet
    ws_summary = wb.create_sheet(title="Summary")
    
    # Summary headers
    summary_headers = ["Account", "Folder", "Email Count"]
    for col_num, header in enumerate(summary_headers, 1):
        cell = ws_summary.cell(row=1, column=col_num)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Create summary data
    summary_data = {}
    for data in email_data:
        key = (data['account'], data['folder'])
        if key in summary_data:
            summary_data[key] += 1
        else:
            summary_data[key] = 1
    
    # Add summary data
    for row_num, ((account, folder), count) in enumerate(summary_data.items(), 2):
        ws_summary.cell(row=row_num, column=1).value = account
        ws_summary.cell(row=row_num, column=2).value = folder
        ws_summary.cell(row=row_num, column=3).value = count
        
        # Apply borders
        for col_num in range(1, 4):
            ws_summary.cell(row=row_num, column=col_num).border = thin_border
    
    # Auto-adjust column widths for summary
    for col in ws_summary.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
        
        adjusted_width = max_length + 2
        ws_summary.column_dimensions[column].width = min(adjusted_width, 50)
    
    # Freeze the header row in summary
    ws_summary.freeze_panes = "A2"
    
    # Save the workbook
    wb.save(output_file)

if __name__ == "__main__":
    print("=" * 60)
    print("Outlook Email Extractor".center(60))
    print("=" * 60)
    
    extract_sender_emails_from_outlook()
    
    print("\nDone!")
```

## How to Use This Script:

1. Install the required libraries:
   ```
   pip install pywin32 openpyxl
   ```

2. Save the script to a file (e.g., `outlook_extractor.py`)

3. Run the script:
   ```
   python outlook_extractor.py
   ```

4. Follow the prompts to:
   - Select multiple accounts (comma-separated numbers or "all")
   - Select multiple folders for each account (comma-separated numbers or "all")
   - Specify the output Excel file name

## Features:

- Select multiple Outlook accounts
- Select multiple folders within each account
- Extracts emails from the last 30 days
- Includes sender email address, subject line, and account information
- Creates a professionally formatted Excel file with:
  - Main sheet containing all email details
  - Summary sheet showing email counts by account and folder
  - Auto-adjusted column widths
  - Frozen header rows
  - Proper formatting with borders

## Requirements:

- Windows operating system
- Outlook installed and configured
- Python 3.x
- pywin32 library
- openpyxl library

This enhanced script provides a comprehensive solution for extracting email sender information across multiple Outlook accounts and folders, with professional Excel output that includes all the requested information.