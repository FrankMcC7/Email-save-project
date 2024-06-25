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
                for folder in root_folder.Folders:
                    if folder.Name.lower() == "inbox":
                        inbox = folder
                        break
                if inbox is not None:
                    break
            except AttributeError as e:
                logs.append(f"Error accessing inbox: {str(e)}")
                continue

    if inbox is None:
        logs.append(f"No Inbox found for the account with the email address: {email_address}")
        pythoncom.CoUninitialize()
        with open(LOG_FILE_PATH, 'w', encoding='utf-8') as f:
            for log in logs:
                f.write(f"{log}\n")
        return

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    items = items.Restrict(f"[ReceivedTime] >= '{specific_date.strftime('%m/%d/%Y')} 00:00 AM' AND [ReceivedTime] <= '{specific_date.strftime('%m/%d/%Y')} 11:59 PM'")

    total_emails = 0
    saved_default = 0
    saved_actual = 0
    not_saved = 0
    failed_emails = []

    for item in items:
        total_emails += 1
        retries = 3
        processed = False

        if hasattr(item, 'SenderEmailAddress') or hasattr(item, 'Sender'):
            sender_email = item.SenderEmailAddress.lower() if hasattr(item, 'SenderEmailAddress') else item.Sender.Address.lower()

        while retries > 0 and not processed:
            try:
                if not sender_email:
                    logs.append(f"Error: Email item has no sender address.")
                    failed_emails.append({'email_address': 'Unknown', 'subject': item.Subject})
                    not_saved += 1
                    break

                sender_rows = sender_path_table[sender_path_table['sender'].str.lower() == sender_email.lower()]
                if sender_rows.empty:
                    # If no matching sender, save in default path
                    year, month = extract_year_and_month(item.Subject, default_year)
                    year_month_path = os.path.join(DEFAULT_SAVE_PATH, sender_email, year, month if month else "")
                    if not os.path.exists(year_month_path):
                        os.makedirs(year_month_path)
                    subject = sanitize_filename(item.Subject)
                    filename = f"{subject}.msg"
                    try:
                        item.SaveAs(os.path.join(year_month_path, filename), 3)
                        logs.append(f"Saved: {filename} to {year_month_path}")
                        saved_default += 1
                        processed = True
                    except Exception as save_err:
                        logs.append(f"Failed to save email to default path: {str(save_err)}")
                        failed_emails.append({'email_address': sender_email, 'subject': item.Subject})
                        not_saved += 1
                    break

                matched_row = None
                if len(sender_rows) > 1:
                    for _, row in sender_rows.iterrows():
                        coper_name = str(row['coper_name']).strip().lower()
                        if coper_name in item.Subject.lower():
                            matched_row = row
                            break
                else:
                    matched_row = sender_rows.iloc[0]

                if matched_row is None:
                    # If no coper_name matched, save in default path
                    year, month = extract_year_and_month(item.Subject, default_year)
                    year_month_path = os.path.join(DEFAULT_SAVE_PATH, sender_email, year, month if month else "")
                    if not os.path.exists(year_month_path):
                        os.makedirs(year_month_path)
                    subject = sanitize_filename(item.Subject)
                    filename = f"{subject}.msg"
                    try:
                        item.SaveAs(os.path.join(year_month_path, filename), 3)
                        logs.append(f"Saved: {filename} to {year_month_path}")
                        saved_default += 1
                        processed = True
                    except Exception as save_err:
                        logs.append(f"Failed to save email to default path: {str(save_err)}")
                        failed_emails.append({'email_address': sender_email, 'subject': item.Subject})
                        not_saved += 1
                else:
                    base_path = matched_row['base_path']
                    keyword_path = matched_row['keyword_path']
                    keywords = str(matched_row['keywords']).split(';')
                    keywords = [kw.lower().strip() for kw in keywords]  # Ensure keywords are lowercase and stripped

                    keyword_matched = any(keyword in item.Subject.lower() for keyword in keywords)

                    if not keyword_matched and item.Attachments.Count > 0:
                        for attachment in item.Attachments:
                            if any(keyword in attachment.FileName.lower() for keyword in keywords):
                                keyword_matched = True
                                break

                    if keyword_matched:
                        subject = sanitize_filename(item.Subject)
                        filename = f"{subject}_{specific_date_str}.msg"
                        try:
                            item.SaveAs(os.path.join(keyword_path, filename), 3)
                            logs.append(f"Saved keyword case: {filename} to {keyword_path}")
                            saved_actual += 1
                            processed = True
                        except Exception as save_err:
                            logs.append(f"Failed to save keyword case email: {str(save_err)}")
                            failed_emails.append({'email_address': sender_email, 'subject': item.Subject})
                            not_saved += 1
                    else:
                        year, month = extract_year_and_month(item.Subject, default_year)
                        year_month_path = os.path.join(base_path, year, month if month else "")
                        if not os.path.exists(year_month_path):
                            os.makedirs(year_month_path)
                        subject = sanitize_filename(item.Subject)
                        filename = f"{subject}.msg"
                        try:
                            item.SaveAs(os.path.join(year_month_path, filename), 3)
                            logs.append(f"Saved: {filename} to {year_month_path}")
                            saved_actual += 1
                            processed = True
                        except Exception as save_err:
                            logs.append(f"Failed to save email: {str(save_err)}")
                            failed_emails.append({'email_address': sender_email, 'subject': item.Subject})
                            not_saved += 1
                break

            except pythoncom.com_error as com_err:
                error_code, _, error_message, _ = com_err.args
                logs.append(f"COM Error handling email with subject '{item.Subject}': {error_message} (Code: {error_code})")
                retries -= 1
                if retries == 0:
                    logs.append(f"Failed to save email '{item.Subject}' after 3 retries.")
                    failed_emails.append({'email_address': sender_email, 'subject': item.Subject})
                    not_saved += 1
            except Exception as e:
                logs.append(f"Error handling email with subject '{item.Subject}': {str(e)}")
                failed_emails.append({'email_address': sender_email, 'subject': item.Subject})
                not_saved += 1
                retries = 0

        if not processed:
            year, month = extract_year_and_month(item.Subject, default_year)
            year_month_path = os.path.join(DEFAULT_SAVE_PATH, sender_email, year, month if month else "")
            if not os.path.exists(year_month_path):
                os.makedirs(year_month_path)
            subject = sanitize_filename(item.Subject)
            filename = f"{subject}.msg"
            try:
                item.SaveAs(os.path.join(year_month_path, filename), 3)
                logs.append(f"Saved: {filename} to {year_month_path}")
                saved_default += 1
            except Exception as save_err:
                logs.append(f"Failed to save email to default path: {str(save_err)}")
                failed_emails.append({'email_address': sender_email, 'subject': item.Subject})
                not_saved += 1

    pythoncom.CoUninitialize()
    with open(LOG_FILE_PATH, 'w', encoding='utf-8') as f:
        for log in logs:
            f.write(f"{log}\n")

    update_excel_summary(specific_date_str, total_emails, saved_default, saved_actual, not_saved, failed_emails)
