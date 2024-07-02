def save_email(item, save_path, special_case):
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    
    valid_extensions = ('.xlsx', '.xls', '.csv', '.pdf', '.doc', '.docx')
    if special_case and special_case.lower() == 'yes' and item.Attachments.Count > 0:
        for attachment in item.Attachments:
            # Only consider attachments with specific file types
            if attachment.FileName.lower().endswith(valid_extensions):
                filename = sanitize_filename(os.path.splitext(attachment.Filename)[0])  # Remove extension
                filename = f"{filename}.msg"
                break
        else:
            # If no valid attachment is found, fallback to using the subject
            filename = f"{sanitize_filename(item.Subject)}.msg"
    else:
        filename = f"{sanitize_filename(item.Subject)}.msg"

    # Ensure the combined length of the save path and filename does not exceed 255 characters
    full_path = os.path.join(save_path, filename)
    if len(full_path) > 255:
        max_filename_length = 255 - len(save_path) - 1  # -1 for the path separator
        filename = filename[:max_filename_length]
        full_path = os.path.join(save_path, filename)

    item.SaveAs(full_path, 3)
    return filename


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

            filename = file.filename
            filepath = os.path.join('uploads', filename)
            file.save(filepath)
            
            try:
                sender_path_table = pd.read_csv(filepath, encoding='utf-8')
            except UnicodeDecodeError:
                sender_path_table = pd.read_csv(filepath, encoding='latin1')

            account_email_address = "hf_data@bofa.com"
            socketio.start_background_task(save_emails_from_senders_on_date, account_email_address, date_str, sender_path_table, default_year)
            return redirect(url_for('results'))

    return render_template('index.html')



def save_email(item, save_path, special_case):
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    
    valid_extensions = ('.xlsx', '.xls', '.csv', '.pdf', '.doc', '.docx')
    if special_case and special_case.lower() == 'yes' and item.Attachments.Count > 0:
        for attachment in item.Attachments:
            # Only consider attachments with specific file types
            if attachment.FileName.lower().endswith(valid_extensions):
                filename_base = sanitize_filename(os.path.splitext(attachment.Filename)[0])  # Remove extension
                filename = f"{filename_base}.msg"
                break
        else:
            # If no valid attachment is found, fallback to using the subject
            filename_base = sanitize_filename(item.Subject)
            filename = f"{filename_base}.msg"
    else:
        filename_base = sanitize_filename(item.Subject)
        filename = f"{filename_base}.msg"

    # Ensure the combined length of the save path and filename does not exceed 255 characters
    extension = ".msg"
    max_filename_length = 255 - len(save_path) - len(extension) - 1  # -1 for the path separator
    if len(filename_base) > max_filename_length:
        filename_base = filename_base[:max_filename_length]

    filename = f"{filename_base}{extension}"
    full_path = os.path.join(save_path, filename)

    item.SaveAs(full_path, 3)
    return filename

