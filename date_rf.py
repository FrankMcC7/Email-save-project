The user is asking for a new Python script that can extract date information (period, month, year) from the email subject column of their data. This is likely a follow-up to our previous script that extracted sender email addresses, subjects, and other information from Outlook.

I'll create a script that:
1. Reads an Excel file (presumably created by our previous script)
2. Extracts date information from the "Subject" column using regular expressions
3. Adds new columns for the extracted period, month, and year
4. Saves the updated data to a new Excel file

Key considerations:
- I'll need to handle different date formats in email subjects
- Common date formats might include: MM/DD/YYYY, DD/MM/YYYY, Month YYYY, Month DD YYYY, etc.
- I should provide informative output about what was found and what wasn't
- The script should be user-friendly and handle potential errors gracefully

I'll use pandas for handling the Excel file, and re (regular expressions) for extracting the dates. I'll also use dateutil.parser as a backup for parsing dates that don't match our regex patterns.

# Date Extractor for Email Subjects

I'll create a Python script that extracts date information (period, month, year) from email subjects in an Excel file. This script can process the output from the previous email extraction script.

```python
import pandas as pd
import re
import os
from datetime import datetime
from dateutil import parser
import calendar

def extract_dates_from_subjects(input_file, output_file=None):
    """
    Extract date information (period, month, year) from email subjects.
    
    Args:
        input_file: Path to the Excel file containing email data
        output_file: Path to save the updated Excel file (default: auto-generated)
    """
    try:
        # Determine output filename if not provided
        if not output_file:
            base, ext = os.path.splitext(input_file)
            output_file = f"{base}_with_dates{ext}"
            
        print(f"Reading email data from {input_file}...")
        
        # Read the Excel file
        df = pd.read_excel(input_file)
        
        if 'Subject' not in df.columns:
            print("Error: Excel file does not contain a 'Subject' column.")
            return False
        
        # Initialize new columns
        df['Extracted_Period'] = None
        df['Extracted_Month'] = None
        df['Extracted_Year'] = None
        df['Extraction_Method'] = None
        
        # Count statistics
        total_rows = len(df)
        dates_found = 0
        
        print(f"Processing {total_rows} email subjects...")
        
        # Process each subject
        for idx, row in df.iterrows():
            subject = str(row['Subject']).strip()
            
            # Update progress every 100 rows
            if idx % 100 == 0 and idx > 0:
                print(f"Processed {idx} subjects ({idx/total_rows*100:.1f}%)...")
            
            # Try different patterns to extract date information
            
            # Pattern 1: Full dates like MM/DD/YYYY, MM-DD-YYYY, etc.
            date_patterns = [
                # MM/DD/YYYY or MM-DD-YYYY
                r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})',
                # YYYY/MM/DD or YYYY-MM-DD
                r'(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})',
                # DD.MM.YYYY
                r'(\d{1,2})\.(\d{1,2})\.(\d{4})',
            ]
            
            date_found = False
            
            # Try standard date patterns first
            for pattern in date_patterns:
                matches = re.search(pattern, subject)
                if matches:
                    try:
                        if len(matches.groups()) == 3:
                            # Determine which group is which (month, day, year)
                            if re.match(r'\d{4}', matches.group(1)):  # YYYY-MM-DD format
                                year = int(matches.group(1))
                                month = int(matches.group(2))
                                day = int(matches.group(3))
                            elif len(matches.group(3)) == 4:  # MM/DD/YYYY format
                                month = int(matches.group(1))
                                day = int(matches.group(2))
                                year = int(matches.group(3))
                            else:  # Handle two-digit years
                                month = int(matches.group(1))
                                day = int(matches.group(2))
                                year = int(matches.group(3))
                                if year < 100:
                                    year = 2000 + year if year < 50 else 1900 + year
                            
                            # Validate the date
                            date_obj = datetime(year, month, day)
                            
                            # Extract information
                            df.at[idx, 'Extracted_Month'] = calendar.month_name[month]
                            df.at[idx, 'Extracted_Year'] = year
                            df.at[idx, 'Extracted_Period'] = date_obj.strftime('%B %Y')
                            df.at[idx, 'Extraction_Method'] = 'Date Pattern'
                            
                            date_found = True
                            dates_found += 1
                            break
                    except (ValueError, IndexError):
                        # Invalid date, try next pattern
                        pass
            
            # If no date found, try month name patterns
            if not date_found:
                # Pattern 2: Month name + Year (e.g., "January 2023" or "Jan 2023")
                month_pattern = r'(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)[, ]+(\d{4})'
                
                matches = re.search(month_pattern, subject, re.IGNORECASE)
                if matches:
                    try:
                        month_name = matches.group(1)
                        year = int(matches.group(2))
                        
                        # Convert month name to number
                        month_dict = {month[:3].lower(): i for i, month in enumerate(calendar.month_name) if month}
                        month_num = month_dict.get(month_name[:3].lower())
                        
                        if month_num:
                            df.at[idx, 'Extracted_Month'] = calendar.month_name[month_num]
                            df.at[idx, 'Extracted_Year'] = year
                            df.at[idx, 'Extracted_Period'] = f"{calendar.month_name[month_num]} {year}"
                            df.at[idx, 'Extraction_Method'] = 'Month Year Pattern'
                            
                            date_found = True
                            dates_found += 1
                    except (ValueError, IndexError):
                        pass
            
            # If still no date found, try a more general approach
            if not date_found:
                # Use dateutil parser as a last resort
                try:
                    # Extract potential date parts from the subject
                    words = re.findall(r'\w+', subject)
                    for i in range(len(words) - 1):
                        potential_date = ' '.join(words[i:i+3])  # Try groups of 3 words
                        try:
                            parsed_date = parser.parse(potential_date, fuzzy=True)
                            
                            # Check if the parsed date is reasonable
                            if 1900 <= parsed_date.year <= datetime.now().year + 1:
                                df.at[idx, 'Extracted_Month'] = calendar.month_name[parsed_date.month]
                                df.at[idx, 'Extracted_Year'] = parsed_date.year
                                df.at[idx, 'Extracted_Period'] = parsed_date.strftime('%B %Y')
                                df.at[idx, 'Extraction_Method'] = 'Fuzzy Parsing'
                                
                                date_found = True
                                dates_found += 1
                                break
                        except:
                            continue
                except:
                    pass
        
        # Save the updated DataFrame
        print(f"Writing results to {output_file}...")
        df.to_excel(output_file, index=False)
        
        # Print statistics
        print(f"\nResults:")
        print(f"Total subjects processed: {total_rows}")
        print(f"Dates extracted: {dates_found} ({dates_found/total_rows*100:.1f}%)")
        print(f"Updated data saved to {os.path.abspath(output_file)}")
        
        return True
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return False

def main():
    print("=" * 70)
    print("Email Subject Date Extractor".center(70))
    print("=" * 70)
    
    # Get input file
    input_file = input("\nEnter the path to the Excel file with email data: ")
    
    # Validate input file
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' does not exist.")
        return
    
    # Get output file (optional)
    output_file = input("Enter the output file name (press Enter for auto-generated name): ")
    if not output_file:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_with_dates{ext}"
    
    # Process the file
    success = extract_dates_from_subjects(input_file, output_file)
    
    if success:
        print("\nDate extraction completed successfully!")
    else:
        print("\nDate extraction failed.")

if __name__ == "__main__":
    main()
```

## How to Use This Script:

1. Install the required libraries:
   ```
   pip install pandas openpyxl python-dateutil
   ```

2. Save the script to a file (e.g., `extract_dates.py`)

3. Run the script:
   ```
   python extract_dates.py
   ```

4. When prompted:
   - Enter the path to your Excel file (e.g., the file created by the previous email extraction script)
   - Optionally specify an output filename, or press Enter for an auto-generated name

## Features:

1. **Multiple Date Format Detection**:
   - Recognizes standard date formats (MM/DD/YYYY, YYYY-MM-DD, etc.)
   - Identifies month names with years (e.g., "January 2023" or "Jan 2023")
   - Uses fuzzy date parsing for less structured formats

2. **Comprehensive Output**:
   - Adds four new columns to the Excel file:
     - `Extracted_Period`: Combined month and year (e.g., "January 2023")
     - `Extracted_Month`: Full month name (e.g., "January")
     - `Extracted_Year`: Four-digit year (e.g., 2023)
     - `Extraction_Method`: How the date was identified

3. **User-Friendly Features**:
   - Progress updates during processing
   - Summary statistics after completion
   - Handles errors gracefully

This script can process the Excel file from your previous Outlook email extractor to identify and extract date information from email subjects, making it easier to categorize emails by time period.