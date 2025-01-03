#!/usr/bin/env python3

import json
import re
import datetime
from collections import defaultdict

import pandas as pd

def parse_dotnet_date(raw_date):
    """
    Convert a .NET date string of the format /Date(XXXXXXXXXXXX+Offset)/
    into a Python datetime object in UTC.
    
    Example date string: /Date(1672531200000+0530)/
    """
    pattern = r"/Date\((\d+)([\+\-]\d+)?\)/"
    match = re.match(pattern, raw_date)
    if match:
        ms_timestamp = int(match.group(1))  # Milliseconds since Unix epoch
        return datetime.datetime.utcfromtimestamp(ms_timestamp / 1000)
    else:
        return None

def summarize_logs(log_file_path, output_excel_path):
    """
    Reads a log file with JSON entries of user-activity data, parses it,
    aggregates usage by date and application, and saves the result (plus
    additional analyses) to an Excel file.
    """
    # 1. Load the raw text from the log file.
    with open(log_file_path, 'r', encoding='utf-8') as f:
        raw_data = f.read()

    # 2. Convert the raw data to a Python object (list of dicts).
    try:
        logs = json.loads(raw_data)
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return

    # 3. Collect timestamps per app
    usage_times = defaultdict(list)
    for entry in logs:
        app_name = entry.get("AppName", "Unknown")
        raw_date_string = entry.get("LastUsedDateTime", "")
        dt = parse_dotnet_date(raw_date_string)
        if dt is not None:
            usage_times[app_name].append(dt)

    # 4. Sort timestamps for each app
    for app in usage_times:
        usage_times[app].sort()

    # 5. Build a structure to hold daily usage statistics
    #    aggregator[ date_string ][ app_name ] = {
    #        "count": int,
    #        "total_seconds": float
    #    }
    aggregator = defaultdict(lambda: defaultdict(lambda: {"count": 0, "total_seconds": 0.0}))

    # For each app, iterate over timestamps in sorted order
    for app, sorted_timestamps in usage_times.items():
        for i in range(len(sorted_timestamps)):
            current_dt = sorted_timestamps[i]
            current_date_str = current_dt.strftime("%Y-%m-%d")

            # Increase the usage count
            aggregator[current_date_str][app]["count"] += 1

            # Look ahead to the next timestamp if it exists
            if i < len(sorted_timestamps) - 1:
                next_dt = sorted_timestamps[i + 1]
                next_date_str = next_dt.strftime("%Y-%m-%d")
                # If the next event is on the same date for the same app,
                # we approximate usage as the difference between the two times.
                if next_date_str == current_date_str:
                    delta_seconds = (next_dt - current_dt).total_seconds()
                    if delta_seconds > 0:
                        aggregator[current_date_str][app]["total_seconds"] += delta_seconds

    # 6. Convert the aggregator into a list of rows for a DataFrame
    rows = []
    for date_str, apps_info in aggregator.items():
        for app, stats in apps_info.items():
            rows.append({
                "date": date_str,
                "app_name": app,
                "usage_count": stats["count"],
                "total_seconds": stats["total_seconds"]
            })

    df = pd.DataFrame(rows)

    # Convert `date` column to an actual date type (not strictly necessary, but helpful)
    df["date"] = pd.to_datetime(df["date"], format="%Y-%m-%d")

    # 7. OPTIONAL: Additional analysis  
    # Example pivot tables:

    # 7a. Pivot for usage frequency: Each row is date, columns are apps, values are usage_count
    pivot_count = pd.pivot_table(
        df,
        index="date",
        columns="app_name",
        values="usage_count",
        aggfunc="sum",
        fill_value=0
    )

    # 7b. Pivot for total usage seconds: Each row is date, columns are apps, values are total_seconds
    pivot_seconds = pd.pivot_table(
        df,
        index="date",
        columns="app_name",
        values="total_seconds",
        aggfunc="sum",
        fill_value=0
    )

    # 7c. You might want total usage time per day, per app:
    #     We'll convert total_seconds to hours for easier reading
    pivot_hours = pivot_seconds / 3600.0

    # 8. Write everything to an Excel file with multiple sheets
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        # Sheet 1: Raw Daily Summaries
        df.to_excel(writer, sheet_name="Daily_Summary", index=False)

        # Sheet 2: Pivot of Frequency
        pivot_count.to_excel(writer, sheet_name="Pivot_Frequency")

        # Sheet 3: Pivot of Usage (seconds)
        pivot_seconds.to_excel(writer, sheet_name="Pivot_Seconds")

        # Sheet 4: Pivot of Usage (hours)
        pivot_hours.to_excel(writer, sheet_name="Pivot_Hours")

    print(f"Excel report generated: {output_excel_path}")


if __name__ == "__main__":
    # Provide the path to your log file and desired output Excel file.
    LOG_FILE_PATH = "activity_log.json"
    OUTPUT_EXCEL_PATH = "usage_summary.xlsx"

    summarize_logs(LOG_FILE_PATH, OUTPUT_EXCEL_PATH)
