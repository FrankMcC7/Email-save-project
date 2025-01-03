#!/usr/bin/env python3

import json
import re
import datetime
from collections import defaultdict

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

def summarize_logs(log_file_path):
    """
    Reads a log file with JSON entries of user-activity data, parses it,
    and prints a daily summary:
    1. Count of usage events (frequency).
    2. Approximate total usage time, computed by summing consecutive usage 
       intervals for the same app within the same date.
    """
    # 1. Load the raw text from the log file.
    with open(log_file_path, 'r', encoding='utf-8') as f:
        raw_data = f.read()

    # 2. Convert the raw data to a Python object (list of dicts).
    #    Adjust if your file isn't a strict JSON array.
    try:
        logs = json.loads(raw_data)
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return

    # 3. Build a structure to hold usage times per app
    #    usage_times[app_name] = list of datetime stamps
    usage_times = defaultdict(list)

    for entry in logs:
        app_name = entry.get("AppName", "Unknown")
        raw_date_string = entry.get("LastUsedDateTime", "")
        dt = parse_dotnet_date(raw_date_string)
        if dt is not None:
            usage_times[app_name].append(dt)

    # Sort timestamps for each app so we can compute consecutive intervals
    for app in usage_times:
        usage_times[app].sort()

    # 4. Build a data structure to summarize usage by date
    #    aggregator[ date_string ][ app_name ] = {
    #        "count": int,
    #        "total_seconds": float
    #    }
    aggregator = defaultdict(lambda: defaultdict(lambda: {"count": 0, "total_seconds": 0.0}))

    # For each app, we look at consecutive timestamps in sorted order
    for app, date_list in usage_times.items():
        for i in range(len(date_list)):
            # Current usage event
            current_dt = date_list[i]
            current_date_str = current_dt.strftime("%Y-%m-%d")  # e.g. "2025-01-03"

            # Increase frequency (count of usage events) for this app/date
            aggregator[current_date_str][app]["count"] += 1

            # Estimate duration by looking at the *next* timestamp, 
            # but only if it’s the same app *and* the same date
            if i < len(date_list) - 1:
                next_dt = date_list[i+1]
                next_date_str = next_dt.strftime("%Y-%m-%d")

                # If next event is on the same date for the same app,
                # compute the difference in seconds.
                if next_date_str == current_date_str:
                    delta_seconds = (next_dt - current_dt).total_seconds()
                    # We only add this difference if it's positive (should be).
                    if delta_seconds > 0:
                        aggregator[current_date_str][app]["total_seconds"] += delta_seconds

    # 5. Print the summary grouped by date
    #    For each date, we list each app, the usage count, and total usage in HH:MM:SS
    print("========== Daily Usage Summary ==========")
    for date_str in sorted(aggregator.keys()):
        print(f"\nDate: {date_str}")
        print("-" * 40)
        for app, stats in sorted(aggregator[date_str].items()):
            count = stats["count"]
            total_sec = stats["total_seconds"]
            # Convert seconds to a human‐readable HH:MM:SS
            usage_td = datetime.timedelta(seconds=total_sec)
            print(f"  App: {app}\n"
                  f"    Frequency: {count}\n"
                  f"    Total Usage (approx): {usage_td}\n")
    print("=========================================")

if __name__ == "__main__":
    # Path to your JSON log file
    LOG_FILE_PATH = "activity_log.json"
    summarize_logs(LOG_FILE_PATH)
