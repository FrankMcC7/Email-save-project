#!/usr/bin/env python3

import json
import re
import datetime
from collections import defaultdict

import pandas as pd

def parse_dotnet_date(raw_date):
    """
    Convert a .NET date string of the format /Date(XXXXXXXXXXXX+Offset)/
    into a Python datetime object (UTC-based by default).
    
    Example date string: /Date(1672531200000+0530)/
    """
    pattern = r"/Date\((\d+)([\+\-]\d+)?\)/"
    match = re.match(pattern, raw_date)
    if match:
        ms_timestamp = int(match.group(1))  # Milliseconds since Unix epoch
        return datetime.datetime.utcfromtimestamp(ms_timestamp / 1000)
    else:
        return None


def build_focus_timeline(log_file_path):
    """
    Reads the log file and returns a list of (timestamp, app_name)
    events sorted in ascending order. 'Idle' and 'Lockapp' are
    preserved to detect breaks in usage.

    Returns:
        events: List of tuples [(dt, app_name), ...] in ascending time.
    """
    # 1. Load the raw log file
    with open(log_file_path, 'r', encoding='utf-8') as f:
        raw_data = f.read()

    # 2. Parse JSON
    try:
        logs = json.loads(raw_data)
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return []

    # 3. Collect all events: (datetime, app_name)
    events = []
    for entry in logs:
        app_name = entry.get("AppName", "Unknown")
        raw_dt = entry.get("LastUsedDateTime", "")
        dt = parse_dotnet_date(raw_dt)
        if dt is not None:
            events.append((dt, app_name))

    # 4. Sort events by datetime
    events.sort(key=lambda x: x[0])
    return events


def compute_usage(events):
    """
    Given a sorted list of (datetime, app_name) events, compute usage
    intervals based on focus transitions. The user is considered actively
    using an app from the time it appears until the next event signals
    a different app or an 'Idle'/'Lockapp' event.

    Returns:
        usage_records: list of dicts like:
            {
                'start': datetime,
                'end': datetime,
                'app': str
            }
        representing each usage session.
    """
    usage_records = []

    active_app = None
    active_start = None

    for i in range(len(events)):
        current_dt, current_app = events[i]

        if active_app is None:
            # No app is currently active:
            # If the current app is not Idle/Lockapp, it becomes active.
            if current_app.lower() not in ("idle", "lockapp"):
                active_app = current_app
                active_start = current_dt
        else:
            # We have an active app
            if current_app.lower() == active_app.lower():
                # Same app repeated or extra logs for the same app
                # Do nothing yet; usage continues.
                pass
            elif current_app.lower() in ("idle", "lockapp"):
                # Idle/Lock event => usage for active app ends now
                usage_records.append({
                    "start": active_start,
                    "end": current_dt,
                    "app": active_app
                })
                active_app = None
                active_start = None
            else:
                # Different app => user switched
                # Close the previous app usage
                usage_records.append({
                    "start": active_start,
                    "end": current_dt,
                    "app": active_app
                })
                # Start new app usage
                active_app = current_app
                active_start = current_dt

    # If an app remained active until the last event, close it at that last event time.
    # (Alternatively, you could leave "end" as None to avoid possible overestimation.)
    if active_app is not None and active_start is not None:
        last_dt, _ = events[-1]
        usage_records.append({
            "start": active_start,
            "end": last_dt,
            "app": active_app
        })

    return usage_records


def break_across_midnights(usage_records):
    """
    Splits usage sessions that cross midnight into separate
    records for each day.

    usage_records: list of dicts with 'start', 'end', 'app'.

    Returns:
        new_records: an updated list where any session that spans
                     midnight is split at the 00:00 boundary.
    """
    new_records = []
    for rec in usage_records:
        start = rec["start"]
        end = rec["end"]
        app = rec["app"]

        current_start = start
        while True:
            # Next midnight after current_start
            next_midnight = (current_start + datetime.timedelta(days=1)).replace(
                hour=0, minute=0, second=0, microsecond=0
            )
            if end <= next_midnight:
                # The entire remaining usage is before the next midnight
                new_records.append({
                    "start": current_start,
                    "end": end,
                    "app": app
                })
                break
            else:
                # The session crosses midnight; split here
                new_records.append({
                    "start": current_start,
                    "end": next_midnight,
                    "app": app
                })
                current_start = next_midnight

    return new_records


def create_daily_summary(usage_records):
    """
    Converts usage records into a daily summary DataFrame:
      - usage_count: how many times an app session starts per day
      - total_seconds: sum of session durations per day/app

    Returns:
        df: DataFrame with columns [date, app_name, usage_count, total_seconds]
    """
    aggregator = defaultdict(lambda: defaultdict(lambda: {"count": 0, "total_seconds": 0.0}))

    for rec in usage_records:
        start = rec["start"]
        end = rec["end"]
        app = rec["app"]

        # Date for aggregation is based on the start time
        date_str = start.strftime("%Y-%m-%d")
        aggregator[date_str][app]["count"] += 1
        delta_sec = (end - start).total_seconds()
        if delta_sec > 0:
            aggregator[date_str][app]["total_seconds"] += delta_sec

    # Convert to a DataFrame
    rows = []
    for date_str, apps_info in aggregator.items():
        for app_name, stats in apps_info.items():
            rows.append({
                "date": date_str,
                "app_name": app_name,
                "usage_count": stats["count"],
                "total_seconds": stats["total_seconds"]
            })
    df = pd.DataFrame(rows)
    df["date"] = pd.to_datetime(df["date"], format="%Y-%m-%d")
    return df


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Analyze app usage logs with Idle/Lock detection.")
    parser.add_argument("--input", default="activity_log.txt", help="Path to the .txt (or .json) log file.")
    parser.add_argument("--output", default="usage_summary.xlsx", help="Output Excel file path.")
    parser.add_argument("--split_midnight", action="store_true",
                        help="If set, usage that crosses midnight is split into separate daily records.")
    args = parser.parse_args()

    # 1. Build a chronological list of (datetime, app) events
    events = build_focus_timeline(args.input)
    if not events:
        print("No events found or failed to parse JSON.")
        return

    # 2. Compute usage sessions based on focus transitions
    usage_records = compute_usage(events)

    # 3. Split sessions that cross midnight (if requested)
    if args.split_midnight:
        usage_records = break_across_midnights(usage_records)

    # 4. Create daily summary
    df = create_daily_summary(usage_records)

    # 5. Create pivot tables
    pivot_freq = pd.pivot_table(
        df,
        index="date",
        columns="app_name",
        values="usage_count",
        aggfunc="sum",
        fill_value=0
    )

    pivot_seconds = pd.pivot_table(
        df,
        index="date",
        columns="app_name",
        values="total_seconds",
        aggfunc="sum",
        fill_value=0
    )

    pivot_hours = pivot_seconds / 3600.0

    # 6. Write everything to Excel (multiple sheets)
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Daily_Summary", index=False)
        pivot_freq.to_excel(writer, sheet_name="Pivot_Frequency")
        pivot_seconds.to_excel(writer, sheet_name="Pivot_Seconds")
        pivot_hours.to_excel(writer, sheet_name="Pivot_Hours")

    print(f"Usage summary has been saved to {args.output}")


if __name__ == "__main__":
    main()
