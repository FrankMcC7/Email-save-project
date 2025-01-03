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
    Reads the raw log file and returns a list of (timestamp, app_name)
    events sorted in ascending order. 'Idle' and 'Lockapp' are preserved 
    as 'app_name' to detect breaks in usage.

    Returns:
        events: List of tuples [(dt, app_name), ...] in ascending time.
    """
    # 1. Load log as text
    with open(log_file_path, 'r', encoding='utf-8') as f:
        raw_data = f.read()

    # 2. Parse JSON
    try:
        logs = json.loads(raw_data)
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return []

    # 3. Collect all events into a list: (dt, app)
    events = []
    for entry in logs:
        app_name = entry.get("AppName", "Unknown")
        raw_dt = entry.get("LastUsedDateTime", "")
        dt = parse_dotnet_date(raw_dt)
        if dt is not None:
            events.append((dt, app_name))

    # 4. Sort by datetime
    events.sort(key=lambda x: x[0])

    return events


def compute_usage(events):
    """
    Given a chronological list of (datetime, app_name) events, compute 
    usage intervals by focusing on 'active app' transitions. The user is 
    considered actively using 'app_name' from the time it appears until 
    the next event or until we see an Idle/Lockapp event.

    Returns:
        usage_records: list of dicts with:
            {
                'start': datetime,
                'end': datetime,
                'app': str
            }
        each representing a usage session for an app.
    """
    usage_records = []

    active_app = None
    active_start = None

    for i in range(len(events)):
        current_dt, current_app = events[i]

        if active_app is None:
            # If no app is active and the new app is neither Idle nor Lockapp, 
            # start a usage session
            if current_app.lower() not in ("idle", "lockapp"):
                active_app = current_app
                active_start = current_dt
            # If it's idle/lock, we do nothing (no active usage).
        else:
            # We have an active app, see if this event changes it
            if current_app.lower() == active_app.lower():
                # Same app, or repeated log. We don't close anything yet.
                # We'll wait for the next event or idle/lock to break usage.
                pass
            elif current_app.lower() in ("idle", "lockapp"):
                # The user went idle or locked the screen, so the usage for the 
                # active app ends now.
                usage_records.append({
                    "start": active_start,
                    "end": current_dt,
                    "app": active_app
                })
                # No active app after this
                active_app = None
                active_start = None
            else:
                # User switched to a different app
                # So we close the usage for the old app at this moment.
                usage_records.append({
                    "start": active_start,
                    "end": current_dt,
                    "app": active_app
                })
                # Now we open a new session for the new app.
                active_app = current_app
                active_start = current_dt

    # If we still have an active app at the end, we can close it
    # at the last known event time. (Alternatively, you can do "end=None" if 
    # you prefer. But let's just use the last timestamp as the end.)
    if active_app is not None and active_start is not None:
        # We'll consider the end as the last event time
        # (or you could ignore it to avoid overestimation).
        last_dt, _ = events[-1]
        usage_records.append({
            "start": active_start,
            "end": last_dt,
            "app": active_app
        })

    return usage_records


def break_across_midnights(usage_records):
    """
    Optionally, split usage records that cross midnight into two 
    or more separate records (one per day). This helps keep daily 
    usage sums accurate.

    usage_records: list of dicts with 'start', 'end', and 'app'.

    Returns:
        new_records: updated list where any record crossing midnight is 
                     split at the 00:00 boundary.
    """
    new_records = []
    for rec in usage_records:
        start = rec["start"]
        end = rec["end"]
        app = rec["app"]

        # We'll iterate day by day from start to end, splitting as needed
        current_start = start

        while True:
            # The boundary is the next midnight after current_start
            midnight = (current_start + datetime.timedelta(days=1)).replace(
                hour=0, minute=0, second=0, microsecond=0
            )
            if end <= midnight:
                # entire usage is within the same day
                new_records.append({
                    "start": current_start,
                    "end": end,
                    "app": app
                })
                break
            else:
                # usage spans beyond midnight, split it
                new_records.append({
                    "start": current_start,
                    "end": midnight,
                    "app": app
                })
                current_start = midnight

    return new_records


def create_daily_summary(usage_records):
    """
    Convert usage records into a daily summary DataFrame:
      - usage_count: how many times an app session starts per day
      - total_seconds: sum of session durations per day/app

    Returns:
        df: pandas DataFrame with columns [date, app_name, usage_count, total_seconds].
    """
    from collections import defaultdict
    import pandas as pd

    # aggregator[date][app_name] = {"count": X, "total_seconds": Y}
    aggregator = defaultdict(lambda: defaultdict(lambda: {"count": 0, "total_seconds": 0.0}))

    for rec in usage_records:
        start = rec["start"]
        end = rec["end"]
        app = rec["app"]

        date_str = start.strftime("%Y-%m-%d")
        aggregator[date_str][app]["count"] += 1  # each record is a "session"
        delta_seconds = (end - start).total_seconds()
        if delta_seconds > 0:
            aggregator[date_str][app]["total_seconds"] += delta_seconds

    # Convert aggregator to a DataFrame
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
    # Convert date to datetime dtype
    df["date"] = pd.to_datetime(df["date"], format="%Y-%m-%d")

    return df


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Compute usage from logs with Idle/Lock detection.")
    parser.add_argument("--input", default="activity_log.txt", help="Path to the log file (txt/JSON).")
    parser.add_argument("--output", default="usage_summary.xlsx", help="Output Excel file.")
    parser.add_argument("--split_midnight", action="store_true",
                        help="If set, will split usage that crosses midnight into separate daily records.")
    args = parser.parse_args()

    # 1. Build the chronological list of events
    events = build_focus_timeline(args.input)
    if not events:
        print("No events loaded or JSON parsing failed.")
        return

    # 2. Compute usage based on focus transitions
    usage_records = compute_usage(events)

    # 3. (Optional) Split usage that crosses midnight
    if args.split_midnight:
        usage_records = break_across_midnights(usage_records)

    # 4. Create daily summary
    df = create_daily_summary(usage_records)

    # 5. Create pivot tables for analysis
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

    # 6. Write to Excel
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Daily_Summary", index=False)
        pivot_freq.to_excel(writer, sheet_name="Pivot_Frequency")
        pivot_seconds.to_excel(writer, sheet_name="Pivot_Seconds")
        pivot_hours.to_excel(writer, sheet_name="Pivot_Hours")

    print(f"Usage summary written to {args.output}")


if __name__ == "__main__":
    main()
