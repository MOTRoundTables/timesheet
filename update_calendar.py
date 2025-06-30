

import datetime
import os
import sys
sys.path.append('C:/Users/Golan-New_PC/timesheet')

from google_calendar_integration import get_calendar_service, get_calendar_events, update_excel_with_calendar_events
import config

def conflict_resolution_callback_default(new_event, existing_event):
    """
    Default conflict resolution: always keep the existing event.
    """
    return ('existing',)

def main():
    """
    Main function to update Excel with Google Calendar events for the current month.
    """
    # Dynamically determine the start and end dates of the current month
    today = datetime.date.today()
    start_date = today.replace(day=1)
    # Calculate the last day of the current month
    # Add one month to the first day of the current month, then subtract one day
    next_month = start_date.replace(day=28) + datetime.timedelta(days=4)  # Go to the 28th to avoid issues with short months
    end_date = next_month - datetime.timedelta(days=next_month.day)

    print(f"Fetching Google Calendar events from {start_date} to {end_date}...")

    try:
        # Ensure credentials file exists
        if not os.path.exists('credentials.json'):
            print("Error: credentials.json not found. Please follow the setup instructions in README.md.")
            return

        service = get_calendar_service()
        events = get_calendar_events(service, start_date, end_date)

        if not events:
            print("No events found in Google Calendar for the specified date range.")
            return

        print(f"Found {len(events)} events. Updating Excel file: {config.excel_file_path}")

        change_log = update_excel_with_calendar_events(
            config.excel_file_path,
            events,
            conflict_resolution_callback_default
        )

        print("\n--- Excel Update Summary ---")
        if change_log:
            for change in change_log:
                print(f"- {change}")
        else:
            print("No changes were made to the Excel file.")
        print("--- Update complete ---")

    except Exception as e:
        print(f"\nAn error occurred during calendar import: {e}")

if __name__ == "__main__":
    main()

