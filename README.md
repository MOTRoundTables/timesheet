# Timesheet Automation Project

This project provides a Python script (`timesheet_filler.py`) to automate the process of filling out an online timesheet using Selenium. It reads timesheet data from an Excel file and inputs it into a web-based timesheet system. It also includes a feature to import events from Google Calendar into the Excel file.

## Features

*   **Automated Login:** Logs into the timesheet system using provided credentials.
*   **Excel Data Input:** Reads daily work entries (start time, end time, notes) from a structured Excel file.
*   **Google Calendar Integration:** Imports events from your Google Calendar, including start time, end time, and title, and appends them to the Excel file.

    *   **Conflict Resolution:** When importing Google Calendar events, the system intelligently handles time overlaps with existing Excel entries. A graphical dialog will appear, allowing you to choose which events to keep (new, existing, both, or neither) using checkboxes, ensuring no data is overwritten without your explicit consent.

*   **Dynamic Row Addition:** Automatically adds new rows on the timesheet webpage for multiple entries on the same day.
*   **Robust Field Filling:** Uses Selenium with explicit waits and JavaScript execution for reliable data entry.

## How to Use

### 1. Prepare your Excel File

*   Ensure your Excel file (`.xlsx`) has the following columns (Hebrew names as used in the script):
    *   `שנה` (Year)
    *   `חודש` (Month)
    *   `יום` (Day)
    *   `זמן התחלה` (Start Time - e.g., "09:00")
    *   `זמן סיום` (End Time - e.g., "17:00")
    *   `שעות` (Hours - this column is read but not directly used for input, calculated by the system)
    *   `מה` (Notes/Description of work)
*   Example data:
    | שנה | חודש | יום | זמן התחלה | זמן סיום | שעות | מה |
    |-----|------|-----|------------|----------|------|----|
    | 2024| 6    | 25  | 09:00      | 13:00    | 4    | Project A |
    | 2024| 6    | 25  | 14:00      | 17:00    | 3    | Project B |

### 2. Install Dependencies

Make sure you have Python installed. Then, install the required libraries:

pip install pandas selenium webdriver-manager openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib customtkinter ttkbootstrap



### 3. Set Up Google Calendar API Credentials

To use the Google Calendar integration, you need to obtain a `credentials.json` file from the Google Cloud Platform.

1.  **Go to the [Google Cloud Console](https://console.cloud.google.com/)**.
2.  **Create a new project**.
3.  **Enable the Google Calendar API**.
4.  **Create an OAuth 2.0 Client ID for a Desktop application**.
5.  **Download the `credentials.json` file** and place it in the same directory as the script.

For detailed, step-by-step instructions, please refer to the official Google documentation on [creating an OAuth 2.0 client ID](https://developers.google.com/workspace/guides/create-credentials).

### 4. Run the Application

Execute the `timesheet_gui.py` script to open the graphical user interface:

<div align="center">

## The App

<img src="https://i.imgur.com/CDFIa97.png" width="400">

This is the Timesheet Automation application.<br>
It loads credentials and file paths from `config.py` and allows integration with Google Calendar.

</div>


From the GUI, you can:
*   **Import from Google Calendar:** Enable the Google Calendar integration using the toggle button, specify a date range, and import your calendar events into the Excel file. Existing Excel entries will always be preserved, and new Google Calendar events will be added on top of them, with conflict resolution for overlapping times.
*   **Run Automation:** Fill the online timesheet based on the data in your Excel file.

## Important Note

*  The script will fill in all the timesheet entries based on your Excel file. **It will NOT automatically click the "Submit" or "Save" button on the webpage.** After the script finishes, you will need to manually review the entries on the webpage and click the appropriate button to finalize your timesheet submission.


