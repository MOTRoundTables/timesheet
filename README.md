# Timesheet Automation Project

This project provides a Python script (`timesheet_filler.py`) to automate the process of filling out an online timesheet using Selenium. It reads timesheet data from an Excel file and inputs it into a web-based timesheet system.

## Features

*   **Automated Login:** Logs into the timesheet system using provided credentials.
*   **Excel Data Input:** Reads daily work entries (start time, end time, notes) from a structured Excel file.
*   **Dynamic Row Addition:** Automatically adds new rows on the timesheet webpage for multiple entries on the same day.
*   **Robust Field Filling:** Uses Selenium with explicit waits and JavaScript execution for reliable data entry.

## How to Use

1.  **Prepare your Excel File:**
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

2.  **Install Dependencies:**
    Make sure you have Python installed. Then, install the required libraries:
    ```bash
    pip install pandas selenium webdriver-manager openpyxl
    ```

3.  **Run the Script:**
    Execute the script from your command line, providing the path to your Excel file, your username, and your password as arguments:

    ```bash
    python C:\Users\Golan-New_PC\timesheet\timesheet_filler.py "C:\path\to\your\timesheet.xlsx" "your_username" "your_password"
    ```
    *   Replace `"C:\path\to\your\timesheet.xlsx"` with the actual absolute path to your Excel file.
    *   Replace `"your_username"` with your actual login username for the timesheet system.
    *   Replace `"your_password"` with your actual login password for the timesheet system.

## Important Note

This script will fill in all the timesheet entries based on your Excel file. **It will NOT automatically click the "Submit" or "Save" button on the webpage.** After the script finishes, you will need to manually review the entries on the webpage and click the appropriate button to finalize your timesheet submission.

```
