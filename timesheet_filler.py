import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import getpass # For securely getting password input
from datetime import datetime
import traceback
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import config  # Import the config file
import argparse # Import argparse

def automate_timesheet(excel_file_path, username, password, dry_run=False, headless=False, job_value=None, job_name=None):
    try:
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.chrome.options import Options
        options = Options()
        options.add_experimental_option("detach", True) # This is the key to keeping the browser open
        if headless:
            options.add_argument("--headless")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print(f"Error initializing WebDriver: {e}")
        print("Please ensure Chrome is installed and the ChromeDriver is correctly set up.")
        print("You might need to update webdriver-manager or manually install ChromeDriver.")
        return # Exit the function if WebDriver fails to initialize

    try:
        driver.get("https://saas.webtime.co.il/wt_periodic.adp")

        print("Attempting to log in...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "email"))
        ).send_keys(username)

        driver.find_element(By.ID, "password").send_keys(password)
        driver.find_element(By.ID, "login-button").click()

        # Wait for successful login (e.g., wait for an element on the dashboard)
        # This ID needs to be an element that appears *after* successful login
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "tableDyn1")) # Assuming tableDyn1 is present after login
        )
        print("Login successful!")

        # Click the 'Show' button to load the timesheet
        print("Clicking 'Show' button...")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "submit1"))
        ).click()

        # Wait for the timesheet table to load after clicking 'Show'
        WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.ID, "tableDyn1")) # Wait for visibility
        )
        print("Timesheet loaded and ready for input.")

        if dry_run:
            print("--- DRY RUN MODE --- Script will not make any changes.")

        print(f"Reading data from {excel_file_path}...")
        df = pd.read_excel(excel_file_path)

        # Rename columns for easier access (using the Hebrew names you provided)
        df.columns = ["שנה", "חודש", "יום", "זמן התחלה", "זמן סיום", "שעות", "מה"]

        # Group by date to handle multiple entries per day
        grouped_by_date = df.groupby(['שנה', 'חודש', 'יום'])

        print("Starting to fill timesheet entries...")
        print(f"Total unique dates to process: {len(grouped_by_date.groups)}")
        print(f"Grouped dates: {grouped_by_date.groups.keys()}")
        for (year, month, day), day_entries in grouped_by_date:
            print(f"--- Processing date group: {year}-{month}-{day} ---")
            # Format date to match website's hidden input value (YYYY-MM-DD)
            formatted_date = f"{year:04d}-{month:02d}-{day:02d}"
            print(f"Processing entries for formatted date: {formatted_date}")

            # Find the TR element for this specific date
            date_row_element = None
            date_inputs = driver.find_elements(By.XPATH, "//input[starts-with(@id, 'day_')]")
            for date_input in date_inputs:
                if date_input.get_attribute('value') == formatted_date:
                    # Get the parent TR element that has the row_no attribute
                    date_row_element = date_input.find_element(By.XPATH, "./ancestor::tr[@row_no]")
                    break

            if not date_row_element:
                print(f"Could not find row for date {formatted_date}. Skipping.")
                continue

            # Scroll the date row into view
            driver.execute_script("arguments[0].scrollIntoView(true);", date_row_element)
            # No time.sleep() here, as scrollIntoView is usually quick and subsequent actions have waits.
            print(f"    Date row element outerHTML: {date_row_element.get_attribute('outerHTML')}")

            # Get the base row number for this day (e.g., '1')
            base_row_num = date_row_element.get_attribute('row_no')

            # Find the 'Add Row' button for this day
            # This assumes the 'Add Row' button is an img tag with a specific onclick attribute
            add_row_button = WebDriverWait(date_row_element, 10).until(
                EC.element_to_be_clickable((By.XPATH, ".//img[contains(@onclick, 'addRow(this,true)')]"))
            )
            print(f"    Add Row button outerHTML: {add_row_button.get_attribute('outerHTML')}")

            # Determine how many additional rows are needed for this day
            num_additional_rows_needed = len(day_entries) - 1
            print(f"    {num_additional_rows_needed} additional rows needed for {formatted_date}")

            # Add the necessary rows first
            for _ in range(num_additional_rows_needed):
                print(f"    Clicking 'Add Row' for {formatted_date}...")
                existing_row_suffixes_for_date = set()
                date_rows_before_add = driver.find_elements(By.XPATH, f"//tr[@row_no][.//input[starts-with(@id, 'day_') and @value='{formatted_date}']]")
                for row in date_rows_before_add:
                    existing_row_suffixes_for_date.add(row.get_attribute('row_no'))

                if not dry_run:
                    driver.execute_script("arguments[0].click();", add_row_button)

                try:
                    WebDriverWait(driver, 20).until(
                        lambda d: len(d.find_elements(By.XPATH, f"//tr[@row_no][.//input[starts-with(@id, 'day_') and @value='{formatted_date}']]")) > len(existing_row_suffixes_for_date)
                    )
                    print(f"    New row added for {formatted_date}.")
                except TimeoutException:
                    print(f"    Timeout waiting for new row for date {formatted_date} to appear after clicking 'Add Row'.")
                    # Decide how to handle this: continue or raise an error. For now, continue.
                    continue

            # Collect all row suffixes for this date after adding all rows
            all_row_elements_for_date = driver.find_elements(By.XPATH, f"//tr[@row_no][.//input[starts-with(@id, 'day_') and @value='{formatted_date}']]")
            # Sort them by their 'row_no' attribute to ensure correct order
            all_row_elements_for_date.sort(key=lambda x: int(x.get_attribute('row_no')))

            # Now, iterate through the entries and populate the fields
            for i, entry_row in enumerate(day_entries.iterrows()):
                # Access columns by their Hebrew names
                start_time = entry_row[1]["זמן התחלה"]
                end_time = entry_row[1]["זמן סיום"]
                notes = entry_row[1]["מה"]

                print(f"  Entry {i+1}: {start_time}-{end_time} - {notes}")

                # Get the current_row_suffix from the pre-added rows
                if i < len(all_row_elements_for_date):
                    current_row_element = all_row_elements_for_date[i]
                    current_row_suffix = current_row_element.get_attribute('row_no')
                    print(f"    Using row suffix: {current_row_suffix} for entry {i+1}")
                else:
                    print(f"    Error: Not enough rows found for entry {i+1} on {formatted_date}. Skipping.")
                    continue

                # Construct dynamic IDs
                start_hh_id = f"time_start_HH_{current_row_suffix}"
                start_mm_id = f"time_start_MM_{current_row_suffix}"
                end_hh_id = f"time_end_HH_{current_row_suffix}"
                end_mm_id = f"time_end_MM_{current_row_suffix}"
                notes_button_id = f"detailsU_CMMPAN_{current_row_suffix}_1" # ID of the notes button
                notes_field_id = f"work_comments_{current_row_suffix}_1" # ID of the actual notes textarea
                print(f"    Constructed IDs: start_hh={start_hh_id}, start_mm={start_mm_id}, end_hh={end_hh_id}, end_mm={end_mm_id}, notes_button={notes_button_id}, notes_field={notes_field_id}")

                # Fill fields
                start_hour = str(start_time).split(':')[0]
                start_minute = str(start_time).split(':')[1]
                try:
                    if not dry_run:
                        WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, start_hh_id))
                        ).send_keys(start_hour)
                        driver.find_element(By.ID, start_mm_id).send_keys(start_minute)

                        end_hour = str(end_time).split(':')[0]
                        end_minute = str(end_time).split(':')[1]
                        driver.find_element(By.ID, end_hh_id).send_keys(end_hour)
                        driver.find_element(By.ID, end_mm_id).send_keys(end_minute)

                        # Click the notes button to make the notes field visible
                        print(f"    Attempting to click notes button with ID: {notes_button_id}")
                        notes_button_element = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, notes_button_id))
                        )
                        notes_button_element.click()
                        

                        # Wait for the notes pop-up div to be visible
                        # The notes popup div ID seems to always end with _1, regardless of sub_row_no
                        notes_popup_div_id = f"detailsDiv_CMMPAN_{current_row_suffix}_1"
                        print(f"    Waiting for notes pop-up div with ID: {notes_popup_div_id} to be visible.")
                        notes_popup_div_element = WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located((By.ID, notes_popup_div_id))
                        )

                        # Use a more robust method to find the textarea and set its value
                        print(f"    Attempting to fill notes field for row suffix {current_row_suffix}")
                        
                        # Find the textarea within the visible popup. This is more robust than relying on a constructed ID.
                        notes_element = WebDriverWait(notes_popup_div_element, 10).until(
                            EC.presence_of_element_located((By.TAG_NAME, "textarea"))
                        )

                        # Use JavaScript to set the value, which is often more reliable for complex fields.
                        driver.execute_script("arguments[0].value = arguments[1];", notes_element, str(notes))

                        # Close the notes pop-up by clicking the original notes button again
                        print(f"    Attempting to close notes pop-up by clicking button: {notes_button_id}")
                        
                        # It's crucial to wait for the button to be clickable again before closing.
                        close_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, notes_button_id))
                        )
                        close_button.click()
                        
                        # Wait for the popup to become invisible to confirm it has closed
                        WebDriverWait(driver, 10).until(
                            EC.invisibility_of_element_located((By.ID, notes_popup_div_id))
                        )

                        # Fill job assignment if provided
                        if job_value and job_name:
                            print(f"    Attempting to fill job assignment for row {current_row_suffix}")
                            try:
                                # Find the assignment input field by name pattern
                                assignment_input_name = f"assignment_name_{current_row_suffix}"
                                assignment_input = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.NAME, assignment_input_name))
                                )

                                # Click the assignment field to open the options
                                print(f"    Clicking assignment field: {assignment_input_name}")
                                assignment_input.click()

                                # Wait a moment for the options to appear
                                time.sleep(0.5)

                                # Find the option with the matching value and click it
                                # The option might be in a dropdown or popup
                                option_xpath = f"//option[@value='{job_value}']"
                                option_element = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.XPATH, option_xpath))
                                )
                                print(f"    Selecting job option with value: {job_value}")
                                option_element.click()

                                # Alternatively, if the above doesn't work, try using JavaScript to set the value directly
                                # Find the hidden input that stores the job ID (jid_{row_suffix})
                                jid_input_name = f"jid_{current_row_suffix}"
                                try:
                                    jid_input = driver.find_element(By.NAME, jid_input_name)
                                    driver.execute_script("arguments[0].value = arguments[1];", jid_input, job_value)
                                    driver.execute_script("arguments[0].value = arguments[1];", assignment_input, job_name)
                                    print(f"    Job assignment filled using JavaScript")
                                except NoSuchElementException:
                                    print(f"    Warning: Could not find jid input field: {jid_input_name}")

                            except (NoSuchElementException, TimeoutException) as e:
                                print(f"    Warning: Could not fill job assignment for row {current_row_suffix}: {e}")
                                # Don't fail the entire entry if job assignment fails

                except (NoSuchElementException, TimeoutException) as e:
                    print(f"    Error filling fields for row {current_row_suffix}: {e}")
                    traceback.print_exc()
                    continue # Skip to the next entry if filling fields failed.


            print(f"Finished processing entries for date: {formatted_date}")
            print(f"    Processed all {len(day_entries)} entries for {formatted_date}")

        print("Timesheet filling complete.")

    except (NoSuchElementException, TimeoutException) as e:
        print(f"A Selenium error occurred: {e}")
        traceback.print_exc()
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        traceback.print_exc()
    finally:
        print("Timesheet filling process completed. Please review and submit manually.")
        # The following lines are removed for GUI compatibility
        # input("Press Enter to close the browser...")
        # driver.quit() # Ensure driver quits after user input

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Automate timesheet filling.')
    parser.add_argument('--headless', action='store_true', help='Run in headless mode.')
    parser.add_argument('--dry-run', action='store_true', help='Run without making any changes.')
    parser.add_argument('--excel-file', type=str, help='Path to Excel file (overrides config.py)')
    parser.add_argument('--job-value', type=str, help='Job assignment value to auto-fill')
    parser.add_argument('--job-name', type=str, help='Job assignment name to auto-fill')
    args = parser.parse_args()

    # Use Excel file path from argument if provided, otherwise from config.py
    excel_file = args.excel_file if args.excel_file else config.excel_file_path
    automate_timesheet(excel_file, config.username, config.password, dry_run=args.dry_run, headless=args.headless, job_value=args.job_value, job_name=args.job_name)
