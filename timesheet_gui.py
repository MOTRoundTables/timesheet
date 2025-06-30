# First, ensure you have installed the required library:
# pip install ttkbootstrap pandas openpyxl

import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap.constants import *
from tkinter import messagebox
import subprocess
import threading
import config
import os
import datetime
import pandas as pd
from google_calendar_integration import get_calendar_service, get_calendar_events, update_excel_with_calendar_events

# --- HELPER FUNCTION for resizing and centering windows ---
def center_window(window, min_width=0, min_height=0):
    window.update_idletasks()
    width = max(window.winfo_reqwidth(), min_width)
    height = max(window.winfo_reqheight(), min_height)
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

# --- EXISTING Google Calendar ConflictResolutionDialog (UNCHANGED) ---
class ConflictResolutionDialog(ttk.Toplevel):
    def __init__(self, parent, new_event, existing_event):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Google Calendar Conflict Detected")
        self.new_event = new_event
        self.existing_event = existing_event
        self.result = None
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        ttk.Label(main_frame, text=f"Conflict on {new_event['day']}/{new_event['month']}/{new_event['year']}:", font="-weight bold").pack(pady=(0, 10))
        existing_frame = ttk.LabelFrame(main_frame, text="Existing Event", padding=10)
        existing_frame.pack(fill=X, pady=5, expand=True)
        ttk.Label(existing_frame, text=f"Summary: {existing_event['summary']}", wraplength=400).pack(anchor=W, padx=5)
        ttk.Label(existing_frame, text=f"Time: {existing_event['start_time']}-{existing_event['end_time']}").pack(anchor=W, padx=5)
        self.var_existing = tk.BooleanVar(value=True)
        ttk.Checkbutton(existing_frame, text="Keep Existing Event", variable=self.var_existing).pack(anchor=W, padx=5, pady=(5,0))
        new_frame = ttk.LabelFrame(main_frame, text="New Event from Calendar", padding=10)
        new_frame.pack(fill=X, pady=5, expand=True)
        ttk.Label(new_frame, text=f"Summary: {new_event['summary']}", wraplength=400).pack(anchor=W, padx=5)
        ttk.Label(new_frame, text=f"Time: {new_event['start_time']}-{new_event['end_time']}").pack(anchor=W, padx=5)
        self.var_new = tk.BooleanVar(value=False)
        ttk.Checkbutton(new_frame, text="Keep New Event", variable=self.var_new).pack(anchor=W, padx=5, pady=(5,0))
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(15, 0), fill=X)
        ttk.Button(button_frame, text="OK", command=self.on_ok, bootstyle="primary").pack(side=RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, bootstyle="secondary").pack(side=RIGHT)
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        center_window(self, min_width=480, min_height=350)
        self.wait_window(self)
    def on_ok(self):
        selected_events = []
        if self.var_existing.get(): selected_events.append('existing')
        if self.var_new.get(): selected_events.append('new')
        self.result = tuple(selected_events)
        self.destroy()
    def on_cancel(self):
        self.result = ()
        self.destroy()

# --- ################################################################## ---
# --- ############### NEW FEATURE: EXCEL VALIDATION DIALOG ############### ---
# --- ################################################################## ---
class OverlapResolutionDialog(ttk.Toplevel):
    def __init__(self, parent, conflicting_entries_by_day):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Resolve Excel Overlaps")
        
        self.conflicting_entries_by_day = conflicting_entries_by_day
        self.day_keys = list(self.conflicting_entries_by_day.keys())
        self.current_day_index = 0
        self.modified_data = {} # Store only the changes
        self.result = None # To pass back the final data

        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        self.day_label = ttk.Label(main_frame, text="", font="-weight bold")
        self.day_label.pack(pady=(0, 10))

        # --- Treeview to display entries ---
        self.tree = ttk.Treeview(main_frame, columns=("Start", "End", "Notes"), show="headings", height=8)
        self.tree.heading("Start", text="Start Time")
        self.tree.heading("End", text="End Time")
        self.tree.heading("Notes", text="Notes")
        self.tree.column("Start", width=100, anchor=CENTER)
        self.tree.column("End", width=100, anchor=CENTER)
        self.tree.column("Notes", width=350)
        self.tree.pack(fill=BOTH, expand=True, pady=5)
        self.tree.tag_configure('overlap', background='#FADBD8') # Light red for overlaps

        # Bind double-click to edit a cell
        self.tree.bind("<Double-1>", self._on_double_click)

        # --- Navigation and Action Buttons ---
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(15, 0), fill=X)

        self.prev_button = ttk.Button(button_frame, text="< Previous Day", command=self._show_previous_day, bootstyle="secondary")
        self.prev_button.pack(side=LEFT)
        
        self.next_button = ttk.Button(button_frame, text="Next Day >", command=self._show_next_day, bootstyle="secondary")
        self.next_button.pack(side=LEFT, padx=10)

        ttk.Button(button_frame, text="Save Changes & Sort", command=self._save_changes, bootstyle="success").pack(side=RIGHT)
        ttk.Button(button_frame, text="Cancel", command=self.destroy, bootstyle="secondary").pack(side=RIGHT, padx=10)

        self._load_current_day_data()
        center_window(self, min_width=700, min_height=450)
        self.wait_window(self)

    def _load_current_day_data(self):
        # Clear previous entries
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        current_day_key = self.day_keys[self.current_day_index]
        self.day_label.config(text=f"Conflicts for: {current_day_key}")
        
        # Use modified data if it exists, otherwise use original
        entries = self.modified_data.get(current_day_key, self.conflicting_entries_by_day[current_day_key])
        
        for i, entry in enumerate(entries):
            tags = ('overlap',) if entry.get('is_overlap') else ()
            self.tree.insert("", END, iid=str(i), values=(entry['start_time'], entry['end_time'], entry['summary']), tags=tags)
        
        self._update_navigation_buttons()

    def _update_navigation_buttons(self):
        self.prev_button.config(state=NORMAL if self.current_day_index > 0 else DISABLED)
        self.next_button.config(state=NORMAL if self.current_day_index < len(self.day_keys) - 1 else DISABLED)

    def _show_previous_day(self):
        if self.current_day_index > 0:
            self.current_day_index -= 1
            self._load_current_day_data()

    def _show_next_day(self):
        if self.current_day_index < len(self.day_keys) - 1:
            self.current_day_index += 1
            self._load_current_day_data()

    def _on_double_click(self, event):
        """Handle double-click to edit a cell."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        item_id = self.tree.focus()
        column_id = self.tree.identify_column(event.x)
        col_index = int(column_id.replace('#', '')) - 1
        
        x, y, width, height = self.tree.bbox(item_id, column_id)

        # Create an Entry widget over the cell
        editor = ttk.Entry(self.tree)
        current_value = self.tree.item(item_id, "values")[col_index]
        editor.insert(0, current_value)
        editor.select_range(0, END)
        editor.focus_set()
        editor.place(x=x, y=y, width=width, height=height)

        editor.bind("<Return>", lambda e: self._save_edit(item_id, col_index, editor))
        editor.bind("<FocusOut>", lambda e: self._save_edit(item_id, col_index, editor))

    def _save_edit(self, item_id, col_index, editor):
        """Save the edited value and destroy the editor widget."""
        new_value = editor.get()
        editor.destroy()

        current_day_key = self.day_keys[self.current_day_index]
        entry_index = int(item_id)
        
        # Ensure we have a modifiable copy of the day's data
        if current_day_key not in self.modified_data:
            self.modified_data[current_day_key] = [e.copy() for e in self.conflicting_entries_by_day[current_day_key]]
        
        # Update the in-memory data
        data_key = ['start_time', 'end_time', 'summary'][col_index]
        self.modified_data[current_day_key][entry_index][data_key] = new_value
        
        # Update the Treeview display
        values = list(self.tree.item(item_id, 'values'))
        values[col_index] = new_value
        self.tree.item(item_id, values=values)

    def _save_changes(self):
        """Finalize all changes and pass them back."""
        self.result = self.modified_data
        self.destroy()

# --- ################################################################## ---
# --- ############ NEW FEATURE: EXCEL VALIDATION FUNCTIONS ############# ---
# --- ################################################################## ---
def check_excel_overlaps():
    """Initiates the Excel overlap check in a new thread."""
    check_overlap_button.config(state=DISABLED)
    output_text.insert(END, "\n--- Checking Excel for Overlapping Hours ---\n", "info")
    thread = threading.Thread(target=execute_excel_overlap_check_in_thread)
    thread.start()

def execute_excel_overlap_check_in_thread():
    """The core logic for finding and resolving overlaps."""
    try:
        excel_path = config.excel_file_path
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"Excel file not found at:\n{excel_path}")
            return

        df = pd.read_excel(excel_path)
        df.columns = ["שנה", "חודש", "יום", "זמן התחלה", "זמן סיום", "שעות", "מה"]
        
        # Convert time strings to datetime.time objects for comparison
        df['start_dt'] = pd.to_datetime(df['זמן התחלה'], format='%H:%M', errors='coerce').dt.time
        df['end_dt'] = pd.to_datetime(df['זמן סיום'], format='%H:%M', errors='coerce').dt.time
        
        # Drop rows where time conversion failed
        original_len = len(df)
        df.dropna(subset=['start_dt', 'end_dt'], inplace=True)
        if len(df) < original_len:
            output_text.insert(END, "Warning: Some rows had invalid time formats and were ignored.\n", "danger")

        conflicting_entries_by_day = {}
        for (year, month, day), day_entries in df.groupby(['שנה', 'חודש', 'יום']):
            day_entries = day_entries.sort_values(by='start_dt').reset_index()
            
            overlaps_found = False
            day_entries['is_overlap'] = False
            for i in range(len(day_entries)):
                for j in range(i + 1, len(day_entries)):
                    if day_entries.loc[i, 'end_dt'] > day_entries.loc[j, 'start_dt']:
                        overlaps_found = True
                        day_entries.loc[i, 'is_overlap'] = True
                        day_entries.loc[j, 'is_overlap'] = True
            
            if overlaps_found:
                date_key = f"{year:04d}-{month:02d}-{day:02d}"
                conflicting_entries_by_day[date_key] = [{
                    'original_index': row['index'],
                    'start_time': row['זמן התחלה'],
                    'end_time': row['זמן סיום'],
                    'summary': row['מה'],
                    'is_overlap': row['is_overlap']
                } for _, row in day_entries.iterrows()]

        if not conflicting_entries_by_day:
            output_text.insert(END, "Success: No overlapping hours found in Excel file.\n", "success")
            return

        # --- Launch Dialog and Process Results ---
        output_text.insert(END, f"Found overlaps on {len(conflicting_entries_by_day)} day(s). Opening resolution dialog...\n", "danger")
        dialog = OverlapResolutionDialog(root, conflicting_entries_by_day)
        
        if dialog.result:
            for date_key, modified_entries in dialog.result.items():
                for entry in modified_entries:
                    # Update the main DataFrame at the original index
                    df.loc[entry['original_index'], 'זמן התחלה'] = entry['start_time']
                    df.loc[entry['original_index'], 'זמן סיום'] = entry['end_time']
                    df.loc[entry['original_index'], 'מה'] = entry['summary']
            
            # Drop helper columns
            df = df.drop(columns=['start_dt', 'end_dt', 'is_overlap'], errors='ignore')
            # Sort the final DataFrame
            df = df.sort_values(by=['שנה', 'חודש', 'יום', 'זמן התחלה'])
            
            # Save back to Excel
            df.to_excel(excel_path, index=False)
            output_text.insert(END, "Excel file updated and sorted successfully.\n", "success")
        else:
            output_text.insert(END, "Overlap resolution cancelled. No changes were made.\n", "info")

    except Exception as e:
        output_text.insert(END, f"\nAn error occurred during Excel validation: {e}\n", "danger")
    finally:
        root.after(100, lambda: check_overlap_button.config(state=NORMAL))


# --- EXISTING SCRIPT FUNCTIONS (UNCHANGED) ---
def run_script():
    run_button.config(state=DISABLED)
    output_text.delete(1.0, END)
    output_text.insert(END, "--- Starting Timesheet Automation ---\n")
    output_text.insert(END, f"Using config: User={config.username}, Excel Path={config.excel_file_path}\n")
    thread = threading.Thread(target=execute_script_in_thread)
    thread.start()

def execute_script_in_thread():
    try:
        command = ["python", "C:\\Users\\Golan-New_PC\\timesheet\\timesheet_filler.py"]
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace', creationflags=subprocess.CREATE_NO_WINDOW)
        for line in iter(process.stdout.readline, ''):
            output_text.insert(END, line)
            output_text.see(END)
        process.stdout.close()
        return_code = process.wait()
        if return_code == 0:
            output_text.insert(END, "\n--- Script finished successfully! ---\n", "success")
        else:
            output_text.insert(END, f"\n--- Script finished with error code: {return_code} ---\n", "danger")
    except FileNotFoundError:
        output_text.insert(END, "\nError: 'python' command not found. Make sure Python is installed and in your system's PATH.\n", "danger")
    except Exception as e:
        output_text.insert(END, f"\nAn unexpected error occurred: {e}\n", "danger")
    finally:
        root.after(100, lambda: run_button.config(state=NORMAL))

def toggle_calendar_fields():
    state = NORMAL if calendar_var.get() else DISABLED
    start_date_entry.config(state=state)
    end_date_entry.config(state=state)
    import_calendar_button.config(state=state)
    prev_month_button.config(state=state)
    next_month_button.config(state=state)
    if calendar_var.get():
        set_current_month()

def set_current_month():
    today = datetime.date.today()
    start_of_month = today.replace(day=1)
    end_of_month = (start_of_month + datetime.timedelta(days=32)).replace(day=1) - datetime.timedelta(days=1)
    start_date_var.set(start_of_month.strftime("%Y-%m-%d"))
    end_date_var.set(end_of_month.strftime("%Y-%m-%d"))

def change_month(delta):
    try:
        start_date = datetime.datetime.strptime(start_date_var.get(), "%Y-%m-%d").date()
        new_start_of_month = (start_date.replace(day=1) + datetime.timedelta(days=delta * 32)).replace(day=1)
        new_end_of_month = (new_start_of_month + datetime.timedelta(days=32)).replace(day=1) - datetime.timedelta(days=1)
        start_date_var.set(new_start_of_month.strftime("%Y-%m-%d"))
        end_date_var.set(new_end_of_month.strftime("%Y-%m-%d"))
    except ValueError:
        messagebox.showerror("Error", "Cannot change month, current date is invalid.")

def import_calendar_data():
    if not calendar_var.get():
        messagebox.showinfo("Info", "Please enable Google Calendar integration.")
        return
    start_date_str = start_date_var.get()
    end_date_str = end_date_var.get()
    try:
        start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
        end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()
    except ValueError:
        messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD.")
        return
    import_calendar_button.config(state=DISABLED)
    output_text.insert(END, "\n--- Updating Excel from Google Calendar ---\n", "info")
    thread = threading.Thread(target=execute_calendar_import_in_thread, args=(start_date, end_date))
    thread.start()

def execute_calendar_import_in_thread(start_date, end_date):
    try:
        service = get_calendar_service()
        events = get_calendar_events(service, start_date, end_date)
        def conflict_resolution_callback(new_event, existing_event):
            dialog = ConflictResolutionDialog(root, new_event, existing_event)
            return dialog.result
        change_log = update_excel_with_calendar_events(config.excel_file_path, events, conflict_resolution_callback)
        output_text.insert(END, "\n--- Excel Update Summary ---\n", "info")
        if change_log:
            for change in change_log:
                output_text.insert(END, f"- {change}\n")
        else:
            output_text.insert(END, "No changes were made to the Excel file.\n")
        output_text.insert(END, "--- Update complete ---\n", "success")
    except Exception as e:
        output_text.insert(END, f"\nAn error occurred during calendar import: {e}\n", "danger")
    finally:
        root.after(100, lambda: import_calendar_button.config(state=NORMAL))

# --- ################################################################## ---
# --- #################### MAIN APPLICATION WINDOW ##################### ---
# --- ################################################################## ---
root = ttk.Window(themename="litera")
root.title("Timesheet Automation")

main_frame = ttk.Frame(root, padding=15)
main_frame.pack(fill=BOTH, expand=True)

# --- Configuration Frame (UNCHANGED) ---
input_frame = ttk.LabelFrame(main_frame, text="Configuration", padding=10)
input_frame.pack(fill=X, padx=5, pady=5)
info_text = f"Username: {config.username}\nPassword: ***\nExcel File: {config.excel_file_path}"
ttk.Label(input_frame, text="Credentials and file path are loaded from config.py:", justify=LEFT).pack(pady=(5,0), padx=5, anchor="w")
ttk.Label(input_frame, text=info_text, justify=LEFT, bootstyle="secondary").pack(pady=(0,5), padx=15, anchor="w")

# --- Google Calendar Integration Frame (UNCHANGED) ---
calendar_frame = ttk.LabelFrame(main_frame, text="Google Calendar Integration", padding=10)
calendar_frame.pack(fill=X, padx=5, pady=5)
calendar_var = tk.BooleanVar()
calendar_check = ttk.Checkbutton(calendar_frame, text="Enable Google Calendar", variable=calendar_var, command=toggle_calendar_fields, bootstyle="round-toggle")
calendar_check.pack(anchor="w", padx=5, pady=5)
date_frame = ttk.Frame(calendar_frame)
date_frame.pack(fill=X, padx=5, pady=(5,0))
start_date_var = tk.StringVar()
end_date_var = tk.StringVar()
prev_month_button = ttk.Button(date_frame, text="<", command=lambda: change_month(-1), state=DISABLED, bootstyle="secondary-outline")
prev_month_button.pack(side=LEFT, padx=(0, 5))
start_date_entry = ttk.Entry(date_frame, textvariable=start_date_var, state=DISABLED, width=12)
start_date_entry.pack(side=LEFT, expand=True, fill=X)
ttk.Label(date_frame, text=" to ").pack(side=LEFT, padx=5)
end_date_entry = ttk.Entry(date_frame, textvariable=end_date_var, state=DISABLED, width=12)
end_date_entry.pack(side=LEFT, expand=True, fill=X)
next_month_button = ttk.Button(date_frame, text=">", command=lambda: change_month(1), state=DISABLED, bootstyle="secondary-outline")
next_month_button.pack(side=LEFT, padx=(5, 0))
import_calendar_button = ttk.Button(calendar_frame, text="Update Excel from Google Calendar", command=import_calendar_data, state=DISABLED, bootstyle="info")
import_calendar_button.pack(pady=10, fill=X)

# --- ################################################################## ---
# --- ################# NEW FEATURE: EXCEL TOOLS FRAME ################# ---
# --- ################################################################## ---
excel_tools_frame = ttk.LabelFrame(main_frame, text="Excel Tools", padding=10)
excel_tools_frame.pack(fill=X, padx=5, pady=5)
check_overlap_button = ttk.Button(excel_tools_frame, text="Validate & Fix Excel Overlaps", command=check_excel_overlaps, bootstyle="warning")
check_overlap_button.pack(fill=X, ipady=4)

# --- Main Action Button (UNCHANGED) ---
run_button = ttk.Button(main_frame, text="Run Automation", command=run_script, bootstyle="primary")
run_button.pack(pady=15, fill=X, ipady=5)

# --- Output Frame (UNCHANGED) ---
out_frame = ttk.LabelFrame(main_frame, text="Output", padding=10)
out_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
output_text = ScrolledText(out_frame, wrap=WORD, height=10, autohide=True)
output_text.pack(fill=BOTH, expand=True, padx=5, pady=5)
output_text.tag_config("success", foreground=root.style.colors.success)
output_text.tag_config("danger", foreground=root.style.colors.danger)
output_text.tag_config("info", foreground=root.style.colors.info)

# --- Finalize Window ---
center_window(root, min_width=750, min_height=750)
root.mainloop()