# First, ensure you have installed the required libraries:
# pip install ttkbootstrap pandas openpyxl

import tkinter as tk
from tkinter import filedialog
import ttkbootstrap as ttk
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap.constants import *
from tkinter import messagebox
import subprocess
import threading
import config
import os
import shutil
import datetime
import pandas as pd
import calendar

# --- ################################################################## ---
# --- ###################### HELPER FUNCTIONS ########################## ---
# --- ################################################################## ---

def center_window(window, min_width=0, min_height=0):
    window.update_idletasks()
    width = max(window.winfo_reqwidth(), min_width)
    height = max(window.winfo_reqheight(), min_height)
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

def create_backup(excel_path, backup_dir):
    if not os.path.exists(excel_path): return
    try:
        os.makedirs(backup_dir, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.basename(excel_path)
        name, ext = os.path.splitext(base_name)
        backup_path = os.path.join(backup_dir, f"{name}_backup_{timestamp}{ext}")
        shutil.copy(excel_path, backup_path)
        output_text.insert(END, f"Backup created at: {backup_path}\n", "info")
    except Exception as e:
        output_text.insert(END, f"Error creating backup: {e}\n", "danger")

def calculate_hours_from_strings(start_str, end_str):
    try:
        start_time = datetime.datetime.strptime(start_str, '%H:%M')
        end_time = datetime.datetime.strptime(end_str, '%H:%M')
        if end_time < start_time: end_time += datetime.timedelta(days=1)
        duration = end_time - start_time
        return round(duration.total_seconds() / 3600, 2)
    except (ValueError, TypeError):
        return 0.0

# --- Google Calendar ConflictResolutionDialog (UNCHANGED) ---
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
        ttk.Label(main_frame, text=f"Conflict on {new_event['day']}/{new_event['month']}/{new_event['year']}:", font="-size 12 -weight bold").pack(pady=(0, 10))
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

# --- EXCEL VALIDATION DIALOG (UNCHANGED) ---
class OverlapResolutionDialog(ttk.Toplevel):
    def __init__(self, parent, conflicting_entries_by_day):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Resolve Excel Overlaps")
        self.conflicting_entries_by_day = conflicting_entries_by_day
        self.day_keys = list(self.conflicting_entries_by_day.keys())
        self.current_day_index = 0
        self.modified_data = {}
        self.result = None
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        self.day_label = ttk.Label(main_frame, text="", font="-size 12 -weight bold", anchor=CENTER)
        self.day_label.pack(fill=X, pady=(0, 10))
        self.tree = ttk.Treeview(main_frame, columns=("Start", "End", "Notes"), show="headings", height=8)
        self.tree.heading("Start", text="Start Time")
        self.tree.heading("End", text="End Time")
        self.tree.heading("Notes", text="Notes")
        self.tree.column("Start", width=100, anchor=CENTER)
        self.tree.column("End", width=100, anchor=CENTER)
        self.tree.column("Notes", width=350)
        self.tree.pack(fill=BOTH, expand=True, pady=5)
        self.tree.tag_configure('overlap', background='#FADBD8')
        self.tree.bind("<Double-1>", self._on_double_click)
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
        for item in self.tree.get_children(): self.tree.delete(item)
        current_day_key = self.day_keys[self.current_day_index]
        self.day_label.config(text=f"Conflicts for: {current_day_key}")
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
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        item_id = self.tree.focus()
        column_id = self.tree.identify_column(event.x)
        col_index = int(column_id.replace('#', '')) - 1
        x, y, width, height = self.tree.bbox(item_id, column_id)
        editor = ttk.Entry(self.tree)
        current_value = self.tree.item(item_id, "values")[col_index]
        editor.insert(0, current_value)
        editor.select_range(0, END)
        editor.focus_set()
        editor.place(x=x, y=y, width=width, height=height)
        editor.bind("<Return>", lambda e: self._save_edit(item_id, col_index, editor))
        editor.bind("<FocusOut>", lambda e: self._save_edit(item_id, col_index, editor))
    def _save_edit(self, item_id, col_index, editor):
        new_value = editor.get()
        editor.destroy()
        current_day_key = self.day_keys[self.current_day_index]
        entry_index = int(item_id)
        if current_day_key not in self.modified_data:
            self.modified_data[current_day_key] = [e.copy() for e in self.conflicting_entries_by_day[current_day_key]]
        data_key = ['start_time', 'end_time', 'summary'][col_index]
        self.modified_data[current_day_key][entry_index][data_key] = new_value
        values = list(self.tree.item(item_id, 'values'))
        values[col_index] = new_value
        self.tree.item(item_id, values=values)
    def _save_changes(self):
        self.result = self.modified_data
        self.destroy()

# --- Manual Entry Dialog (UNCHANGED) ---
class ManualEntryDialog(ttk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Add Manual Entry")
        self.result = None
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=True)
        ttk.Label(main_frame, text="Date:").pack(anchor=W)
        self.date_entry = ttk.DateEntry(main_frame, bootstyle="primary", dateformat='%Y-%m-%d')
        self.date_entry.pack(fill=X, pady=(0, 10))
        time_frame = ttk.Frame(main_frame)
        time_frame.pack(fill=X, pady=(0, 10))
        ttk.Label(time_frame, text="Start Time (HH:MM):").pack(side=LEFT)
        self.start_time_entry = ttk.Entry(time_frame, width=8)
        self.start_time_entry.pack(side=LEFT, padx=5)
        ttk.Label(time_frame, text="End Time (HH:MM):").pack(side=LEFT)
        self.end_time_entry = ttk.Entry(time_frame, width=8)
        self.end_time_entry.pack(side=LEFT, padx=5)
        ttk.Label(main_frame, text="Notes:").pack(anchor=W)
        self.notes_entry = ttk.Entry(main_frame)
        self.notes_entry.pack(fill=X, pady=(0, 10))
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=10)
        ttk.Button(button_frame, text="Save Entry", command=self._save, bootstyle="success").pack(side=RIGHT)
        ttk.Button(button_frame, text="Cancel", command=self.destroy, bootstyle="secondary").pack(side=RIGHT, padx=10)
        center_window(self, min_width=400)
        self.wait_window(self)
    def _save(self):
        try:
            date = self.date_entry.entry.get()
            start_time = self.start_time_entry.get()
            end_time = self.end_time_entry.get()
            notes = self.notes_entry.get()
            datetime.datetime.strptime(date, '%Y-%m-%d')
            datetime.datetime.strptime(start_time, '%H:%M')
            datetime.datetime.strptime(end_time, '%H:%M')
            self.result = {"date": date, "start_time": start_time, "end_time": end_time, "notes": notes}
            self.destroy()
        except ValueError:
            messagebox.showerror("Invalid Format", "Please ensure date is YYYY-MM-DD and times are HH:MM.", parent=self)

# --- ################################################################## ---
# --- ################## BACKEND FUNCTION DEFINITIONS ################## ---
# --- ################################################################## ---

def browse_excel_file():
    filepath = filedialog.askopenfilename(title="Select Excel Timesheet", filetypes=(("Excel Files", "*.xlsx"), ("All files", "*.*")))
    if filepath:
        excel_path_var.set(filepath)
        update_backup_path_default()
        update_total_hours_display()

def browse_backup_folder():
    folderpath = filedialog.askdirectory(title="Select Backup Folder")
    if folderpath:
        backup_path_var.set(folderpath)

def update_backup_path_default():
    excel_path = excel_path_var.get()
    if excel_path and os.path.exists(excel_path):
        backup_path_var.set(os.path.dirname(excel_path))
    else:
        backup_path_var.set("")

def toggle_backup_fields():
    state = NORMAL if backup_enabled_var.get() else DISABLED
    backup_path_entry.config(state=state)
    backup_browse_button.config(state=state)

def update_total_hours_display():
    excel_path = excel_path_var.get()
    if not os.path.exists(excel_path):
        total_hours_var.set("Total Hours: N/A (File not found)")
        return
    try:
        df = pd.read_excel(excel_path)
        if "שעות" in df.columns:
            total = df["שעות"].sum()
            total_hours_var.set(f"Total Hours: {total:.2f}")
        else:
            total_hours_var.set("Total Hours: N/A ('Hours' column missing)")
    except Exception as e:
        total_hours_var.set(f"Total Hours: Error")
        output_text.insert(END, f"Error reading total hours: {e}\n", "danger")

def add_manual_entry():
    dialog = ManualEntryDialog(root)
    if dialog.result:
        excel_path = excel_path_var.get()
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", "Cannot add entry: Excel file not found.")
            return
        if backup_enabled_var.get(): create_backup(excel_path, backup_path_var.get())
        try:
            df = pd.read_excel(excel_path)
            entry = dialog.result
            date_obj = datetime.datetime.strptime(entry['date'], '%Y-%m-%d')
            new_row = {"שנה": date_obj.year, "חודש": date_obj.month, "יום": date_obj.day, "זמן התחלה": entry['start_time'], "זמן סיום": entry['end_time'], "שעות": calculate_hours_from_strings(entry['start_time'], entry['end_time']), "מה": entry['notes']}
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            df = df.sort_values(by=['שנה', 'חודש', 'יום', 'זמן התחלה'])
            df.to_excel(excel_path, index=False)
            output_text.insert(END, f"Successfully added manual entry for {entry['date']}.\n", "success")
            update_total_hours_display()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save manual entry: {e}")

def clear_sheet():
    if not messagebox.askyesno("Confirm Clear", "Are you sure you want to delete ALL entries from the Excel file?\nThis cannot be undone, but a backup will be created if enabled."):
        return
    excel_path = excel_path_var.get()
    if not os.path.exists(excel_path):
        messagebox.showerror("Error", "Cannot clear: Excel file not found.")
        return
    if backup_enabled_var.get(): create_backup(excel_path, backup_path_var.get())
    try:
        headers = ["שנה", "חודש", "יום", "זמן התחלה", "זמן סיום", "שעות", "מה"]
        df = pd.DataFrame(columns=headers)
        df.to_excel(excel_path, index=False)
        output_text.insert(END, "Excel sheet has been cleared.\n", "success")
        update_total_hours_display()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to clear sheet: {e}")

def check_excel_overlaps():
    check_overlap_button.config(state=DISABLED)
    output_text.insert(END, "\n--- Checking Excel for Overlapping Hours ---\n", "info")
    thread = threading.Thread(target=execute_excel_overlap_check_in_thread)
    thread.start()

def execute_excel_overlap_check_in_thread():
    try:
        excel_path = excel_path_var.get()
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"Excel file not found at:\n{excel_path}")
            return
        if backup_enabled_var.get(): create_backup(excel_path, backup_path_var.get())
        df = pd.read_excel(excel_path, dtype={'זמן התחלה': str, 'זמן סיום': str})
        df.columns = ["שנה", "חודש", "יום", "זמן התחלה", "זמן סיום", "שעות", "מה"]
        df['start_dt'] = pd.to_datetime(df['זמן התחלה'], format='%H:%M', errors='coerce').dt.time
        df['end_dt'] = pd.to_datetime(df['זמן סיום'], format='%H:%M', errors='coerce').dt.time
        invalid_rows = df[df['start_dt'].isna() | df['end_dt'].isna()]
        if not invalid_rows.empty:
            output_text.insert(END, "Warning: Some rows had invalid time formats and were ignored:\n", "danger")
            for index, row in invalid_rows.iterrows():
                msg = f"  - Row {index + 2}: Start='{row['זמן התחלה']}', End='{row['זמן סיום']}'\n"
                output_text.insert(END, msg, "danger")
        df.dropna(subset=['start_dt', 'end_dt'], inplace=True)
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
                conflicting_entries_by_day[date_key] = [{'original_index': row['index'], 'start_time': row['זמן התחלה'], 'end_time': row['זמן סיום'], 'summary': str(row['מה']) if pd.notna(row['מה']) else '', 'is_overlap': row['is_overlap']} for _, row in day_entries.iterrows()]
        
        if not conflicting_entries_by_day:
            output_text.insert(END, "Success: No overlapping hours found in Excel file.\n", "success")
            messagebox.showinfo("Validation Success", "The Excel file is valid!\nNo overlapping entries were found.", icon='info')
            return

        output_text.insert(END, f"Found overlaps on {len(conflicting_entries_by_day)} day(s). Opening resolution dialog...\n", "danger")
        dialog = OverlapResolutionDialog(root, conflicting_entries_by_day)
        if dialog.result:
            for date_key, modified_entries in dialog.result.items():
                for entry in modified_entries:
                    idx = entry['original_index']
                    df.loc[idx, 'שעות'] = calculate_hours_from_strings(entry['start_time'], entry['end_time'])
                    df.loc[idx, 'זמן התחלה'] = entry['start_time']
                    df.loc[idx, 'זמן סיום'] = entry['end_time']
                    df.loc[idx, 'מה'] = entry['summary']
            df = df.drop(columns=['start_dt', 'end_dt', 'is_overlap'], errors='ignore')
            df = df.sort_values(by=['שנה', 'חודש', 'יום', 'זמן התחלה'])
            df.to_excel(excel_path, index=False)
            output_text.insert(END, "Excel file updated and sorted successfully.\n", "success")
            update_total_hours_display()
        else:
            output_text.insert(END, "Overlap resolution cancelled. No changes were made.\n", "info")
    except Exception as e:
        output_text.insert(END, f"\nAn error occurred during Excel validation: {e}\n", "danger")
    finally:
        root.after(100, lambda: check_overlap_button.config(state=NORMAL))

def run_script():
    run_button.config(state=DISABLED)
    output_text.delete(1.0, END)
    output_text.insert(END, "--- Starting Timesheet Automation ---\n")
    output_text.insert(END, f"Using config: User={config.username}, Excel Path={excel_path_var.get()}\n")
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
    if calendar_var.get(): set_current_month()

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
    excel_path = excel_path_var.get()
    if not excel_path:
        messagebox.showerror("Error", "Excel file path cannot be empty.")
        return

    if not os.path.exists(excel_path):
        if not messagebox.askyesno("Create New File?", f"The Excel file was not found at:\n\n{excel_path}\n\nDo you want to create it?"):
            output_text.insert(END, "File creation cancelled by user.\n", "info")
            return
        # If user says yes, we just continue. The creation is handled by the backend function.
        output_text.insert(END, f"A new Excel file will be created at the specified path.\n", "info")

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
    thread = threading.Thread(target=execute_calendar_import_in_thread, args=(start_date, end_date, excel_path))
    thread.start()

def execute_calendar_import_in_thread(start_date, end_date, excel_path):
    try:
        from google_calendar_integration import get_calendar_service, get_calendar_events, update_excel_with_calendar_events
        if backup_enabled_var.get(): create_backup(excel_path, backup_path_var.get())
        service = get_calendar_service()
        events = get_calendar_events(service, start_date, end_date)
        def conflict_resolution_callback(new_event, existing_event):
            dialog = ConflictResolutionDialog(root, new_event, existing_event)
            return dialog.result
        change_log = update_excel_with_calendar_events(excel_path, events, conflict_resolution_callback)
        output_text.insert(END, "\n--- Excel Update Summary ---\n", "info")
        if change_log:
            for change in change_log: output_text.insert(END, f"- {change}\n")
        else:
            output_text.insert(END, "No changes were made to the Excel file.\n")
        output_text.insert(END, "--- Update complete ---\n", "success")
        update_total_hours_display()
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

# --- Configuration Frame ---
config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding=10)
config_frame.pack(fill=X, padx=5, pady=5)
ttk.Label(config_frame, text=f"Username: {config.username} (from config.py)").pack(anchor=W)
file_frame = ttk.Frame(config_frame)
file_frame.pack(fill=X, pady=5)
ttk.Label(file_frame, text="Excel File:").pack(side=LEFT, anchor=W)
excel_path_var = tk.StringVar(value=config.excel_file_path)
excel_path_entry = ttk.Entry(file_frame, textvariable=excel_path_var)
excel_path_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
ttk.Button(file_frame, text="Browse...", command=browse_excel_file, bootstyle="secondary").pack(side=LEFT)
backup_frame = ttk.Frame(config_frame)
backup_frame.pack(fill=X, pady=(5,0))
backup_enabled_var = tk.BooleanVar(value=True)
backup_check = ttk.Checkbutton(backup_frame, text="Create backup", variable=backup_enabled_var, command=toggle_backup_fields, bootstyle="primary")
backup_check.pack(side=LEFT)
backup_path_var = tk.StringVar()
backup_path_entry = ttk.Entry(backup_frame, textvariable=backup_path_var)
backup_path_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
backup_browse_button = ttk.Button(backup_frame, text="Browse...", command=browse_backup_folder, bootstyle="secondary")
backup_browse_button.pack(side=LEFT)

# --- Step 1: Google Calendar Frame ---
calendar_frame = ttk.LabelFrame(main_frame, text="Step 1: Get Events from Calendar", padding=10)
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

# --- Step 2: Excel Tools Frame ---
excel_tools_frame = ttk.LabelFrame(main_frame, text="Step 2: Review and Edit Excel Data", padding=10)
excel_tools_frame.pack(fill=X, padx=5, pady=5)
tools_button_frame = ttk.Frame(excel_tools_frame)
tools_button_frame.pack(fill=X)
check_overlap_button = ttk.Button(tools_button_frame, text="Validate & Fix Overlaps", command=check_excel_overlaps, bootstyle="warning")
check_overlap_button.pack(side=LEFT, fill=X, expand=True, ipady=4, padx=(0,5))
add_manual_button = ttk.Button(tools_button_frame, text="Add Manual Entry", command=add_manual_entry, bootstyle="secondary")
add_manual_button.pack(side=LEFT, fill=X, expand=True, ipady=4, padx=5)
clear_sheet_button = ttk.Button(tools_button_frame, text="Clear All Entries", command=clear_sheet, bootstyle="danger-outline")
clear_sheet_button.pack(side=LEFT, fill=X, expand=True, ipady=4, padx=(5,0))

# --- Step 3: Run Automation Frame ---
run_frame = ttk.LabelFrame(main_frame, text="Step 3: Run Automation", padding=10)
run_frame.pack(fill=X, padx=5, pady=5)
run_button = ttk.Button(run_frame, text="Run Automation on Webtime", command=run_script, bootstyle="primary")
run_button.pack(fill=X, ipady=5)

# --- Output Frame ---
out_frame = ttk.LabelFrame(main_frame, text="Output Log", padding=10)
out_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
output_text = ScrolledText(out_frame, wrap=WORD, height=8, autohide=True)
output_text.pack(fill=BOTH, expand=True, padx=5, pady=5)
output_text.tag_config("success", foreground=root.style.colors.success)
output_text.tag_config("danger", foreground=root.style.colors.danger)
output_text.tag_config("info", foreground=root.style.colors.info)

# --- Status Bar ---
status_frame = ttk.Frame(main_frame, padding=(5, 2))
status_frame.pack(fill=X, padx=5, pady=(5,0))
total_hours_var = tk.StringVar(value="Total Hours: N/A")
# --- MODIFIED: Increased font size and set bootstyle to primary for black text ---
total_hours_label = ttk.Label(status_frame, textvariable=total_hours_var, bootstyle="primary", font="-size 8 -weight bold")
total_hours_label.pack(side=RIGHT)

# --- Finalize Window ---
center_window(root, min_width=750, min_height=800)
update_backup_path_default()
toggle_backup_fields()
update_total_hours_display()
root.mainloop()