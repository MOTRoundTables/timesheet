# First, ensure you have installed the required library:
# pip install ttkbootstrap

import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap.constants import *
from tkinter import messagebox
import subprocess
import threading
import config
import datetime
from google_calendar_integration import get_calendar_service, get_calendar_events, update_excel_with_calendar_events

# --- NEW HELPER FUNCTION for resizing and centering windows ---
def center_window(window, min_width=0, min_height=0):
    """
    Calculates the required size for a window's content, enforces a minimum size,
    and centers the window on the screen.
    """
    window.update_idletasks()  # Update geometry calculations
    
    # Get required size and apply minimums
    width = max(window.winfo_reqwidth(), min_width)
    height = max(window.winfo_reqheight(), min_height)
    
    # Get screen dimensions
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    
    # Calculate position for centering
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    
    window.geometry(f'{width}x{height}+{x}+{y}')

# --- The ConflictResolutionDialog is also updated for a modern look ---
class ConflictResolutionDialog(ttk.Toplevel):
    def __init__(self, parent, new_event, existing_event):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Conflict Detected")
        # self.geometry("450x350") # <-- REMOVED fixed size

        self.new_event = new_event
        self.existing_event = existing_event
        self.result = None

        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        ttk.Label(main_frame, text=f"Conflict on {new_event['day']}/{new_event['month']}/{new_event['year']}:", font="-weight bold").pack(pady=(0, 10))

        # Existing Event Frame
        existing_frame = ttk.LabelFrame(main_frame, text="Existing Event", padding=10)
        existing_frame.pack(fill=X, pady=5, expand=True)
        ttk.Label(existing_frame, text=f"Summary: {existing_event['summary']}", wraplength=400).pack(anchor=W, padx=5)
        ttk.Label(existing_frame, text=f"Time: {existing_event['start_time']}-{existing_event['end_time']}").pack(anchor=W, padx=5)
        self.var_existing = tk.BooleanVar(value=True)
        ttk.Checkbutton(existing_frame, text="Keep Existing Event", variable=self.var_existing).pack(anchor=W, padx=5, pady=(5,0))

        # New Event Frame
        new_frame = ttk.LabelFrame(main_frame, text="New Event from Calendar", padding=10)
        new_frame.pack(fill=X, pady=5, expand=True)
        ttk.Label(new_frame, text=f"Summary: {new_event['summary']}", wraplength=400).pack(anchor=W, padx=5)
        ttk.Label(new_frame, text=f"Time: {new_event['start_time']}-{new_event['end_time']}").pack(anchor=W, padx=5)
        self.var_new = tk.BooleanVar(value=False)
        ttk.Checkbutton(new_frame, text="Keep New Event", variable=self.var_new).pack(anchor=W, padx=5, pady=(5,0))

        # Button Frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(15, 0), fill=X)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, bootstyle="primary").pack(side=RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, bootstyle="secondary").pack(side=RIGHT)

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        # --- FIX: Automatically resize and center the dialog ---
        center_window(self, min_width=480, min_height=350)
        
        self.wait_window(self)

    def on_ok(self):
        selected_events = []
        if self.var_existing.get():
            selected_events.append('existing')
        if self.var_new.get():
            selected_events.append('new')
        self.result = tuple(selected_events)
        self.destroy()

    def on_cancel(self):
        self.result = ()
        self.destroy()

def run_script():
    run_button.config(state=DISABLED)
    output_text.delete(1.0, END)
    output_text.insert(END, "--- Starting Timesheet Automation ---\n")
    output_text.insert(END, f"Using config: User={config.username}, Excel Path={config.excel_file_path}\n")
    thread = threading.Thread(target=execute_script_in_thread)
    thread.start()

def execute_script_in_thread():
    try:
        command = [
            "python",
            "C:\\Users\\Golan-New_PC\\timesheet\\timesheet_filler.py"
        ]
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace',
            creationflags=subprocess.CREATE_NO_WINDOW
        )
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

# --- Main Application Window ---
# --- FIX: Changed theme to "litera" for better visibility ---
# Other good light themes: "flatly", "sandstone", "lumen"
root = ttk.Window(themename="litera")
root.title("Timesheet Automation")
# root.geometry("750x700") # <-- REMOVED fixed size

main_frame = ttk.Frame(root, padding=15)
main_frame.pack(fill=BOTH, expand=True)

# --- Configuration Frame ---
input_frame = ttk.LabelFrame(main_frame, text="Configuration", padding=10)
input_frame.pack(fill=X, padx=5, pady=5)
info_text = f"Username: {config.username}\nPassword: ***\nExcel File: {config.excel_file_path}"
ttk.Label(input_frame, text="Credentials and file path are loaded from config.py:", justify=LEFT).pack(pady=(5,0), padx=5, anchor="w")
ttk.Label(input_frame, text=info_text, justify=LEFT, bootstyle="secondary").pack(pady=(0,5), padx=15, anchor="w")

# --- Google Calendar Integration Frame ---
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

# --- Main Action Button ---
run_button = ttk.Button(main_frame, text="Run Automation", command=run_script, bootstyle="primary")
run_button.pack(pady=15, fill=X, ipady=5)

# --- Output Frame ---
out_frame = ttk.LabelFrame(main_frame, text="Output", padding=10)
out_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
output_text = ScrolledText(out_frame, wrap=WORD, height=15, autohide=True)
output_text.pack(fill=BOTH, expand=True, padx=5, pady=5)

output_text.tag_config("success", foreground=root.style.colors.success)
output_text.tag_config("danger", foreground=root.style.colors.danger)
output_text.tag_config("info", foreground=root.style.colors.info)

# --- FIX: Automatically resize and center the main window before showing it ---
center_window(root, min_width=750, min_height=700)

root.mainloop()