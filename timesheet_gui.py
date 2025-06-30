import tkinter as tk
from tkinter import scrolledtext, messagebox
import subprocess
import threading
import config
import datetime
from google_calendar_integration import get_calendar_service, get_calendar_events, update_excel_with_calendar_events

class ConflictResolutionDialog(tk.Toplevel):
    def __init__(self, parent, new_event, existing_event):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Conflict Detected")

        self.new_event = new_event
        self.existing_event = existing_event
        self.result = None  # Will store ('new', 'existing'), ('new',), ('existing',), or ()

        tk.Label(self, text=f"Conflict on {new_event['day']}/{new_event['month']}/{new_event['year']}:").pack(pady=5)

        # Existing Event
        # Existing Event
        tk.Label(self, text="Existing Event:", font=('TkDefaultFont', 10, 'bold')).pack(anchor='w', padx=10)
        tk.Label(self, text=f"  Summary: {existing_event['summary']}").pack(anchor='w', padx=20)
        tk.Label(self, text=f"  Time: {existing_event['start_time']}-{existing_event['end_time']}").pack(anchor='w', padx=20)
        self.var_existing = tk.BooleanVar(value=True) # Default to keeping existing
        tk.Checkbutton(self, text="Keep Existing Event", var=self.var_existing).pack(anchor='w', padx=10)

        # New Event
        tk.Label(self, text="\nNew Event:", font=('TkDefaultFont', 10, 'bold')).pack(anchor='w', padx=10)
        tk.Label(self, text=f"  Summary: {new_event['summary']}").pack(anchor='w', padx=20)
        tk.Label(self, text=f"  Time: {new_event['start_time']}-{new_event['end_time']}").pack(anchor='w', padx=20)
        self.var_new = tk.BooleanVar(value=False) # Default to not keeping new
        tk.Checkbutton(self, text="Keep New Event", var=self.var_new).pack(anchor='w', padx=10)

        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="OK", command=self.on_ok).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Cancel", command=self.on_cancel).pack(side=tk.LEFT, padx=5)

        self.protocol("WM_DELETE_WINDOW", self.on_cancel) # Handle window close button
        self.parent = parent
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
        self.result = () # Return empty tuple if cancelled (keep none)
        self.destroy()

def run_script():
    run_button.config(state=tk.DISABLED)
    output_text.delete(1.0, tk.END)
    output_text.insert(tk.END, "--- Starting Timesheet Automation ---\n")
    output_text.insert(tk.END, f"Using config: User={config.username}, Excel Path={config.excel_file_path}\n")
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
            output_text.insert(tk.END, line)
            output_text.see(tk.END)
        process.stdout.close()
        return_code = process.wait()
        if return_code == 0:
            output_text.insert(tk.END, "\n--- Script finished successfully! ---\n")
        else:
            output_text.insert(tk.END, f"\n--- Script finished with error code: {return_code} ---\n")
    except FileNotFoundError:
        output_text.insert(tk.END, "\nError: 'python' command not found. Make sure Python is installed and in your system's PATH.\n")
    except Exception as e:
        output_text.insert(tk.END, f"\nAn unexpected error occurred: {e}\n")
    finally:
        root.after(100, lambda: run_button.config(state=tk.NORMAL))

def toggle_calendar_fields():
    state = tk.NORMAL if calendar_var.get() else tk.DISABLED
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
    start_date = datetime.datetime.strptime(start_date_var.get(), "%Y-%m-%d").date()
    new_start_of_month = (start_date.replace(day=1) + datetime.timedelta(days=delta * 32)).replace(day=1)
    new_end_of_month = (new_start_of_month + datetime.timedelta(days=32)).replace(day=1) - datetime.timedelta(days=1)
    start_date_var.set(new_start_of_month.strftime("%Y-%m-%d"))
    end_date_var.set(new_end_of_month.strftime("%Y-%m-%d"))

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

    import_calendar_button.config(state=tk.DISABLED)
    output_text.insert(tk.END, "\n--- Updating Excel from Google Calendar ---\n")
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
        output_text.insert(tk.END, "\n--- Excel Update Summary ---\n")
        if change_log:
            for change in change_log:
                output_text.insert(tk.END, f"- {change}\n")
        else:
            output_text.insert(tk.END, "No changes were made to the Excel file.\n")
        output_text.insert(tk.END, "--- Update complete ---\n")

    except Exception as e:
        output_text.insert(tk.END, f"\nAn error occurred during calendar import: {e}\n")
    finally:
        root.after(100, lambda: import_calendar_button.config(state=tk.NORMAL))

root = tk.Tk()
root.title("Timesheet Automation")
main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(fill=tk.BOTH, expand=True)
input_frame = tk.LabelFrame(main_frame, text="Configuration")
input_frame.pack(fill=tk.X, padx=5, pady=5)
info_text = f"Username: {config.username}\nPassword: ***\nExcel File: {config.excel_file_path}"
tk.Label(input_frame, text="Credentials and file path are loaded from config.py:", justify=tk.LEFT).pack(pady=(5,0), padx=5, anchor="w")
tk.Label(input_frame, text=info_text, justify=tk.LEFT, fg="grey").pack(pady=(0,5), padx=15, anchor="w")

# Google Calendar Integration Frame
calendar_frame = tk.LabelFrame(main_frame, text="Google Calendar Integration")
calendar_frame.pack(fill=tk.X, padx=5, pady=5)
calendar_var = tk.BooleanVar()
calendar_check = tk.Checkbutton(calendar_frame, text="Enable Google Calendar", var=calendar_var, command=toggle_calendar_fields)
calendar_check.pack(anchor="w", padx=5, pady=5)

date_frame = tk.Frame(calendar_frame)
date_frame.pack(fill=tk.X, padx=5)

start_date_var = tk.StringVar()
end_date_var = tk.StringVar()

prev_month_button = tk.Button(date_frame, text="<", command=lambda: change_month(-1), state=tk.DISABLED)
prev_month_button.pack(side=tk.LEFT)

start_date_entry = tk.Entry(date_frame, textvariable=start_date_var, state=tk.DISABLED, width=12)
start_date_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
end_date_entry = tk.Entry(date_frame, textvariable=end_date_var, state=tk.DISABLED, width=12)
end_date_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

next_month_button = tk.Button(date_frame, text=">", command=lambda: change_month(1), state=tk.DISABLED)
next_month_button.pack(side=tk.LEFT)

import_calendar_button = tk.Button(calendar_frame, text="Update Excel from Google Calendar", command=import_calendar_data, state=tk.DISABLED)
import_calendar_button.pack(pady=5)


run_button = tk.Button(main_frame, text="Run Automation", command=run_script, bg="lightblue", font=('Helvetica', 10, 'bold'))
run_button.pack(pady=10)
out_frame = tk.LabelFrame(main_frame, text="Output")
out_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
output_text = scrolledtext.ScrolledText(out_frame, wrap=tk.WORD, height=15, width=80)
output_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
root.mainloop()
