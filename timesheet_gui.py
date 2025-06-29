import tkinter as tk
from tkinter import scrolledtext
import subprocess
import threading
import config

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

root = tk.Tk()
root.title("Timesheet Automation")
main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(fill=tk.BOTH, expand=True)
input_frame = tk.LabelFrame(main_frame, text="Configuration")
input_frame.pack(fill=tk.X, padx=5, pady=5)
info_text = f"Username: {config.username}\nPassword: ***\nExcel File: {config.excel_file_path}"
tk.Label(input_frame, text="Credentials and file path are loaded from config.py:", justify=tk.LEFT).pack(pady=(5,0), padx=5, anchor="w")
tk.Label(input_frame, text=info_text, justify=tk.LEFT, fg="grey").pack(pady=(0,5), padx=15, anchor="w")
run_button = tk.Button(main_frame, text="Run Automation", command=run_script, bg="lightblue", font=('Helvetica', 10, 'bold'))
run_button.pack(pady=10)
output_frame = tk.LabelFrame(main_frame, text="Output")
output_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, height=15, width=80)
output_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
root.mainloop()
