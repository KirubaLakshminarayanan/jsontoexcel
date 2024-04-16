import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Button
import json
from openpyxl import Workbook
from datetime import datetime
import logging
import os

# Create a logs directory if it doesn't exist
logs_dir = 'logs'
if not os.path.exists(logs_dir):
    os.makedirs(logs_dir)


# Function to configure logging with a custom log file name
def configure_logging(log_filename):
    logging.basicConfig(filename=log_filename, level=logging.DEBUG,
                        format='%(asctime)s - %(levelname)s - %(message)s')


# Function to check if a file is a JSON file
def is_json(file_path):
    try:
        with open(file_path) as f:
            json.load(f)
            return True
    except (ValueError, FileNotFoundError):
        return False


# Function to flatten JSON data
def flatten_json(data, parent_key='', sep='_'):
    if isinstance(data, dict):
        items = {}
        for k, v in data.items():
            new_key = parent_key + sep + k if parent_key else k
            items.update(flatten_json(v, new_key, sep=sep))
        return items
    elif isinstance(data, list):
        items = {}
        for i, v in enumerate(data):
            new_key = parent_key + sep + str(i)
            items.update(flatten_json(v, new_key, sep=sep))
        return items
    else:
        return {parent_key: data}


# Function to convert JSON data to Excel
def convert_to_excel(file_paths, progress_bar, status_label, info_label, completion_label):
    log_filename = os.path.join(logs_dir, f"main_errors_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    configure_logging(log_filename)  # Configure logging with a new log file name
    total_files = len(file_paths)
    progress = 0
    success_count = 0  # Track the number of successful conversions

    for file_path in file_paths:
        if not file_path.endswith('.json'):
            logging.error(f"Skipping {file_path}: Invalid file format. Please provide a JSON file.")
            messagebox.showerror("Error", f"Conversion of JSON to Excel file is failed because of invalid file format.")
            continue

        if os.path.getsize(file_path) == 0:
            logging.error(f"Skipping {file_path}: File is empty.")
            messagebox.showerror("Error", f"{file_path} is empty. Please provide a valid JSON file.")
            continue

        try:
            with open(file_path, 'r') as json_file:
                try:
                    data = json.load(json_file)
                except ValueError as e:
                    error_message = f"Invalid JSON format in {file_path}: {str(e)}"
                    logging.error(error_message)
                    messagebox.showerror("Error", error_message)
                    continue

                if isinstance(data, dict):
                    records = [data]
                elif isinstance(data, list):
                    records = data
                else:
                    logging.error(f"Skipping {file_path}: Invalid JSON data format.")
                    continue

                if not records:
                    logging.error(f"No records found in {file_path}.")
                    continue

                flattened_records = [flatten_json(record) for record in records]

                # Get the headers in the order of keys from the first record
                headers = list(flattened_records[0].keys())

                excel_file_name = file_path[:-5] + '_' + datetime.now().strftime("%Y%m%d%H%M%S") + '.xlsx'
                wb = Workbook()
                ws = wb.active

                ws.append(headers)

                for record in flattened_records:
                    row_data = []
                    for header in headers:
                        row_data.append(
                            record.get(header, ''))  # Get value for each header or empty string if not present
                    ws.append(row_data)

                wb.save(excel_file_name)
                logging.info(f"{file_path} conversion successful. Converted to {excel_file_name}.")
                logging.info(f"Size of converted file: {os.path.getsize(excel_file_name)} bytes")
                success_count += 1  # Increment the success count
        except Exception as e:
            logging.error(f"{file_path} conversion failed! Error: {str(e)}")
            messagebox.showinfo("Conversion Failed",
                                f"Conversion of {file_path} failed. Please refer to the log files for more information.")

        progress += 1
        progress_bar['value'] = (progress / total_files) * 100
        progress_bar.update_idletasks()
        completion_label.config(text=f"{int(progress / total_files * 100)}% completed")

    failed_files_count = total_files - success_count
    if failed_files_count == 0:
        status_label.config(
            text=f"Completed {progress}/{total_files} files. Conversion of JSON to Excel file is successful.")
        info_label.config(text="")
    else:
        status_label.config(
            text=f"Completed {progress}/{total_files} files. Conversion of JSON to Excel file is failed for {failed_files_count} file(s).")
        info_label.config(text="Please refer to the log files for more information.")


# Function to browse for JSON files
def browse_files(file_entry, convert_to_excel_button):
    file_paths = filedialog.askopenfilenames(filetypes=[("JSON Files", "*.json")])
    if file_paths:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, ", ".join(file_paths))
        convert_to_excel_button.config(state=tk.NORMAL)


# Function to clear selected files
def clear_files(file_entry, convert_to_excel_button, progress_bar, status_label, info_label, completion_label):
    files = file_entry.get()
    if not files:
        status_label.config(text="")
        messagebox.showinfo("No Files Selected", "There are no files selected for clear.")
    else:
        file_entry.delete(0, tk.END)
        status_label.config(text="")
        completion_label.config(text="")
        info_label.config(text="")
        convert_to_excel_button.config(state=tk.DISABLED)
        progress_bar['value'] = 0
        progress_bar.update_idletasks()
        progress_bar.pack_forget()  # Hide the progress bar
        messagebox.showinfo("Files Cleared", "Selected files are cleared.")


# Function to handle conversion process
def handle_conversion(file_entry, progress_bar, status_label, info_label, completion_label):
    file_paths = file_entry.get().split(", ")
    progress_bar.pack()  # Show the progress bar
    convert_to_excel(file_paths, progress_bar, status_label, info_label, completion_label)


# Function to confirm exiting the application
def confirm_exit(root):
    confirmation = messagebox.askquestion("Confirm Exit", "Do you want to exit from the application?")
    if confirmation == 'yes':
        root.destroy()


class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tip_window, text=self.text, justify="left", background="#ffffe0", relief="solid",
                         borderwidth=1)
        label.pack()

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()


root = tk.Tk()
root.title("JSON to Excel Converter")
root.geometry("600x400")  # Increase window size

file_label = tk.Label(root, text="Select JSON Files:")
file_label.pack()

file_entry = tk.Entry(root, width=70)
file_entry.pack()
ToolTip(file_entry, "Select JSON file(s) for conversion")  # Add tooltip to file entry

browse_button = tk.Button(root, text="Browse", command=lambda: browse_files(file_entry, convert_to_excel_button))
browse_button.pack()

clear_button = tk.Button(root, text="Clear",
                         command=lambda: clear_files(file_entry, convert_to_excel_button, progress_bar, status_label,
                                                     info_label, completion_label))
clear_button.pack()

convert_to_excel_button = tk.Button(root, text="Convert to Excel", state=tk.DISABLED,
                                    command=lambda: handle_conversion(file_entry, progress_bar, status_label,
                                                                      info_label, completion_label))
convert_to_excel_button.pack()

exit_button = tk.Button(root, text="Exit", command=lambda: confirm_exit(root))
exit_button.pack()

progress_bar = Progressbar(root, orient=tk.HORIZONTAL, length=200, mode='determinate')

status_label = tk.Label(root, text="")
status_label.pack()

info_label = tk.Label(root, text="", fg="blue")
info_label.pack()

completion_label = tk.Label(root, text="", fg="green")
completion_label.pack()

root.mainloop()
