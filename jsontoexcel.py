import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Button, Style
from openpyxl import Workbook
import json
import os
from datetime import datetime
import logging

# Create a logs directory if it doesn't exist
logs_dir = 'logs'
if not os.path.exists(logs_dir):
    os.makedirs(logs_dir)

# Set up logging
log_filename = os.path.join(logs_dir, f"main_errors_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
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
            if isinstance(v, dict):
                items.update(flatten_json(v, new_key, sep=sep))
            elif isinstance(v, list):
                for i, item in enumerate(v):
                    items.update(flatten_json(item, new_key + sep + str(i), sep=sep))
            else:
                items[new_key] = v
        return items
    else:
        return {parent_key: data}

# Function to convert JSON data to Excel
def convert_to_excel(file_paths, progress_bar, status_label):
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

                flattened_data = flatten_json(data)

                headers = sorted(list(flattened_data.keys()))

                excel_file_name = file_path[:-5] + '_' + datetime.now().strftime("%Y%m%d%H%M%S") + '.xlsx'
                wb = Workbook()
                ws = wb.active

                ws.append(headers)

                row_data = [flattened_data.get(header, "") for header in headers]
                ws.append(row_data)

                wb.save(excel_file_name)
                logging.info(f"{file_path} conversion successful. Converted to {excel_file_name}.")
                logging.info(f"Size of converted file: {os.path.getsize(excel_file_name)} bytes")
                success_count += 1  # Increment the success count
        except Exception as e:
            logging.error(f"{file_path} conversion failed! Error: {str(e)}")

        progress += 1
        progress_bar['value'] = (progress / total_files) * 100
        progress_bar.update_idletasks()

    if success_count == total_files:
        status_label.config(
            text=f"Completed {progress}/{total_files} files. Conversion of JSON to Excel file is successful.")
    else:
        status_label.config(
            text=f"Completed {progress}/{total_files} files. Conversion of JSON to Excel file is failed.")

# Function to browse for JSON files
def browse_files(file_entry, convert_to_excel_button):
    file_paths = filedialog.askopenfilenames(filetypes=[("JSON Files", "*.json")])
    if file_paths:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, ", ".join(file_paths))
        convert_to_excel_button.config(state=tk.NORMAL)

# Function to clear selected files
def clear_files(file_entry, convert_to_excel_button, progress_bar, status_label):
    files = file_entry.get()
    if not files:
        status_label.config(text="")
        messagebox.showinfo("No Files Selected", "There are no files selected for clear.")
    else:
        file_entry.delete(0, tk.END)
        status_label.config(text="")
        convert_to_excel_button.config(state=tk.DISABLED)
        progress_bar.pack_forget()  # Hide the progress bar
        messagebox.showinfo("Files Cleared", "Selected files are cleared.")

# Function to handle conversion process
def handle_conversion(file_entry, progress_bar, status_label):
    file_paths = file_entry.get().split(", ")
    progress_bar.pack()  # Show the progress bar
    convert_to_excel(file_paths, progress_bar, status_label)

# Function to confirm exiting the application
def confirm_exit(root):
    confirmation = messagebox.askquestion("Confirm Exit", "Do you want to exit from the application?")
    if confirmation == 'yes':
        root.destroy()

root = tk.Tk()
root.title("JSON to Excel Converter")
root.geometry("500x300")

style = Style()
style.configure('TButton', font=('calibri', 10, 'bold'), borderwidth='4')

file_label = tk.Label(root, text="Select JSON Files:", font=('calibri', 12, 'bold'))
file_label.pack()

file_entry = tk.Entry(root, width=50, font=('calibri', 10))
file_entry.pack()

browse_button = Button(root, text="Browse", command=lambda: browse_files(file_entry, convert_to_excel_button))
browse_button.pack()

clear_button = Button(root, text="Clear", command=lambda: clear_files(file_entry, convert_to_excel_button, progress_bar, status_label))
clear_button.pack()

progress_bar = Progressbar(root, orient=tk.HORIZONTAL, length=200, mode='determinate')

convert_to_excel_button = Button(root, text="Convert to Excel", state=tk.DISABLED,
                                    command=lambda: handle_conversion(file_entry, progress_bar, status_label))
convert_to_excel_button.pack()

exit_button = Button(root, text="Exit", command=lambda: confirm_exit(root))
exit_button.pack()

status_label = tk.Label(root, text="", font=('calibri', 12))
status_label.pack()

root.mainloop()
