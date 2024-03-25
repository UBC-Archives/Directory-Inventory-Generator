#Directory Inventory Generator
#Version: 1.1
#https://recordsmanagement.ubc.ca
#https://www.gnu.org/licenses/gpl-3.0.en.html

import os
import time
from datetime import datetime
import csv
import tkinter as tk
from tkinter import filedialog
import threading
import subprocess

rows_written = 0

def list_files_and_folders(path, csv_writer, max_level, current_level=0):
    rows = []
    batch_size = 1000
    global rows_written
    global progress_var

    try:    
        with os.scandir(path) as entries:
            for entry in entries:
                try:
                    full_path = entry.path
                    name, ext = os.path.splitext(entry.name)
                    size_bytes = entry.stat().st_size
                    created_time = datetime.fromtimestamp(entry.stat().st_birthtime).strftime('%Y-%m-%d %H:%M:%S')
                    last_modified = datetime.fromtimestamp(entry.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                    last_accessed = datetime.fromtimestamp(entry.stat().st_atime).strftime('%Y-%m-%d %H:%M:%S')
                    path_length = len(full_path)

                    path_components = full_path.split(os.path.sep)  

                    if path_components[0].count('/') == 1:
                        splitted_components = path_components[0].split('/')
                        path_components[0] = splitted_components[0]
                        path_components.insert(1, splitted_components[1])

                    file_type = 'File' if entry.is_file() else 'Folder'

                    if entry.is_file():
                        rows.append([full_path.replace('/', '\\'), name, ext, file_type, path_length, current_level, size_bytes, created_time, last_modified, last_accessed, *path_components[1:]])
                    else:
                        rows.append([full_path.replace('/', '\\'), name, '-', file_type, path_length, current_level, '-', created_time, last_modified, last_accessed, *path_components[1:]])
                        if current_level < max_level:
                            list_files_and_folders(full_path, csv_writer, max_level, current_level + 1)

                    if len(rows) >= batch_size:
                        csv_writer.writerows(rows)
                        rows_written += len(rows)
                        rows.clear()

                        if progress_var:
                            progress_var.set(str(rows_written))

                except PermissionError as pe:
                    log_error(f"PermissionError: {pe} (Access denied for '{entry.path}')")
                    continue
                except Exception as e:
                    log_error(f"Unexpected error for '{entry.path}': {e}")
                    continue
    except PermissionError as pe:
        log_error(f"PermissionError: {pe} (Access denied for '{path}')")
        return
    except Exception as e:
        log_error(f"Unexpected error for '{path}': {e}")
        return

    if rows:
        csv_writer.writerows(rows)
        rows_written += len(rows)
        rows.clear()

        if progress_var:
            progress_var.set(str(rows_written))

def log_error(error_message):
    global timestamp
    with open(f'Error-Log_{timestamp}.txt', 'a', encoding='utf-8') as log_file:
        log_file.write(f"{error_message}\n")

def show_completion_message(execution_time_str, output_file, error_log_path):
    completion_window = tk.Toplevel(root)
    completion_window.title("Inventory Generation Completed!")

    message = f"Execution completed in {execution_time_str}.\n\nOutput file saved as {output_file}"

    completion_label = tk.Label(completion_window, text=message)
    completion_label.pack(padx=20, pady=5)

    open_button = tk.Button(completion_window, text="Open Output File", command=lambda: open_output_file(output_file))
    open_button.pack(pady=5)

    if os.path.exists(error_log_path):
        error_log_label = tk.Label(completion_window, text=f"Error log saved as {error_log_path}")
        error_log_label.pack(pady=5)

        open_log_button = tk.Button(completion_window, text="Open Error Log", command=lambda: open_output_file(error_log_path))
        open_log_button.pack(pady=5)

    # Clear all fields of the main window
    clear_fields()

def open_output_file(output_file):
    try:
        subprocess.run(['start', '', output_file], shell=True)
    except Exception as e:
        tk.messagebox.showerror("Error", f"Error opening file: {str(e)}")

def list_files_in_thread():
    global thread
    global progress_var
    global timestamp
    global rows_written
    rows_written = 0
    
    path_to_list = input_path_entry.get()
    if not os.path.exists(path_to_list):
        output_label.config(fg="red")
        output_text.set("Invalid path!")
        return

    output_text.set("")

    # Directory level input
    max_level_str = directory_level_entry.get()
    if max_level_str:
        try:
            max_level = int(max_level_str)
        except ValueError:
            output_label.config(fg="red")
            output_text.set("Invalid directory level!")
            return
    else:
        # If directory level field empty
        max_level = float('inf')

    start_time = time.time()
    timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
    
    output_file = os.path.join(os.getcwd(), f'Inventory_{timestamp}.csv')
    error_log_path = os.path.join(os.getcwd(), f'Error-Log_{timestamp}.txt')

    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(['Full Path', 'Name', 'Extension', 'Type', 'Path Length (Chars)', 'Directory Level', 'Size (Bytes)', 'Created', 'Last Modified', 'Last Accessed', 'Path Components'])
        list_files_and_folders(path_to_list, csv_writer, max_level)

    # Sort the CSV file based on the first column (full path)
    sort_csv_file(output_file)

    end_time = time.time()
    execution_time = end_time - start_time
    execution_time_str = time.strftime("%H:%M:%S", time.gmtime(execution_time))

    show_completion_message(execution_time_str, output_file, error_log_path)

thread = threading.Thread(target=list_files_in_thread, daemon=True)

def sort_csv_file(csv_file):
    try:
        with open(csv_file, 'r', newline='', encoding='utf-8') as csvfile:
            rows = list(csv.reader(csvfile))
        
        header = rows[0]
        data_rows = rows[1:]
        
        # Sort only the data rows based on the first column (full path)
        data_rows.sort(key=lambda x: x[0])
        
        # Combine the sorted data rows with the original header
        sorted_rows = [header] + data_rows
        
        with open(csv_file, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerows(sorted_rows)
    except Exception as e:
        tk.messagebox.showerror("Error", f"Error sorting file: {str(e)}")

def execute_button_callback():
    global thread

    if thread.is_alive():
        output_text.set("A process is already running.")
    else:
        input_path_entry.config(state=tk.DISABLED)
        execute_button.config(state=tk.DISABLED)
        browse_button.config(state=tk.DISABLED)
        directory_level_entry.config(state=tk.DISABLED)

        thread = threading.Thread(target=list_files_in_thread, daemon=True)
        thread.start()

        root.after(100, check_thread_status)

def check_thread_status():
    if thread.is_alive():
        root.after(100, check_thread_status)
    else:
        input_path_entry.config(state=tk.NORMAL)
        execute_button.config(state=tk.NORMAL)
        browse_button.config(state=tk.NORMAL)
        directory_level_entry.config(state=tk.NORMAL)

def browse_button_callback():
    selected_path = filedialog.askdirectory()
    if selected_path:
        input_path_entry.delete(0, tk.END)
        input_path_entry.insert(0, selected_path)
        output_label.config(fg="black")
        output_text.set("")

# File menu functions
def exit_app():
    if threading.active_count() > 1:
        confirm = tk.messagebox.askyesno("Exit?", "A process is running. Are you sure you want to exit?")
        if not confirm:
            return
    root.destroy()

def clear_fields():
    # Enable the Path entry
    input_path_entry.config(state=tk.NORMAL)

    # Enable the Max Directory Depth entry
    directory_level_entry.config(state=tk.NORMAL)

    # Clear the fields
    input_path_entry.delete(0, tk.END)
    directory_level_entry.delete(0, tk.END)
    output_text.set("")
    progress_var.set("0")

def show_help():
    help_message = "This program lists files and folders in a directory.\n\n" \
                   "Use 'Browse' to select the path, set the maximum directory depth, and click Generate Inventory.\n\n" \
                   "The output is saved in a CSV file."
    tk.messagebox.showinfo("Help", help_message)

def show_about():
    about_message = "Directory Inventory Generator\nVersion 1.1"
   
    about_window = tk.Toplevel(root)
    about_window.title("About")
    about_window.resizable(False, False)

    about_label = tk.Label(about_window, text=about_message)
    about_label.pack(padx=20, pady=10)

    # Frame for UBC RMO
    ubc_frame = tk.Frame(about_window)
    ubc_frame.pack(padx=20, pady=20)

    ubc_label = tk.Label(ubc_frame, text="Developed by\nRecords Management Office\nThe University of British Columbia")
    ubc_label.pack()

    ubc_link_label = tk.Label(ubc_frame, text="https://recordsmanagement.ubc.ca", fg="blue", cursor="hand2")
    ubc_link_label.pack()

    def open_ubc_link(event):
        import webbrowser
        webbrowser.open("https://recordsmanagement.ubc.ca")

    ubc_link_label.bind("<Button-1>", open_ubc_link)

    # Frame for license
    license_frame = tk.Frame(about_window)
    license_frame.pack(padx=20, pady=10)

    license_label = tk.Label(license_frame, text="License: ")
    license_label.pack(side='left')

    license_link_label = tk.Label(license_frame, text="GPL-3.0", fg="blue", cursor="hand2")
    license_link_label.pack(side='left')

    def open_license_link(event):
        import webbrowser
        webbrowser.open("https://www.gnu.org/licenses/gpl-3.0.en.html")

    license_link_label.bind("<Button-1>", open_license_link)

# Main window
root = tk.Tk()
root.title("Directory Inventory Generator")
root.geometry("600x150")
root.resizable(False, False)

# Menu bar
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# File menu
file_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)

# File menu options
file_menu.add_command(label="Clear Fields", command=clear_fields)
file_menu.add_command(label="Exit", command=exit_app)

# Help menu
help_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Help", menu=help_menu)

# Help menu options
help_menu.add_command(label="Help", command=show_help)
help_menu.add_command(label="About", command=show_about)

# Main window padding
root.option_add('*TButton*Padding', 5)
root.option_add('*TButton*highlightThickness', 0)
root.option_add('*TButton*highlightColor', 'SystemButtonFace')

# Frame to hold the path input field and the "Browse" button
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Input field for the path
input_label = tk.Label(frame, text="Directory Path:")
input_label.pack(side='left')

input_path_entry = tk.Entry(frame, width=60)
input_path_entry.pack(side='left', padx=10)

# "Browse" button
browse_button = tk.Button(frame, text="Browse", command=browse_button_callback)
browse_button.pack(side='left', padx=10)

# Frame to hold the directory level input field
frameD = tk.Frame(root)
frameD.pack(padx=10, pady=10)

# Label for the directory level input field
directory_level_label = tk.Label(frameD, text="Max Directory Depth:")
directory_level_label.pack(side='left')

# Entry field for the directory level
directory_level_entry = tk.Entry(frameD, width=5)
directory_level_entry.pack(side='left', padx=10)

# Hint label for directory level field
hint_label_text = " (Root = 0; Leave empty for deepest level)"
hint_label = tk.Label(frameD, text=hint_label_text, font=("Arial", 8), fg="gray")
hint_label.pack(side='left')

# Frame to hold the "Execute" button
frameE = tk.Frame(root)
frameE.pack(padx=10, pady=10)

# "Execute" button
execute_button = tk.Button(frameE, text="Generate Inventory", command=execute_button_callback)
execute_button.pack(side='left', padx=10)

# Output label to display invalid path messages
output_text = tk.StringVar()
output_label = tk.Label(frameE, textvariable=output_text)
output_label.pack(side='left', padx=10)

# Progress label
progress_label = tk.Label(frameE, text="Processed files & folders:")
progress_label.pack(side='left', padx=10)

# Shared variable to store the number of rows written
progress_var = tk.StringVar()
progress_var.set("0")

# Label to display the number of rows written
rows_written_label = tk.Label(frameE, textvariable=progress_var)
rows_written_label.pack(side='left')

root.protocol("WM_DELETE_WINDOW", exit_app)

root.mainloop()