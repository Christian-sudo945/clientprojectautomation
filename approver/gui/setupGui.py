import os
import tkinter as tk
from tkinter import messagebox, filedialog
from gui.pyGui import run_gui
from Objects.BrowserAutomation import EdgeAutomation

'''
This file is for the setup gui, where the users would have to input the file path of their edgebrowser.
This also calls the run_gui function for the main gui (pyGui.py).
'''

LINK_FILE = "link.txt"  # File to store the path of the selected file

def load_link_from_file():
    ''' Load the file path from link.txt if it exists '''
    if os.path.exists(LINK_FILE):
        with open(LINK_FILE, "r") as file:
            link = file.read().strip()
            if os.path.isfile(link):
                return link
    return None

def submit_form(entry_filepath):
    filepath = entry_filepath.get()
    return filepath

def cancel_action():
    root.destroy() 

def select_file(entry_filepath):
    file_path = filedialog.askopenfilename()
    if file_path:
        entry_filepath.delete(0, tk.END)
        entry_filepath.insert(0, file_path)

def save_link_to_file(link):
    with open(LINK_FILE, "w") as file:
        file.write(link)

def run_setup():
    global root

    def on_button_click():
        filepath = submit_form(entry_filepath)
        save_link_to_file(filepath)
        message = f"Selected File: {filepath}"
        messagebox.showinfo("Form Submission", message)

        browser_automation = EdgeAutomation()
        if browser_automation.initialize_driver():
            root.destroy()
            run_gui(browser_automation)

    # Check if link.txt already has a file path
    existing_file_path = load_link_from_file()

    if existing_file_path:
        # If a file path is already saved in link.txt, skip the file selection step
        result = messagebox.askyesno("File Found", f"Previously selected file found: {existing_file_path}. Do you want to use this file?")
        if result:
            browser_automation = EdgeAutomation()
            if browser_automation.initialize_driver():
                run_gui(browser_automation)
            return  # Exit early since we're using the existing file

    # If no file is found in link.txt, or the user chose not to use it, show the file selection GUI

    # Main window
    root = tk.Tk()
    root.title("Work Inbox Web Automation")

    # File path
    file_label = tk.Label(root, text="Select File:")
    file_label.grid(row=2, column=0, padx=10, pady=5)
    entry_filepath = tk.Entry(root)
    entry_filepath.grid(row=2, column=1, padx=10, pady=5)
    file_button = tk.Button(root, text="Browse", command=lambda: select_file(entry_filepath))
    file_button.grid(row=2, column=2, padx=10, pady=5)

    # Cancel button
    cancelBtn = tk.Button(root, text="Cancel", command=cancel_action)
    cancelBtn.grid(row=3, column=0, padx=10, pady=5, sticky="ew")

    # Submit button
    submit_button = tk.Button(root, text="Submit", command=on_button_click)
    submit_button.grid(row=3, column=1, padx=5, pady=10, sticky="ew")

    # Run the application
    root.mainloop()

import tkinter as tk
import os
import sys
from tkinter import messagebox, filedialog

# Fix the import paths
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from approver.gui.pyGui import run_gui
from approver.Objects.BrowserAutomation import EdgeAutomation

def submit_form(entry_filepath):
    filepath = entry_filepath.get()
    return filepath

def cancel_action():
    root.destroy()

def select_file(entry_filepath):
    file_path = filedialog.askopenfilename()
    if file_path:
        entry_filepath.delete(0, tk.END)
        entry_filepath.insert(0, file_path)

def save_link_to_file(link):
    with open("approver_link.txt", "w") as file:
        file.write(link)

def run_setup():
    global root

    def on_button_click():
        filepath = submit_form(entry_filepath)
        save_link_to_file(filepath)
        message = f"Selected File: {filepath}"
        messagebox.showinfo("Form Submission", message)

        browser_automation = EdgeAutomation()
        if browser_automation.initialize_driver():
            root.destroy()
            run_gui(browser_automation)

    # main window
    root = tk.Tk()
    root.title("Approver Tool")
    root.geometry("600x400")
    root.configure(bg="#f5f7fa")
    
    # Create frames to separate pack and grid
    header_frame = tk.Frame(root, bg="#f5f7fa")
    header_frame.pack(fill="x", pady=10)
    
    content_frame = tk.Frame(root, bg="#f5f7fa")
    content_frame.pack(fill="both", expand=True, padx=20, pady=10)
    
    # Header in its own frame using pack
    label = tk.Label(
        header_frame, 
        text="SAP NetWeaver Approver Tool", 
        font=("Helvetica", 18, "bold"),
        bg="#f5f7fa",
        fg="#2c3e50"
    )
    label.pack(pady=10)
    
    # Content in its own frame using grid
    file_label = tk.Label(content_frame, text="Select File:", bg="#f5f7fa")
    file_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")
    
    entry_filepath = tk.Entry(content_frame, width=40)
    entry_filepath.grid(row=0, column=1, padx=10, pady=5)
    
    file_button = tk.Button(content_frame, text="Browse", command=lambda: select_file(entry_filepath))
    file_button.grid(row=0, column=2, padx=10, pady=5)
    
    # Button frame for the buttons
    button_frame = tk.Frame(content_frame, bg="#f5f7fa")
    button_frame.grid(row=1, column=0, columnspan=3, pady=20)
    
    # Buttons in the button frame
    cancelBtn = tk.Button(button_frame, text="Cancel", command=cancel_action, width=10)
    cancelBtn.pack(side="left", padx=10)
    
    submit_button = tk.Button(button_frame, text="Submit", command=on_button_click, width=10)
    submit_button.pack(side="left", padx=10)
    
    # Run the application
    root.mainloop()

import tkinter as tk
import os
import sys
import traceback
from tkinter import messagebox, filedialog

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from approver.gui.pyGui import run_gui
from approver.Objects.BrowserAutomation import EdgeAutomation

def run_setup():
    print("[LOG] Starting Approver setup")
    try:
        # Use the WebDriver path from link.txt
        driver_path = None
        link_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "link.txt")
        
        if os.path.exists(link_file):
            with open(link_file, 'r') as f:
                driver_path = f.read().strip()
                print(f"[LOG] Found WebDriver path in link.txt: {driver_path}")
        
        browser_automation = EdgeAutomation()
        print("[LOG] Created EdgeAutomation instance")
        
        if browser_automation.initialize_driver():
            print("[LOG] WebDriver initialized successfully, launching main GUI")
            run_gui(browser_automation)
        else:
            print("[LOG] Failed to initialize WebDriver")
            messagebox.showerror("Error", "Failed to initialize WebDriver")
    except Exception as e:
        error_details = traceback.format_exc()
        print(f"[LOG] ERROR in run_setup: {str(e)}")
        print(f"[LOG] {error_details}")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    run_setup()
