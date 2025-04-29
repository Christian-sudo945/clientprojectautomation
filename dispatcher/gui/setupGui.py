import tkinter as tk
import os
import sys
import traceback
from tkinter import messagebox, filedialog
from dispatcher.gui.pyGui import run_gui
from dispatcher.Objects.BrowserAutomation import EdgeAutomation

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from dispatcher.gui.pyGui import run_gui
from dispatcher.Objects.BrowserAutomation import EdgeAutomation

def run_setup():
    print("Starting Dispatcher setup")
    try:
        # Use the WebDriver path from link.txt
        driver_path = None
        link_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "link.txt")
        
        if os.path.exists(link_file):
            with open(link_file, 'r') as f:
                driver_path = f.read().strip()
                print(f"Found WebDriver path in link.txt: {driver_path}")
        
        browser_automation = EdgeAutomation()
        print("Created EdgeAutomation instance")
        
        if browser_automation.initialize_driver():
            print("WebDriver initialized successfully, launching main GUI")
            run_gui(browser_automation)
        else:
            print("Failed to initialize WebDriver")
            messagebox.showerror("Error", "Failed to initialize WebDriver")
    except Exception as e:
        error_details = traceback.format_exc()
        print(f"ERROR in run_setup: {str(e)}")
        print(f"{error_details}")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    run_setup()
