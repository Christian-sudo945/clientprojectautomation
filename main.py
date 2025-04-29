import tkinter as tk
from tkinter import messagebox
import os, sys, threading, winreg, urllib.request, zipfile, ssl, csv
from datetime import datetime
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

class ModernSAPApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Automation Tool - IT Authorization")
        self.geometry("1000x680")
        self.resizable(False, False)
        
        # Define colors from flatly theme
        self.colors = {
            "primary": "#2C3E50",
            "secondary": "#95A5A6",
            "success": "#18BC9C",
            "info": "#3498DB",
            "warning": "#F39C12",
            "danger": "#E74C3C",
            "light": "#ECF0F1",
            "dark": "#7B8A8B",
            "bg": "#FFFFFF",
            "fg": "#2C3E50",
            "border": "#DEE2E6"
        }
        
        # Setup directories
        self.app_dir = os.path.dirname(os.path.abspath(__file__))
        self.driver_dir = os.path.join(self.app_dir, 'edgedriver_win64')
        self.driver_path = os.path.join(self.driver_dir, 'msedgedriver.exe')
        self.log_dir = os.path.join(self.app_dir, 'logs')
        self.main_log = os.path.join(self.log_dir, 'main_logs.csv')
        os.makedirs(self.driver_dir, exist_ok=True)
        os.makedirs(self.log_dir, exist_ok=True)
        self.webdriver_setup_complete = False
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Apply custom styles
        self.create_custom_styles()
        
        # Create UI
        self.create_ui_components()
        
        # Track running tools
        self.running_approver = False
        self.running_dispatcher = False
        
        # Start webdriver check
        threading.Thread(target=self.check_webdriver_on_startup, daemon=True).start()

    def create_custom_styles(self):
        """Create custom ttkbootstrap styles"""
        # Configure custom styles for consistent look and feel
        self.custom_style = ttk.Style()
        
        # Card frame style
        self.custom_style.configure("Card.TFrame", background=self.colors["bg"])
        
        # Custom button styles
        self.custom_style.configure("Modern.TButton", 
                            font=("Aptos", 11),
                            padding=8)
        
        # Label styles
        self.custom_style.configure("Title.TLabel", 
                           font=("Aptos", 22, "bold"),
                           foreground=self.colors["primary"])
        
        self.custom_style.configure("Subtitle.TLabel", 
                           font=("Aptos", 13),
                           foreground=self.colors["secondary"])
        
        self.custom_style.configure("SectionHeader.TLabel", 
                           font=("Aptos", 16, "bold"),
                           foreground=self.colors["primary"])
                           
        self.custom_style.configure("CardTitle.TLabel", 
                           font=("Aptos", 14, "bold"),
                           foreground=self.colors["primary"])
        
        self.custom_style.configure("CardDescription.TLabel", 
                           font=("Aptos", 12),
                           foreground=self.colors["secondary"])
        
        # Status label style
        self.custom_style.configure("Status.TLabel", 
                           font=("Aptos", 12),
                           padding=(5, 5))

    def create_ui_components(self):
        self.create_header()
        self.create_main_content()
        self.create_footer()

    def create_header(self):
        header_frame = ttk.Frame(self, padding=(25, 30, 25, 5))
        header_frame.grid(row=0, column=0, sticky="ew")
        
        # Logo and subtitle
        logo_label = ttk.Label(
            header_frame, 
            text="AUTOMATION TOOL", 
            style="Title.TLabel"
        )
        logo_label.pack(anchor="center")
        
        subtitle_label = ttk.Label(
            header_frame, 
            text="IT Authorization", 
            style="Subtitle.TLabel"
        )
        subtitle_label.pack(anchor="center", pady=(3, 0))

    def create_main_content(self):
        main_frame = ttk.Frame(self, padding=(25, 10, 25, 10))
        main_frame.grid(row=1, column=0, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        
        self.content_frame = ttk.Frame(main_frame)
        self.content_frame.grid(row=0, column=0, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        
        # Create status section
        self.create_status_section(self.content_frame)
        
        # Create options section
        self.create_options_section(self.content_frame)
        
        # Create tools section
        self.create_tools_section(self.content_frame)

    def create_status_section(self, parent):
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=0, column=0, sticky="ew", pady=(0, 25))
        
        # Add border around status section
        status_labelframe = ttk.Labelframe(
            status_frame, 
            text="System Status",
            bootstyle="primary",
            padding=(15, 12)
        )
        status_labelframe.pack(fill="x", expand=True)
        
        # Status indicator
        status_container = ttk.Frame(status_labelframe)
        status_container.pack(fill="x", expand=True)
        
        self.status_indicator = ttk.Label(
            status_container,
            text="●",
            font=("Aptos", 14),
            bootstyle="secondary"
        )
        self.status_indicator.pack(side="left", padx=(0, 8))
        
        self.status_label = ttk.Label(
            status_container,
            text="Checking WebDriver...",
            style="Status.TLabel"
        )
        self.status_label.pack(side="left")

    def create_options_section(self, parent):
        options_frame = ttk.Frame(parent)
        options_frame.grid(row=1, column=0, sticky="ew", pady=(0, 25))
        
        options_label = ttk.Label(
            options_frame,
            text="Automation Options",
            style="SectionHeader.TLabel"
        )
        options_label.pack(anchor="w", pady=(0, 15))
        
        options_container = ttk.Frame(options_frame)
        options_container.pack(fill="x")
        
        # Option buttons
        self.option_buttons = {}
        options = ["All Tools", "Approver", "Dispatcher"]
        
        for option in options:
            btn = ttk.Button(
                options_container,
                text=option,
                style="Modern.TButton",
                bootstyle=f"{'primary' if option == 'All Tools' else 'outline-primary'}",
                command=lambda opt=option: self.select_option(opt)
            )
            btn.pack(side="left", padx=(0, 12))
            self.option_buttons[option] = btn

    def create_tools_section(self, parent):
        tools_container = ttk.Frame(parent)
        tools_container.grid(row=2, column=0, sticky="ew")
        tools_container.grid_columnconfigure(0, weight=1)
        tools_container.grid_columnconfigure(1, weight=1)
        
        # Create both tool cards with same dimensions
        self.approver_card = self.create_tool_card(
            tools_container,
            "Approver Tool",
            "Automate approval processes in Work Inbox",
            0, 0,
            self.run_approver_thread
        )
        
        self.dispatcher_card = self.create_tool_card(
            tools_container,
            "Dispatcher Tool",
            "Automate dispatching in SAP CUP Tool",
            0, 1,
            self.run_dispatcher_thread
        )

    def create_tool_card(self, parent, title, description, row, col, command):
        # Card container with subtle padding around it
        card_container = ttk.Frame(parent, padding=(5, 5))
        card_container.grid(row=row, column=col, sticky="nsew", padx=5, pady=5)
        
        # Actual card with border
        card = ttk.Labelframe(
            card_container,
            text=title,
            bootstyle="primary",
            padding=(20, 18)
        )
        card.pack(fill="both", expand=True)
        
        # Tool icon (using emoji as placeholder)
        icon_text = "🔄" if "Approver" in title else "📋"
        icon_label = ttk.Label(
            card,
            text=icon_text,
            font=("Aptos", 22),
            bootstyle="primary"
        )
        icon_label.pack(anchor="w", pady=(0, 10))
        
        # Description with better styling
        desc_label = ttk.Label(
            card,
            text=description,
            style="CardDescription.TLabel",
            wraplength=400,
            justify="left"
        )
        desc_label.pack(fill="x", pady=(0, 20))
        
        # Separator to visually divide content from actions
        separator = ttk.Separator(card, bootstyle="secondary")
        separator.pack(fill="x", pady=(0, 20))
        
        # Action frame
        action_frame = ttk.Frame(card)
        action_frame.pack(fill="x")
        
        # Status indicator (when tool is running)
        status_frame = ttk.Frame(action_frame)
        status_frame.pack(side="left", fill="y")
        
        # Play button with updated design
        button_text = "Run"
        play_button = ttk.Button(
            action_frame,
            text=button_text,
            bootstyle="success",
            style="Modern.TButton",
            command=command,
            width=12
        )
        play_button.pack(side="right")
        
        # Store button reference
        if title == "Approver Tool":
            self.approver_button = play_button
        elif title == "Dispatcher Tool":
            self.dispatcher_button = play_button
        
        return card_container

    def create_footer(self):
        footer_frame = ttk.Frame(self, padding=(0, 5, 0, 15))
        footer_frame.grid(row=2, column=0, sticky="ew")
        
        # Add separator above footer
        separator = ttk.Separator(footer_frame, bootstyle="secondary")
        separator.pack(fill="x", pady=(0, 10))
        
        credits_label = ttk.Label(
            footer_frame,
            text="IT Authorization Automation Tool | Created by Aurora Benavidez | April 2025 ",
            font=("Aptos", 9),
            bootstyle="secondary"
        )
        credits_label.pack(side="bottom", anchor="center")

    def select_option(self, option):
        # Update button styles
        for opt, btn in self.option_buttons.items():
            if opt == option:
                btn.configure(bootstyle="primary")
            else:
                btn.configure(bootstyle="outline-primary")
        
        # Find all tool cards in the tools frame
        tools_frame = None
        for widget in self.content_frame.winfo_children():
            if widget.grid_info().get('row') == 2:
                tools_frame = widget
                break
        
        if not tools_frame:
            self.log_message("Error: Could not find tools frame")
            return
        
        # Apply the selected filter with smooth transition (visual effect)
        if option == "Approver":
            if self.approver_card:
                self.approver_card.grid(row=0, column=0, sticky="nsew", padx=5, pady=5, columnspan=2)
            if self.dispatcher_card:
                self.dispatcher_card.grid_remove()
            self.log_message("Showing Approver Tool only")
        elif option == "Dispatcher":
            if self.dispatcher_card:
                self.dispatcher_card.grid(row=0, column=0, sticky="nsew", padx=5, pady=5, columnspan=2)
            if self.approver_card:
                self.approver_card.grid_remove()
            self.log_message("Showing Dispatcher Tool only")
        else:  # "All Tools"
            if self.approver_card:
                self.approver_card.grid(row=0, column=0, sticky="nsew", padx=5, pady=5, columnspan=1)
            if self.dispatcher_card:
                self.dispatcher_card.grid(row=0, column=1, sticky="nsew", padx=5, pady=5, columnspan=1)
            self.log_message("Showing all tools")

    def update_status_indicator(self, status):
        status_map = {
            "idle": ("secondary", "Idle"),
            "loading": ("warning", "Loading..."),
            "success": ("success", "Success"),
            "error": ("danger", "Error")
        }
        bootstyle, _ = status_map.get(status, ("secondary", "Idle"))
        self.status_indicator.configure(bootstyle=bootstyle)

    def update_status_text(self, text):
        self.status_label.configure(text=text)

    def check_webdriver_on_startup(self):
        self.update_status_indicator("loading")
        self.update_status_text("Checking Edge WebDriver...")
        
        if not os.path.exists(self.driver_path):
            self.update_status_text("WebDriver not found. Will download when needed.")
            self.log_message("WebDriver not found. Will download when needed.")
            self.update_status_indicator("idle")
        else:
            self.update_status_text("Edge WebDriver ready")
            self.log_message("WebDriver found on startup")
            self.update_status_indicator("success")
            self.webdriver_setup_complete = True

    def run_approver_thread(self):
        if self.running_approver:
            return  # Prevent multiple instances
            
        self.running_approver = True
        self.approver_button.configure(text="Running", bootstyle="success-outline", state="disabled")
        threading.Thread(target=self.open_approver, daemon=True).start()

    def run_dispatcher_thread(self):
        if self.running_dispatcher:
            return  # Prevent multiple instances
            
        self.running_dispatcher = True
        self.dispatcher_button.configure(text="Running", bootstyle="success-outline", state="disabled")
        threading.Thread(target=self.open_dispatcher, daemon=True).start()

    def log_message(self, message):
        file_exists = os.path.isfile(self.main_log)
        with open(self.main_log, mode='a', newline='') as file:
            writer = csv.writer(file)
            if not file_exists:
                writer.writerow(['Timestamp', 'Message'])
            writer.writerow([datetime.now(), message])

    def open_approver(self):
        self.update_status_indicator("loading")
        self.update_status_text("Preparing WebDriver...")
        
        if not self.ensure_webdriver_setup():
            messagebox.showerror("Error", "Failed to set up WebDriver. Please try again.")
            self.update_status_indicator("error")
            self.update_status_text("WebDriver setup failed")
            # Reset button
            self.reset_approver_button()
            return
            
        self.update_status_text("Launching Approver Tool...")
        self.withdraw()
        
        try:
            approver_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "approver")
            if approver_path not in sys.path:
                sys.path.append(approver_path)
                
            with open(os.path.join(approver_path, "link.txt"), "w") as f:
                f.write(self.driver_path)
                
            self.log_message("Launching approver module")
            from approver.gui.setupGui import run_setup
            run_setup()
            
            self.log_message("Approver module closed")
            self.update_status_indicator("success")
            self.update_status_text("Approver task completed")
            
        except Exception as e:
            self.log_message(f"Error launching approver: {str(e)}")
            messagebox.showerror("Error", f"Error launching Approver: {str(e)}")
            self.update_status_indicator("error")
            self.update_status_text("Approver error occurred")
            
        finally:
            self.deiconify()
            self.reset_approver_button()

    def reset_approver_button(self):
        """Reset approver button to initial state"""
        self.running_approver = False
        self.approver_button.configure(text="Run", bootstyle="success", state="normal")

    def reset_dispatcher_button(self):
        """Reset dispatcher button to initial state"""
        self.running_dispatcher = False
        self.dispatcher_button.configure(text="Run", bootstyle="success", state="normal")

    def open_dispatcher(self):
        self.update_status_indicator("loading")
        self.update_status_text("Preparing WebDriver...")
        
        if not self.ensure_webdriver_setup():
            messagebox.showerror("Error", "Failed to set up WebDriver. Please try again.")
            self.update_status_indicator("error")
            self.update_status_text("WebDriver setup failed")
            # Reset button
            self.reset_dispatcher_button()
            return
            
        self.update_status_text("Launching Dispatcher Tool...")
        self.withdraw()
        
        try:
            dispatcher_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dispatcher")
            if dispatcher_path not in sys.path:
                sys.path.append(dispatcher_path)
                
            with open(os.path.join(dispatcher_path, "link.txt"), "w") as f:
                f.write(self.driver_path)
                
            self.log_message("Launching dispatcher module")
            from dispatcher.gui.setupGui import run_setup
            run_setup()
            
            self.log_message("Dispatcher module closed")
            self.update_status_indicator("success")
            self.update_status_text("Dispatcher task completed")
            
        except Exception as e:
            self.log_message(f"Error launching dispatcher: {str(e)}")
            messagebox.showerror("Error", f"Error launching Dispatcher: {str(e)}")
            self.update_status_indicator("error")
            self.update_status_text("Dispatcher error occurred")
            
        finally:
            self.deiconify()
            self.reset_dispatcher_button()

    def ensure_webdriver_setup(self):
        if self.webdriver_setup_complete:
            return True
            
        if not os.path.exists(self.driver_path):
            self.update_status_text("WebDriver not found. Downloading...")
            if not self.download_edge_driver():
                self.update_status_text("Failed to download WebDriver.")
                return False
                
        self.webdriver_setup_complete = True
        return True

    def get_edge_version(self):
        try:
            key_path = r"SOFTWARE\Microsoft\Edge\BLBeacon"
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
            version, _ = winreg.QueryValueEx(key, "version")
            winreg.CloseKey(key)
            self.log_message(f"Detected Edge version: {version}")
            return version
        except Exception as e:
            self.log_message(f"Error detecting Edge version: {str(e)}")
            return None

    def download_edge_driver(self):
        edge_version = self.get_edge_version()
        if not edge_version:
            messagebox.showerror("Error", "Could not detect Microsoft Edge version.")
            return False
            
        major_version = edge_version.split('.')[0]
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        
        try:
            driver_version_url = f"https://msedgedriver.azureedge.net/LATEST_RELEASE_{major_version}"
            req = urllib.request.Request(driver_version_url)
            req.add_header('User-Agent', 'Mozilla/5.0')
            
            with urllib.request.urlopen(req, timeout=15, context=ctx) as response:
                driver_version = response.read().decode('utf-8').strip()
                
            driver_url = f"https://msedgedriver.azureedge.net/{driver_version}/edgedriver_win64.zip"
            zip_path = os.path.join(self.driver_dir, "msedgedriver.zip")
            
            req = urllib.request.Request(driver_url)
            req.add_header('User-Agent', 'Mozilla/5.0')
            
            with urllib.request.urlopen(req, timeout=60, context=ctx) as response:
                with open(zip_path, 'wb') as f:
                    f.write(response.read())
                    
            with zipfile.ZipFile(zip_path) as zip_ref:
                zip_ref.extractall(self.driver_dir)
                
            os.remove(zip_path)
            self.log_message("WebDriver downloaded successfully")
            self.update_status_indicator("success")
            self.update_status_text("WebDriver downloaded successfully")
            return True
            
        except Exception as e:
            self.log_message(f"Error downloading WebDriver: {str(e)}")
            self.update_status_indicator("error")
            self.update_status_text(f"WebDriver download failed")
            return False


if __name__ == "__main__":
    app = ModernSAPApp()
    app.mainloop()