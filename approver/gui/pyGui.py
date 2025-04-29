import tkinter as tk
from tkinter import ttk, messagebox
import os
from auto_approve import auto_approve
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class RoundedButton(tk.Canvas):
    def __init__(self, parent, width, height, corner_radius, padding=0, color="#0c1236", text="", command=None, text_color="white", hover_color="#161e41", **kwargs):
        tk.Canvas.__init__(self, parent, width=width, height=height, highlightthickness=0, **kwargs)
        self.command = command
        self.color = color
        self.hover_color = hover_color
        self.corner_radius = corner_radius

        if corner_radius > min(width, height) / 2:
            self.corner_radius = min(width, height) / 2

        self.rect = self.rounded_rectangle(padding, padding, width - padding, height - padding, self.corner_radius, fill=color, outline="")
        self.text_id = self.create_text(width / 2, height / 2, text=text, fill=text_color, font=("Arial", 10, "bold"))
        
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def rounded_rectangle(self, x1, y1, x2, y2, radius, **kwargs):
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1
        ]
        return self.create_polygon(points, **kwargs, smooth=True)

    def _on_press(self, event):
        self.itemconfig(self.rect, fill=self.color)

    def _on_release(self, event):
        self.itemconfig(self.rect, fill=self.hover_color)
        if self.command:
            self.command()

    def _on_enter(self, event):
        self.itemconfig(self.rect, fill=self.hover_color)
        self.config(cursor="hand2")

    def _on_leave(self, event):
        self.itemconfig(self.rect, fill=self.color)
        self.config(cursor="")


def run_gui(browser_automation):
    root = ttk.Window(themename="flatly")
    root.title("Work Inbox Automation")
    root.geometry("800x500")
    root.resizable(False, False)
    
    # Define colors from flatly theme
    colors = {
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
    
    # Configure grid
    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(1, weight=1)
    
    # Create header section
    header_frame = ttk.Frame(root, padding=(25, 30, 25, 5))
    header_frame.grid(row=0, column=0, sticky="ew")
    
    # Logo and subtitle
    header_title = ttk.Label(
        header_frame, 
        text="WORK INBOX AUTOMATION", 
        style="Title.TLabel"
    )
    header_title.pack(anchor="center")
    
    subtitle = ttk.Label(
        header_frame, 
        text="Approver Tool", 
        style="Subtitle.TLabel"
    )
    subtitle.pack(anchor="center", pady=(3, 0))
    
    # Apply custom styles
    custom_style = ttk.Style()
    
    # Title styles
    custom_style.configure("Title.TLabel", 
                 font=("Aptos", 22, "bold"),
                 foreground=colors["primary"])
    
    custom_style.configure("Subtitle.TLabel", 
                 font=("Aptos", 13),
                 foreground=colors["secondary"])
    
    # Label styles
    custom_style.configure("SectionHeader.TLabel", 
                 font=("Aptos", 16, "bold"),
                 foreground=colors["primary"])
                 
    custom_style.configure("CardTitle.TLabel", 
                 font=("Aptos", 14, "bold"),
                 foreground=colors["primary"])
    
    custom_style.configure("CardDescription.TLabel", 
                 font=("Aptos", 12),
                 foreground=colors["secondary"])
    
    # Button style
    custom_style.configure("Modern.TButton", 
                  font=("Aptos", 11),
                  padding=8)
    
    # Status label style
    custom_style.configure("Status.TLabel", 
                 font=("Aptos", 12),
                 padding=(5, 5))
    
    # Main content section
    main_frame = ttk.Frame(root, padding=(25, 10, 25, 10))
    main_frame.grid(row=1, column=0, sticky="nsew")
    main_frame.grid_columnconfigure(0, weight=1)
    
    # Create card for form
    content_frame = ttk.Frame(main_frame)
    content_frame.grid(row=0, column=0, sticky="nsew")
    content_frame.grid_columnconfigure(0, weight=1)
    
    # Card content
    card_labelframe = ttk.Labelframe(
        content_frame,
        text="Enter Work Inbox Details",
        bootstyle="primary",
        padding=15
    )
    card_labelframe.pack(fill="both", expand=True)
    
    # Form contents
    form_frame = ttk.Frame(card_labelframe)
    form_frame.pack(fill="both", expand=True, padx=5, pady=10)
    form_frame.grid_columnconfigure(1, weight=1)
    
    # Form fields
    link_label = ttk.Label(
        form_frame,
        text="Work Inbox URL:",
        font=("Aptos", 10),
        bootstyle="primary"
    )
    link_label.grid(row=0, column=0, sticky="w", pady=15)
    
    entry_link = ttk.Entry(
        form_frame,
        font=("Aptos", 10),
        bootstyle="primary"
    )
    entry_link.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=15)
    
    # Status section
    status_frame = ttk.Frame(card_labelframe)
    status_frame.pack(fill="x", expand=False, padx=5, pady=(0, 10))
    
    status_container = ttk.Frame(status_frame)
    status_container.pack(fill="x", expand=True)
    
    status_indicator = ttk.Label(
        status_container,
        text="●",
        font=("Aptos", 14),
        bootstyle="secondary"
    )
    status_indicator.pack(side="left", padx=(0, 8))
    
    status_label = ttk.Label(
        status_container,
        text="Ready",
        style="Status.TLabel"
    )
    status_label.pack(side="left")
    
    # Button section
    button_frame = ttk.Frame(card_labelframe)
    button_frame.pack(fill="x", expand=False, padx=5, pady=(10, 5))
    
    # Configure button frame columns
    button_frame.grid_columnconfigure(0, weight=1)
    button_frame.grid_columnconfigure(1, weight=1)
    
    def cancel_action():
        root.destroy()
        
    def on_button_click():
        link = entry_link.get().strip()
        
        if not link:
            messagebox.showerror("Input Error", "Work Inbox Link cannot be empty.")
            return
        
        status_indicator.configure(bootstyle="warning")
        status_label.configure(text="Processing...")
        root.update_idletasks()
        
        try:
            auto_approve(browser_automation, link)
            status_indicator.configure(bootstyle="success")
            status_label.configure(text="Completed successfully")
            messagebox.showinfo("Success", "Task completed successfully")
        except Exception as e:
            status_indicator.configure(bootstyle="danger")
            status_label.configure(text="Error occurred")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    cancel_btn = ttk.Button(
        button_frame,
        text="CANCEL",
        bootstyle="secondary-outline",
        style="Modern.TButton",
        command=cancel_action,
        width=15
    )
    cancel_btn.grid(row=0, column=0, sticky="e", padx=(0, 10))
    
    submit_btn = ttk.Button(
        button_frame,
        text="START",
        bootstyle="success",
        style="Modern.TButton",
        command=on_button_click,
        width=15
    )
    submit_btn.grid(row=0, column=1, sticky="w", padx=(10, 0))
    
    # Footer section
    footer_frame = ttk.Frame(root, padding=(0, 5, 0, 15))
    footer_frame.grid(row=2, column=0, sticky="ew")
    
    # Add separator above footer
    separator = ttk.Separator(footer_frame, bootstyle="secondary")
    separator.pack(fill="x", pady=(0, 10))
    
    credits_label = ttk.Label(
        footer_frame,
        text="IT Authorization Automation Tool | Created by Aurora Benavidez | April 2025",
        font=("Aptos", 9),
        bootstyle="secondary"
    )
    credits_label.pack(side="bottom", anchor="center")
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()

if __name__ == "__main__":
    print("This module should be imported, not run directly.")
