import os
import webbrowser
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from PIL import Image, ImageTk

# Loading indicator state
_loading_window = None
_loading_label = None
_loading_progress = None

def toggle_loading(root, state, message="Processing..."):
    """Modern styled loading indicator"""
    global _loading_window, _loading_label, _loading_progress

    if state:  # Show loading window
        if _loading_window is None:
            # Create a modern loading overlay
            _loading_window = tk.Toplevel(root)
            _loading_window.title("Processing")
            _loading_window.overrideredirect(True)  # Remove window decorations
            
            # Semi-transparent background
            _loading_window.attributes('-alpha', 0.95)
            
            # Configure colors
            bg_color = "#ffffff"
            accent_color = "#2196f3"
            
            # Window size
            width, height = 400, 150
            
            # Center the window
            root.update_idletasks()
            root_x = root.winfo_x()
            root_y = root.winfo_y()
            root_width = root.winfo_width()
            root_height = root.winfo_height()
            
            x = root_x + (root_width - width) // 2
            y = root_y + (root_height - height) // 2
            
            _loading_window.geometry(f"{width}x{height}+{x}+{y}")
            _loading_window.configure(bg=bg_color, highlightbackground=accent_color, highlightthickness=2)
            
            # Create content
            content_frame = tk.Frame(_loading_window, bg=bg_color)
            content_frame.pack(expand=True, fill="both", padx=30, pady=30)
            
            # Title
            title_label = tk.Label(content_frame, text="Processing", 
                                  font=("Segoe UI", 14, "bold"),
                                  bg=bg_color, fg=accent_color)
            title_label.pack(pady=(0, 15))
            
            # Message
            _loading_label = tk.Label(content_frame, text=message,
                                     font=("Segoe UI", 10),
                                     bg=bg_color, fg="#546e7a")
            _loading_label.pack(pady=(0, 20))
            
            # Progress bar with custom style
            style = ttk.Style()
            style.theme_use('clam')
            style.configure("Modern.Horizontal.TProgressbar",
                           troughcolor="#f0f0f0",
                           background=accent_color,
                           borderwidth=0,
                           lightcolor=accent_color,
                           darkcolor=accent_color)
            
            _loading_progress = ttk.Progressbar(
                content_frame,
                mode="indeterminate",
                length=300,
                style="Modern.Horizontal.TProgressbar"
            )
            _loading_progress.pack()
            _loading_progress.start(10)
            
            # Force immediate display
            _loading_window.update_idletasks()
            
        else:
            # Update existing window
            _loading_label.config(text=message)
            _loading_window.update_idletasks()

    elif _loading_window:  # Hide loading window
        _loading_progress.stop()
        _loading_window.destroy()
        _loading_window = None
        root.update_idletasks()

def create_status_bar(parent):
    """Create a modern status bar"""
    status_bar = tk.Frame(parent, bg="#f8f9fa", height=30)
    status_bar.pack(fill="x", side="bottom")
    status_bar.pack_propagate(False)
    
    # Status message
    status_label = tk.Label(status_bar, text="Ready", font=("Segoe UI", 9),
                           bg="#f8f9fa", fg="#6c757d", anchor="w")
    status_label.pack(side="left", padx=15)
    
    # Version info
    version_label = tk.Label(status_bar, text="Version 0.3.1", font=("Segoe UI", 9),
                            bg="#f8f9fa", fg="#adb5bd", anchor="e")
    version_label.pack(side="right", padx=15)
    
    return status_label

def create_quick_links(parent):
    """Create quick access links"""
    links_frame = tk.Frame(parent, bg=parent.cget('bg'))
    links_frame.pack(pady=(20, 0))
    
    links = [
        ("📁 Results", "Results"),
        ("📁 Imported", "Imported_Results"),
        ("📁 Logs", "Logs"),
        ("❓ Help", "help")
    ]
    
    def open_directory(directory):
        """Open the specified directory"""
        if directory == "help":
            webbrowser.open("https://info.support.huawei.com/storageinfo/refer/#/home")
        else:
            path = os.path.join(os.path.dirname(os.path.abspath(__file__)), directory)
            if os.path.exists(path):
                os.startfile(path)
            else:
                messagebox.showinfo("Info", f"Directory not found:\n{path}")
    
    for icon_text, directory in links:
        link = tk.Label(
            links_frame,
            text=icon_text,
            font=("Segoe UI", 9),
            fg="#2196f3",
            bg=parent.cget('bg'),
            cursor="hand2"
        )
        link.pack(side="left", padx=15)
        
        # Add hover effect
        link.bind("<Enter>", lambda e, l=link: l.config(font=("Segoe UI", 9, "underline")))
        link.bind("<Leave>", lambda e, l=link: l.config(font=("Segoe UI", 9)))
        
        # Bind click event
        link.bind("<Button-1>", lambda e, d=directory: open_directory(d))

def create_info_panel(parent):
    """Create an information panel with tips"""
    panel = tk.Frame(parent, bg="#e3f2fd", relief="flat", bd=1,
                    highlightbackground="#bbdefb", highlightthickness=1)
    
    # Header
    header = tk.Label(panel, text="💡 Quick Tips", font=("Segoe UI", 10, "bold"),
                     bg="#e3f2fd", fg="#1565c0")
    header.pack(anchor="w", padx=15, pady=(10, 5))
    
    # Tips
    tips = [
        "• Select a device type and resource to begin",
        "• Click 'Open Excel' to create/edit templates",
        "• Fill in Excel sheets and click 'Run Script'",
        "• Use 'Import Excel' to process existing files"
    ]
    
    for tip in tips:
        tip_label = tk.Label(panel, text=tip, font=("Segoe UI", 9),
                           bg="#e3f2fd", fg="#37474f", anchor="w")
        tip_label.pack(anchor="w", padx=20, pady=2)
    
    panel.pack(fill="x", padx=20, pady=(20, 0))
    
    return panel