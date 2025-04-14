import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont
from tkinter import messagebox
import json
import webbrowser
import os

# Global variable to track loading state
_loading_window = None
_loading_label = None
_loading_progress = None

def apply_style(root):
    """Apply a professional and consistent style to the GUI."""
    root.configure(bg="#FFFFFF")  # White background
    root.option_add("*Font", "Arial 11")  # Use Arial as the default font
    root.option_add("*Foreground", "#333333")  # Dark gray text
    root.option_add("*Button.Background", "#F0F0F0")  # Light gray background for buttons
    root.option_add("*Button.Foreground", "#333333")  # Dark gray text for buttons
    root.option_add("*Button.Relief", "flat")  # Flat button style
    root.option_add("*Button.BorderWidth", 1)  # Thin border
    root.option_add("*Button.Padding", 8)  # Padding for buttons

    # Custom font for the title
    title_font = tkfont.Font(family="Arial", size=14, weight="bold")  # Use Arial for the title
    label = tk.Label(
        root,
        text="Select a script to execute or manage configuration:",
        font=title_font,
        bg="#FFFFFF",
        fg="#333333",
        pady=10,
    )
    label.pack(pady=(20, 10))

def create_buttons_and_dropdown(root, open_excel, run_script, script_var, command_var, clear_excel, import_excel, device_var, config):
    """Create buttons and dropdowns for the GUI with a professional look."""
    style = ttk.Style()
    style.theme_use("clam")
    
    # Style for the Combobox
    style.configure("TCombobox",
                    fieldbackground="#FFFFFF",
                    background="#FFFFFF",
                    foreground="#333333",
                    padding=5,
                    bordercolor="#CCCCCC",
                    lightcolor="#CCCCCC",
                    darkcolor="#CCCCCC")
    style.map("TCombobox",
              fieldbackground=[("readonly", "#FFFFFF")],
              selectbackground=[("readonly", "#F0F0F0")],
              selectforeground=[("readonly", "#333333")])

    # Script selection dropdown
    script_menu = ttk.Combobox(
        root,
        textvariable=script_var,
        values=list(config.keys()),
        state="readonly",
        font=("Arial", 11),
        justify="center",
    )
    script_menu.pack(pady=10)

    # Frame for command buttons
    command_buttons_frame = tk.Frame(root, bg="#FFFFFF")
    command_buttons_frame.pack(pady=(20, 20))

    def on_enter(button, hover_color):
        button.config(bg=hover_color)

    def on_leave(button, original_color):
        button.config(bg=original_color)

    def update_command_buttons(config):
        """Updates command buttons without redundant calls."""
        selected_script = script_var.get()
        
        if selected_script not in config:
            return

        # Clear previous buttons
        for button in command_buttons_frame.winfo_children():
            button.destroy()

        # Create new buttons
        for operation in config[selected_script].keys():
            button = tk.Button(
                command_buttons_frame,
                text=operation,
                command=lambda op=operation: select_command(op),
                width=10, height=1, bg="#F0F0F0", fg="#333333",
                activebackground="#E0E0E0", relief="flat", bd=1, padx=8, pady=5
            )
            button.pack(side="left", padx=5)
            button.bind("<Enter>", lambda e, b=button: on_enter(b, "#E0E0E0"))
            button.bind("<Leave>", lambda e, b=button: on_leave(b, "#F0F0F0"))

        # Default selection
        if "Create" in config[selected_script]:
            select_command("Create")

    def select_command(command):
        command_var.set(command)
        for button in command_buttons_frame.winfo_children():
            if button.cget("text") == command:
                button.config(relief="sunken", bg="#E0E0E0")
            else:
                button.config(relief="flat", bg="#F0F0F0")

    # Initialize buttons
    update_command_buttons(config)

    # Frame for Excel buttons - now with equal widths
    excel_buttons_frame = tk.Frame(root, bg="#FFFFFF")
    excel_buttons_frame.pack(pady=(20, 10))

    # Button settings
    button_width = 14
    button_height = 1
    button_padx = 4
    
    # Open Excel button (Green)
    open_button = tk.Button(
        excel_buttons_frame,
        text="Open Excel",
        font=("Arial", 11, "bold"),
        command=lambda: open_excel(script_var.get(), command_var.get(), device_var.get()),
        width=button_width,
        height=button_height,
        bg="#4CAF50",
        fg="white",
        activebackground="#45A049",
        relief="flat",
        bd=1,
        padx=button_padx,
        pady=5,
    )
    open_button.pack(side="left", padx=5)
    open_button.bind("<Enter>", lambda e: on_enter(open_button, "#45A049"))
    open_button.bind("<Leave>", lambda e: on_leave(open_button, "#4CAF50"))

    # Clear Excel button (Red)
    clear_button = tk.Button(
        excel_buttons_frame,
        text="Clear Excel",
        font=("Arial", 11, "bold"),
        command=lambda: clear_excel(script_var.get(), device_var.get()),
        width=button_width,
        height=button_height,
        bg="#F44336",
        fg="white",
        activebackground="#D32F2F",
        relief="flat",
        bd=1,
        padx=button_padx,
        pady=5,
    )
    clear_button.pack(side="left", padx=5)
    clear_button.bind("<Enter>", lambda e: on_enter(clear_button, "#D32F2F"))
    clear_button.bind("<Leave>", lambda e: on_leave(clear_button, "#F44336"))

    # Import Excel button (Orange)
    import_button = tk.Button(
        excel_buttons_frame,
        text="Import Excel",
        font=("Arial", 11, "bold"),
        command=lambda: import_excel(root),
        width=button_width,
        height=button_height,
        bg="#FF9800",  # Orange color
        fg="white",
        activebackground="#F57C00",  # Darker orange
        relief="flat",
        bd=1,
        padx=button_padx,
        pady=5,
    )
    import_button.pack(side="left", padx=5)
    import_button.bind("<Enter>", lambda e: on_enter(import_button, "#F57C00"))
    import_button.bind("<Leave>", lambda e: on_leave(import_button, "#FF9800"))

    # Run Script button (Blue) - now same size as others
    run_button = tk.Button(
        root,
        text="Run Script",
        font=("Arial", 11, "bold"),
        command=lambda: run_script(script_var.get(), command_var.get(), device_var.get()),
        width=button_width,
        height=button_height,
        bg="#2196F3",
        fg="white",
        activebackground="#1976D2",
        relief="flat",
        bd=1,
        padx=8,
        pady=5,
    )
    run_button.pack(pady=(10, 10))

    # Add directory opening links
    def open_directory(directory):
        import os
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), directory)
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showinfo("Info", f"Directory not found: {path}")

    # Frame for directory links (horizontal layout)
    dir_links_frame = tk.Frame(root, bg="#FFFFFF")
    dir_links_frame.pack(pady=(20, 20))

    # Open Results link (grey)
    results_link = tk.Label(
        dir_links_frame,
        text="Open Results",
        font=("Arial", 10, "underline"),
        fg="#666666",  # Grey color
        cursor="hand2",
        bg="#FFFFFF"
    )
    results_link.pack(side="left", padx=10)
    results_link.bind("<Button-1>", lambda e: open_directory("Results"))

    # Open Imported Results link (grey)
    imported_results_link = tk.Label(
        dir_links_frame,
        text="Open Imported Results",
        font=("Arial", 10, "underline"),
        fg="#666666",  # Grey color
        cursor="hand2",
        bg="#FFFFFF"
    )
    imported_results_link.pack(side="left", padx=10)
    imported_results_link.bind("<Button-1>", lambda e: open_directory("Imported_Results"))
    
    # Open Logs link (grey)
    logs_link = tk.Label(
        dir_links_frame,
        text="Open Logs",
        font=("Arial", 10, "underline"),
        fg="#666666",  # Grey color
        cursor="hand2",
        bg="#FFFFFF"
    )
    logs_link.pack(side="left", padx=10)
    logs_link.bind("<Button-1>", lambda e: open_directory("Logs"))

    return script_menu, open_button, run_button, clear_button, update_command_buttons


def toggle_loading(root, state, message="Processing..."):
    """
    Show or hide a loading indicator.
    
    Args:
        root (tk.Tk|tk.Toplevel): The root window
        state (bool): True to show loading, False to hide
        message (str): Optional message to display
    """
    global _loading_window, _loading_label, _loading_progress
    
    if state:
        # Create loading overlay if it doesn't exist
        if _loading_window is None:
            _loading_window = tk.Toplevel(root)
            _loading_window.title("Please Wait")
            _loading_window.geometry("300x100")
            _loading_window.resizable(False, False)
            _loading_window.transient(root)  # Show above main window
            _loading_window.grab_set()  # Make modal
            
            # Center the window
            root.update_idletasks()
            x = root.winfo_x() + (root.winfo_width() // 2) - 150
            y = root.winfo_y() + (root.winfo_height() // 2) - 50
            _loading_window.geometry(f"+{x}+{y}")
            
            # Loading label
            _loading_label = ttk.Label(
                _loading_window,
                text=message,
                font=("Arial", 10)
            )
            _loading_label.pack(pady=10)
            
            # Progress bar
            _loading_progress = ttk.Progressbar(
                _loading_window,
                mode='indeterminate',
                length=200
            )
            _loading_progress.pack()
            _loading_progress.start()
            
        _loading_window.deiconify()  # Show if hidden
    else:
        # Hide and clean up loading overlay
        if _loading_window is not None:
            _loading_progress.stop()
            _loading_window.grab_release()
            _loading_window.destroy()
            _loading_window = None
            _loading_label = None
            _loading_progress = None
            
    root.update_idletasks()  # Force UI update