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

def create_script_menu(root, script_var, config):
    """Create and return the script selection dropdown"""
    script_menu = ttk.Combobox(
        root,
        textvariable=script_var,
        values=list(config.keys()),
        state="readonly",
        font=("Arial", 11),
        justify="center",
    )
    script_menu.pack(pady=10)
    return script_menu

def create_command_buttons(root, script_var, command_var, config):
    """Create command operation buttons frame and return it with update function"""
    command_frame = tk.Frame(root, bg="#FFFFFF")
    command_frame.pack(pady=(10, 15))

    def update_command_buttons(config):
        """Update command buttons based on current script selection"""
        for widget in command_frame.winfo_children():
            widget.destroy()
        
        current_script = script_var.get()
        if current_script in config:
            for operation in config[current_script]['operations'].keys():
                btn = create_command_button(operation, command_frame)
                btn.pack(side="left", padx=5)
                btn.config(
                    command=lambda op=operation: select_command(op, command_var, command_frame),
                    **button_style(operation == command_var.get())
                )
                setup_button_hover(btn, "#E0E0E0", "#F0F0F0")

    # Initialize buttons
    if config:
        script_var.set(list(config.keys())[0])
        update_command_buttons(config)

    return command_frame, update_command_buttons

def create_command_button(text, command_frame):
    """Create a standard command button"""
    return tk.Button(
        command_frame,
        text=text,
        width=12,
        height=1,
        bg="#F0F0F0",
        fg="#333333",
        activebackground="#E0E0E0",
        relief="flat",
        bd=1,
        padx=5,
        pady=3,
        font=("Arial", 10)
    )

def button_style(active):
    """Return style properties based on active state"""
    return {
        'bg': "#E0E0E0" if active else "#F0F0F0",
        'relief': "sunken" if active else "flat"
    }

def select_command(command, command_var, command_frame):
    """Handle command button selection"""
    command_var.set(command)
    for btn in command_frame.winfo_children():
        btn.config(**button_style(btn.cget("text") == command))

def setup_button_hover(button, hover_color, default_color):
    """Configure hover effects for buttons"""
    button.bind("<Enter>", lambda e, b=button: b.config(bg=hover_color))
    button.bind("<Leave>", lambda e, b=button: b.config(bg=default_color))

def create_action_buttons(root, open_excel, run_script, script_var, command_var, clear_excel, import_excel, device_var):
    """Create and return action buttons frame"""
    action_frame = tk.Frame(root, bg="#FFFFFF")
    action_frame.pack(pady=(15, 10))

    # Create action buttons
    open_button = create_colored_button(
        action_frame,
        "Open Excel", "#4CAF50",
        lambda: open_excel(script_var.get(), command_var.get(), device_var.get())
    )
    clear_button = create_colored_button(
        action_frame,
        "Clear Excel", "#F44336",
        lambda: clear_excel(script_var.get(), device_var.get())
    )
    import_button = create_colored_button(
        action_frame,
        "Import Excel", "#FF9800",
        lambda: import_excel(root)
    )
    run_button = create_colored_button(
        root,
        "Run Script", "#2196F3",
        lambda: run_script(script_var.get(), command_var.get(), device_var.get()),
        pady=(15, 10)
    )

    return action_frame, open_button, run_button, clear_button

def create_colored_button(parent, text, color, command, **pack_args):
    """Create a colored action button with hover effects"""
    button = tk.Button(
        parent,
        text=text,
        width=12,
        height=1,
        bg=color,
        fg="white",
        activebackground=darker_color(color),
        relief="flat",
        bd=1,
        padx=5,
        pady=5,
        font=("Arial", 10, "bold"),
        command=command
    )
    setup_button_hover(button, darker_color(color, 0.9), color)
    button.pack(**pack_args) if pack_args else button.pack(side="left", padx=5)
    return button

def darker_color(hex_color, factor=0.8):
    """Return a darker version of the given hex color"""
    rgb = tuple(int(hex_color[i:i+2], 16) for i in (1, 3, 5))
    darker = tuple(max(0, int(c * factor)) for c in rgb)
    return f"#{darker[0]:02x}{darker[1]:02x}{darker[2]:02x}"

def add_help_button(root):
    """Add help button to top-right corner"""
    try:
        icon_path = os.path.join("Icons", "help.ico")
        help_icon = Image.open(icon_path).resize((20, 20), Image.LANCZOS)
        help_photo = ImageTk.PhotoImage(help_icon)
        
        help_button = tk.Button(
            root,
            image=help_photo,
            command=lambda: webbrowser.open("https://info.support.huawei.com/storageinfo/refer/#/home"),
            bg="#FFFFFF",
            activebackground="#FFFFFF",
            relief="flat",
            bd=0,
            highlightthickness=0
        )
        help_button.image = help_photo
        help_button.place(relx=0.98, rely=0.02, anchor="ne")
        help_button.bind("<Enter>", lambda e: help_button.config(cursor="hand2"))
    except Exception:
        # Fallback text button
        help_button = tk.Button(
            root,
            text="?",
            command=lambda: webbrowser.open("https://info.support.huawei.com/storageinfo/refer/#/home"),
            bg="#FFFFFF",
            fg="#333333",
            relief="flat",
            bd=0,
            font=("Arial", 10, "bold")
        )
        help_button.place(relx=0.98, rely=0.02, anchor="ne")
        
def add_version_label(root):
    """Add version label at bottom right"""
    version_label = tk.Label(root, 
                             text="Version 0.2.0", 
                             bg="#FFFFFF", 
                             font=("Helvetica", 9), 
                             fg="#808080")  # Set the text color to greyish
    version_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)  # Positioned at bottom-right corner with some margin
    return version_label


def add_directory_links(root):
    """Add directory links at bottom"""
    links_frame = tk.Frame(root, bg="#FFFFFF")
    links_frame.pack(pady=(30, 15))

    for text, directory in [("Open Results", "Results"), 
                          ("Open Imported", "Imported_Results"),
                          ("Open Logs", "Logs")]:
        create_directory_link(links_frame, text, directory)

def create_directory_link(parent, text, directory):
    """Create a single directory link"""
    link = tk.Label(
        parent,
        text=text,
        font=("Arial", 9, "underline"),
        fg="#666666",
        cursor="hand2",
        bg="#FFFFFF"
    )
    link.pack(side="left", padx=10)
    link.bind("<Button-1>", lambda e: open_directory(directory))

def open_directory(directory):
    """Open the specified directory"""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), directory)
    if os.path.exists(path):
        os.startfile(path)
    else:
        messagebox.showinfo("Info", f"Directory not found:\n{path}")

def toggle_loading(root, state, message="Processing..."):
    """Show/hide loading indicator"""
    global _loading_window, _loading_label, _loading_progress
    
    if state:
        if _loading_window is None:
            _loading_window = tk.Toplevel(root)
            _loading_window.overrideredirect(True)
            _loading_window.geometry("300x100")
            _loading_window.transient(root)
            _loading_window.grab_set()
            
            # Center window
            root.update_idletasks()
            x = root.winfo_x() + (root.winfo_width() - 300) // 2
            y = root.winfo_y() + (root.winfo_height() - 100) // 2
            _loading_window.geometry(f"+{x}+{y}")
            
            # Content
            _loading_label = tk.Label(
                _loading_window,
                text=message,
                font=("Arial", 10),
                bg="#F0F0F0",
                pady=10
            )
            _loading_label.pack(fill="x")
            
            _loading_progress = ttk.Progressbar(
                _loading_window,
                mode="indeterminate",
                length=250
            )
            _loading_progress.pack(pady=5)
            _loading_progress.start()
    elif _loading_window:
        _loading_progress.stop()
        _loading_window.grab_release()
        _loading_window.destroy()
        _loading_window = None