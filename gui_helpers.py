import tkinter as tk
from tkinter import ttk, font as tkfont
from gui_functions import *

def apply_style(root):
    """Apply consistent styling to the GUI"""
    root.configure(bg="#FFFFFF")
    style = ttk.Style()
    style.theme_use("clam")
    
    # Configure combobox style
    style.configure("TCombobox",
                    font=("Arial", 11),
                    padding=8,
                    relief="flat",
                    anchor="center",
                    #foreground="#333333",  # Dark grey text
                    background="#F5F5F5",  # Light grey background
                    highlightthickness=0,
                    arrowsize=15)
    
    style.map("TCombobox",
             fieldbackground=[("readonly", "#FFFFFF")],
             selectbackground=[("readonly", "#F0F0F0")],
             selectforeground=[("readonly", "#333333")])
    
    # Title label
    title_font = tkfont.Font(family="Arial", size=14, weight="bold")
    title = tk.Label(
        root,
        text="OceanStor Command Generator",
        font=title_font,
        bg="#FFFFFF",
        fg="#333333",
        pady=10
    )
    title.pack(pady=(10, 5))

def create_buttons_and_dropdown(root, open_excel, run_script, script_var, command_var, clear_excel, import_excel, device_var, config):
    """Create the main UI components"""
    # Create script selection dropdown
    script_menu = create_script_menu(root, script_var, config)
    
    # Create command operation buttons
    command_frame, update_command_buttons = create_command_buttons(root, script_var, command_var, config)
    
    # Create action buttons
    action_frame = create_action_buttons(
        root, 
        open_excel, 
        run_script, 
        script_var, 
        command_var, 
        clear_excel, 
        import_excel, 
        device_var
    )
    
    # Add helper elements
    add_help_button(root)
    add_directory_links(root)
    add_version_label(root)
    
    return script_menu, *action_frame[:3], update_command_buttons