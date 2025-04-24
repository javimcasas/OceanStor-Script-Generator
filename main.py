import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from excel_operations import open_excel, clear_excel, import_excel
from gui_helpers import apply_style, create_buttons_and_dropdown, toggle_loading
from command_generator import main as generate_commands
from utils import load_config

def run_script(script_type, command_type, device_type):
    try:
        config = load_config(device_type)
        if script_type not in config:
            raise ValueError(f"Invalid script type: {script_type}. Valid types are: {list(config.keys())}")

        documents_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Documents")
        excel_path = os.path.join(documents_folder, f'{device_type}_{script_type}_commands.xlsx')

        if not os.path.exists(excel_path):
            messagebox.showinfo("Info", f"Please create the Excel file first by pressing the 'Create Excel' button.")
            return

        # Show loading indicator
        toggle_loading(root, True, f"Generating {command_type} commands...")
        root.update()  # Force UI update
        
        # Call the command generator
        output_file_path = generate_commands(script_type, device_type)
        
        # Hide loading indicator
        toggle_loading(root, False)

        if os.path.exists(output_file_path):
            messagebox.showinfo("Success", f"Commands generated successfully!\nOutput saved to:\n{output_file_path}")
            os.startfile(output_file_path)
        else:
            messagebox.showerror("Error", f"Output file not found at:\n{output_file_path}")
    except Exception as e:
        toggle_loading(root, False)
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

def main():
    global root
    root = tk.Tk()
    root.title("OceanStor Command Generator")
    
    try:
        icon_path = os.path.join("Icons", "exe_icon.ico")
        root.iconbitmap(icon_path)
    except Exception as e:
        try:
            base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")
            icon_path = os.path.join(base_path, "Icons", "exe_icon.ico")
            root.iconbitmap(icon_path)
        except Exception:
            pass

    # Window sizing and positioning
    window_width = 550
    window_height = 420
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_x = (screen_width - window_width) // 2
    position_y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    apply_style(root)

    # Device selection
    device_var = tk.StringVar(value="OceanStor Dorado")
    script_var = tk.StringVar()
    command_var = tk.StringVar()

    # Device dropdown
    device_menu = ttk.Combobox(
        root,
        textvariable=device_var,
        values=["OceanStor Dorado", "OceanStor Pacific"],
        state="readonly",
        font=("Arial", 11),
        justify="center",
    )
    device_menu.pack(pady=10)

    # Load initial config
    def load_current_config():
        return load_config(device_var.get())

    # Create UI elements
    config = load_current_config()
    script_menu, open_button, run_button, clear_button, update_command_buttons = create_buttons_and_dropdown(
        root, open_excel, run_script, script_var, command_var, clear_excel, import_excel, device_var, config
    )

    # Update UI when device changes
    def on_device_change(*args):
        new_config = load_current_config()
        script_menu['values'] = list(new_config.keys())
        if script_var.get() not in new_config:
            script_var.set(list(new_config.keys())[0] if new_config else "")
        update_command_buttons(new_config)

    device_var.trace_add("write", on_device_change)

    # Update command buttons when script changes
    def on_script_change(*args):
        update_command_buttons(load_current_config())

    script_var.trace_add("write", on_script_change)

    root.mainloop()

if __name__ == "__main__":
    main()