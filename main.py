import os, sys
import tkinter as tk
from tkinter import messagebox, ttk
from excel_operations import open_excel, clear_excel, import_excel
from gui_helpers import apply_style, create_buttons_and_dropdown
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

        # Call the command generator logic directly
        output_file_path = generate_commands(script_type, device_type)

        if os.path.exists(output_file_path):
            messagebox.showinfo("Success", f"Script executed successfully. Output saved to:\n{output_file_path}")
            os.startfile(output_file_path)  # Open the output file
        else:
            messagebox.showerror("Error", f"The file '{output_file_path}' does not exist.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def main():
    root = tk.Tk()
    root.title("Script Selector")
    
    try:
        icon_path = os.path.join("Icons", "exe_icon.ico")
        root.iconbitmap(icon_path)
    except Exception as e:
        try:
            base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")
            icon_path = os.path.join(base_path, "Icons", "exe_icon.ico")
            root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not load window icon: {str(e)}")

    window_width = 550
    window_height = 450
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_x = (screen_width - window_width) // 2
    position_y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    apply_style(root)

    # Add a new dropdown for device selection (OceanStor Dorado or OceanStor Pacific)
    device_var = tk.StringVar(value="OceanStor Dorado")  # Default to OceanStor Dorado
    script_var = tk.StringVar(value="CIFS")  # Default script type
    command_var = tk.StringVar(value="Create")  # Default command type

    # Device selection dropdown
    device_menu = ttk.Combobox(
        root,
        textvariable=device_var,
        values=["OceanStor Dorado", "OceanStor Pacific"],  # Options for device selection
        state="readonly",
        font=("Arial", 11),
        justify="center",
    )
    device_menu.pack(pady=10)

    # Load config based on the selected device
    def load_selected_config():
        selected_device = device_var.get()
        return load_config(selected_device)

    # Create buttons and dropdowns for script and command selection
    config = load_selected_config()
    script_menu, open_button, run_button, clear_button, update_command_buttons = create_buttons_and_dropdown(
        root, open_excel, run_script, script_var, command_var, clear_excel, import_excel, device_var, config
    )

    # Update the script dropdown and buttons when the device changes
    def update_script_dropdown(*args):
        selected_device = device_var.get()
        print(f"Selected device: {selected_device}")  # Debug print
        new_config = load_config(selected_device)  

        if script_var.get() not in new_config:
            script_var.set(list(new_config.keys())[0])  # Reset only if needed
        
        script_menu['values'] = list(new_config.keys())
        update_command_buttons(new_config)

    device_var.trace_add("write", update_script_dropdown)

    # Update the buttons when the script type changes
    def update_buttons(*args):
        selected_script = script_var.get()
        selected_device = device_var.get()
        new_config = load_config(selected_device)

        if selected_script in new_config:
            update_command_buttons(new_config)

    script_var.trace_add("write", update_buttons)

    root.mainloop()

if __name__ == "__main__":
    main()