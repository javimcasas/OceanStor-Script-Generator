import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import subprocess
import os
import sys
import json
from openpyxl import load_workbook
from styled_gui import apply_style, create_buttons_and_dropdown
from excel_manager import create_excel_for_resource

def load_config():
    """Load the configuration from the JSON file."""
    with open("commands_config.json", "r") as file:
        return json.load(file)

def get_data_file_path(filename):
    """Returns the correct path to the file, depending on whether it's packaged or not."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, filename)

def run_script(script_type, command_type):
    try:
        # Load configuration to validate script type
        config = load_config()
        if script_type not in config:
            raise ValueError(f"Invalid script type: {script_type}. Valid types are: {list(config.keys())}")

        # Define the script path and output file path
        script_path = get_data_file_path('command_generator.py')
        output_file_path = os.path.join('Results', f'{script_type.lower()}_commands.txt')

        # Run the command_generator script with the selected resource type
        result = subprocess.run(['python', script_path, script_type], capture_output=True, text=True)

        if result.returncode == 0:
            # Show success message
            messagebox.showinfo("Success", f"Script executed successfully:\n{result.stdout}")

            # Open the output file after the user acknowledges the message
            if os.path.exists(output_file_path):
                os.startfile(output_file_path)
            else:
                messagebox.showerror("Error", f"The file '{output_file_path}' does not exist.")
        else:
            messagebox.showerror("Error", f"Error executing the script:\n{result.stderr}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_excel_with_sheet(excel_path, sheet_name):
    """
    Opens the Excel file and activates the specified sheet.

    :param excel_path: The path to the Excel file.
    :param sheet_name: The name of the sheet to activate.
    """
    try:
        # Load the workbook
        workbook = load_workbook(excel_path)

        # Check if the sheet exists
        if sheet_name in workbook.sheetnames:
            # Activate the sheet
            workbook.active = workbook[sheet_name]
            workbook.save(excel_path)  # Save to ensure the active sheet is updated

        # Open the Excel file
        os.startfile(excel_path)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while opening the Excel sheet: {e}")

def open_excel(script_type, command_type):
    try:
        # Load configuration to validate script type
        config = load_config()
        if script_type not in config:
            raise ValueError(f"Invalid script type: {script_type}. Valid types are: {list(config.keys())}")

        # Define the path to the Documents folder
        documents_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Documents")

        # Ensure the Documents folder exists
        if not os.path.exists(documents_folder):
            os.makedirs(documents_folder)

        # Define the Excel file path based on the script type
        excel_path = os.path.join(documents_folder, f'{script_type}_commands.xlsx')

        # Create the Excel file only if it doesn't exist
        if not os.path.exists(excel_path):
            create_excel_for_resource(script_type, excel_path, documents_folder)

        # Open the Excel file and activate the selected sheet
        if os.path.exists(excel_path):
            open_excel_with_sheet(excel_path, command_type)
        else:
            messagebox.showerror("Error", f"The file '{excel_path}' does not exist.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while trying to open the Excel file: {e}")

def clear_excel(script_type):
    """
    Deletes the Excel file corresponding to the selected script type.

    :param script_type: The selected script type (e.g., CIFS, NFS, FileSystem).
    """
    try:
        # Load configuration to validate script type
        config = load_config()
        if script_type not in config:
            raise ValueError(f"Invalid script type: {script_type}. Valid types are: {list(config.keys())}")

        # Define the path to the Documents folder
        documents_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Documents")

        # Define the Excel file path based on the script type
        excel_path = os.path.join(documents_folder, f'{script_type}_commands.xlsx')

        # Check if the file exists
        if os.path.exists(excel_path):
            # Ask for confirmation before deleting
            confirm = messagebox.askyesno(
                "Confirm Deletion",
                f"Are you sure you want to delete the Excel file for {script_type}?",
            )
            if confirm:
                os.remove(excel_path)
                messagebox.showinfo("Success", f"The Excel file for {script_type} has been deleted.")
        else:
            messagebox.showinfo("Info", f"No Excel file found for {script_type}.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while deleting the Excel file: {e}")

# Create the main window
root = tk.Tk()
root.title("Script Selector")

# Set window size and position
window_width = 500
window_height = 400
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_x = (screen_width - window_width) // 2
position_y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

# Apply style to the window
apply_style(root)

# Variables for selectors
script_var = tk.StringVar(value="CIFS")  # Default to CIFS
command_var = tk.StringVar(value="Create")  # Default to Create

# Load JSON configuration
config = load_config()

# Create the selector and buttons in the interface
script_menu, open_button, run_button, clear_button = create_buttons_and_dropdown(
    root, open_excel, run_script, script_var, command_var, clear_excel, config
)

# Run the interface
root.mainloop()