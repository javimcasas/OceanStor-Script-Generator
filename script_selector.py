import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import subprocess
import os
import sys
from openpyxl import load_workbook
from styled_gui import apply_style, create_buttons_and_dropdown
from excel_manager import create_excel_for_resource

def get_data_file_path(filename):
    """Devuelve la ruta correcta al archivo, dependiendo de si está empaquetado o no."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, filename)

def run_script(script_type, command_type):
    try:
        # Define the script path and output file path based on the selected command
        if script_type == 'CIFS':
            script_path = get_data_file_path('cifs_share_script.py')
            output_file_path = os.path.join('Results', 'cifs_shares_commands.txt')
        elif script_type == 'NFS':
            script_path = get_data_file_path('nfs_share_script.py')
            output_file_path = os.path.join('Results', 'nfs_shares_commands.txt')
        else:
            raise ValueError("Invalid script type")

        # Run the selected script
        result = subprocess.run(['python', script_path, command_type], capture_output=True, text=True)

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
        # Define the path to the Documents folder
        documents_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Documents")

        # Ensure the Documents folder exists
        if not os.path.exists(documents_folder):
            os.makedirs(documents_folder)

        # Define the Excel file path based on the script type
        if script_type == 'CIFS':
            excel_path = os.path.join(documents_folder, 'CIFS_commands.xlsx')
        elif script_type == 'NFS':
            excel_path = os.path.join(documents_folder, 'NFS_commands.xlsx')
        else:
            raise ValueError("Invalid script type")

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

# Crear la ventana principal
root = tk.Tk()
root.title("Script Selector")

# Tamaño de la ventana
window_width = 500
window_height = 400
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_x = (screen_width - window_width) // 2
position_y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

# Aplicar estilo a la ventana
apply_style(root)

# Variables para los selectores
script_var = tk.StringVar(value="CIFS")
command_var = tk.StringVar(value="Create")  # Por defecto, "Create"

# Crear el selector y botones en la interfaz
script_menu, open_button, run_button, command_buttons = create_buttons_and_dropdown(
    root, open_excel, run_script, script_var, command_var
)

# Ejecutar la interfaz
root.mainloop()