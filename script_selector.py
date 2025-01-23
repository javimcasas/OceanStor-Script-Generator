import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import subprocess
import os
import sys
from styled_gui import apply_style, create_buttons_and_dropdown
from excel_manager import create_excel_if_not_exists

def get_data_file_path(filename):
    """Devuelve la ruta correcta al archivo, dependiendo de si está empaquetado o no."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, filename)

def run_script(script_type):
    try:
        if script_type == 'CIFS':
            script_path = get_data_file_path('cifs_share_script.py')
            output_file_path = 'Results\\cifs_shares_commands.txt'
        elif script_type == 'NFS':
            script_path = get_data_file_path('nfs_share_script.py')
            output_file_path = 'Results\\nfs_shares_commands.txt'
        else:
            raise ValueError("Invalid script type")

        result = subprocess.run(['python', script_path], capture_output=True, text=True)
        if result.returncode == 0:
            messagebox.showinfo("Success", f"Script executed successfully:\n{result.stdout}")
            if os.path.exists(output_file_path):
                os.startfile(output_file_path)
            else:
                messagebox.showerror("Error", f"The file '{output_file_path}' does not exist.")
        else:
            messagebox.showerror("Error", f"Error executing the script:\n{result.stderr}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_excel():
    try:
        script_type = script_var.get()
        if script_type == 'CIFS':
            excel_path = 'Documents\\CIFSShares.xlsx'
        elif script_type == 'NFS':
            excel_path = 'Documents\\NFSShares.xlsx'
        else:
            raise ValueError("Invalid script type")

        create_excel_if_not_exists(script_type, excel_path)
        if os.path.exists(excel_path):
            os.startfile(excel_path)
        else:
            messagebox.showerror("Error", f"The file '{excel_path}' does not exist.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while trying to open the Excel file: {e}")

root = tk.Tk()
root.title("Script Selector")

# Configuración de tamaño y posición de la ventana
window_width = 550
window_height = 350
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_x = (screen_width - window_width) // 2
position_y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

# Aplicar estilo desde styled_gui.py
apply_style(root)

# Crear los elementos de la interfaz
script_var = tk.StringVar(root)
script_var.set("CIFS")  # Valor predeterminado

script_menu, open_button, run_button = create_buttons_and_dropdown(
    root, open_excel, run_script, script_var
)

# Ejecutar la interfaz
root.mainloop()
