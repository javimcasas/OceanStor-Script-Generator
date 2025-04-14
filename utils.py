import os
import sys
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from tkinter import messagebox

def load_config(device_type):
    config_file = f"{device_type.lower().replace(' ', '_')}_commands.json"
    with open(config_file, 'r') as file:
        config = json.load(file)
    return config

def get_data_file_path(filename):
    """Get the correct path to a file, whether the app is running as a script or as an executable."""
    if getattr(sys, 'frozen', False):
        # Running as an executable
        base_path = sys._MEIPASS
    else:
        # Running as a script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

def read_file(file_path, sheet_name):
    """Read an Excel sheet into a DataFrame."""
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading sheet '{sheet_name}' from file '{file_path}': {e}")
        return None
    
def open_directory(directory):
    import os
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), directory)
    if os.path.exists(path):
        os.startfile(path)
    else:
        messagebox.showinfo("Info", f"Directory not found: {path}")