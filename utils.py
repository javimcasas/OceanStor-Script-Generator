import os
import sys
import json
import zipfile
from datetime import datetime
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
        
def encapsulate_results(results_dir="Imported_Results"):
    """Encapsulate all .txt files in the results directory into a zip file and clean up."""
    try:
        # Create timestamp for zip filename
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        zip_filename = f"import_commands_{timestamp}.zip"
        zip_path = os.path.join(results_dir, zip_filename)
        
        # Get all .txt files in the results directory
        txt_files = [f for f in os.listdir(results_dir) 
                   if f.endswith('.txt') and os.path.isfile(os.path.join(results_dir, f))]
        
        # Find the latest log file
        logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logs")
        log_files = []
        if os.path.exists(logs_dir):
            log_files = [f for f in os.listdir(logs_dir) 
                       if f.startswith('import_log_') and f.endswith('.txt')]
        
        # Sort log files by creation time (newest first)
        if log_files:
            log_files.sort(key=lambda x: os.path.getmtime(os.path.join(logs_dir, x)), reverse=True)
            latest_log = log_files[0]
        else:
            latest_log = None
        
        # Create the zip file
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add all command files
            for txt_file in txt_files:
                file_path = os.path.join(results_dir, txt_file)
                zipf.write(file_path, os.path.basename(file_path))
            
            # Add the latest log file if found
            if latest_log:
                log_path = os.path.join(logs_dir, latest_log)
                zipf.write(log_path, f"logs/{latest_log}")
        
        # Delete original .txt files (only if zip was created successfully)
        for txt_file in txt_files:
            try:
                os.remove(os.path.join(results_dir, txt_file))
            except Exception as e:
                print(f"Warning: Could not delete {txt_file}: {str(e)}")
        
        return zip_path
    except Exception as e:
        print(f"Error creating zip file: {str(e)}")
        return None