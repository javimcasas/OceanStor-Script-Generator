import os
import sys
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

def load_config():
    """Load the configuration from the JSON file."""
    with open("commands_config.json", "r") as file:
        return json.load(file)

def get_data_file_path(filename):
    """Get the correct path to a file, whether the app is running as a script or as an executable."""
    if getattr(sys, 'frozen', False):
        # Running as an executable
        base_path = sys._MEIPASS
    else:
        # Running as a script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

def load_config():
    """Load the configuration from the JSON file."""
    config_path = get_data_file_path("commands_config.json")
    try:
        with open(config_path, "r") as file:
            return json.load(file)
    except FileNotFoundError:
        print(f"Error: Configuration file not found at {config_path}")
        raise
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON format in {config_path}")
        raise

def read_file(file_path, sheet_name):
    """Read an Excel sheet into a DataFrame."""
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading sheet '{sheet_name}' from file '{file_path}': {e}")
        return None