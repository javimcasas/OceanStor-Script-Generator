import os
import pandas as pd
import json

def load_config():
    """Load the configuration from the JSON file."""
    with open("commands_config.json", "r") as file:
        return json.load(file)

def create_excel_for_resource(resource_type, excel_path, output_folder):
    """
    Creates an Excel file for a specific resource type (e.g., CIFS or NFS) only if it doesn't exist.
    Each sheet in the Excel file corresponds to a command type (Create, Modify, Show).

    :param resource_type: The resource type (e.g., "CIFS" or "NFS").
    :param excel_path: The full path to the Excel file.
    :param output_folder: The folder where the Excel file will be saved.
    """
    if not os.path.exists(excel_path):
        # Load configuration from the JSON file
        config = load_config()

        # Create an ExcelWriter object to write multiple sheets
        with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
            # Create sheets for each command type (Create, Modify, Show)
            for command_type in config[resource_type]:
                # Create an empty DataFrame for each command
                df = pd.DataFrame(columns=config[resource_type][command_type])
                sheet_name = command_type  # Use the command type as the sheet name
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Excel file created for {resource_type}: {excel_path}")
    else:
        print(f"Excel file already exists: {excel_path}")