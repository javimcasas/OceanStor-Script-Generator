import os
import sys
import pandas as pd
import json

def load_config():
    """Load the configuration from the JSON file."""
    with open("commands_config.json", "r") as file:
        return json.load(file)

def read_file(file_path, sheet_name):
    """
    Reads an Excel sheet into a DataFrame.

    :param file_path: Path to the Excel file.
    :param sheet_name: Name of the sheet to read.
    :return: DataFrame containing the sheet data.
    """
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading sheet '{sheet_name}' from file '{file_path}': {e}")
        return None

def generate_commands(data_frame, resource_type, command_type, mandatory_fields):
    """
    Generates commands from a DataFrame based on the resource and command type.

    :param data_frame: DataFrame containing the data.
    :param resource_type: The resource type (e.g., 'CIFS', 'NFS', 'FileSystem').
    :param command_type: Type of command (e.g., 'Create', 'Change', 'Show').
    :param mandatory_fields: List of mandatory field names for the command type.
    :return: List of generated commands.
    """
    commands = []

    # Handle Show command separately
    if command_type == "Show":
        # If the sheet is empty, generate a single command without parameters
        if data_frame.empty:
            if resource_type == "FileSystem":
                commands.append(f"{command_type.lower()} file_system general")
            else:
                commands.append(f"{command_type.lower()} share {resource_type.lower()}")
            return commands

    # Process rows for all command types
    for index, row in data_frame.iterrows():
        try:
            # Check if all mandatory fields are present (except for Show)
            if command_type != "Show":
                missing_mandatory = [field for field in mandatory_fields if pd.isna(row.get(field)) or row.get(field) == '']
                if missing_mandatory:
                    print(f"Warning: Missing mandatory data in row {index + 1} of sheet '{command_type}'. Missing fields: {missing_mandatory}. Skipping.")
                    continue

            # Start building the command
            if resource_type == "FileSystem":
                command = f"{command_type.lower()} file_system general"
            else:
                command = f"{command_type.lower()} share {resource_type.lower()}"

            # Add mandatory fields (except for Show)
            if command_type != "Show":
                for field in mandatory_fields:
                    value = row.get(field)
                    if field == 'local_path' and not str(value).startswith('/'):
                        command += f" {field}=/{str(value).lstrip('/')}"  # Ensure local_path starts with '/'
                    else:
                        command += f" {field}={value}"

            # Add optional fields
            for param, value in row.items():
                if param in mandatory_fields or pd.isna(value) or value == '' or value is None:
                    continue  # Skip mandatory, empty, or None fields
                command += f" {param}={value}"

            # Add the command to the list
            commands.append(command)
        except KeyError as e:
            print(f"Warning: Missing column '{e}' in row {index + 1} of sheet '{command_type}'. Skipping.")
            continue
        except Exception as e:
            print(f"Error processing row {index + 1}: {e}")
            continue

    return commands

def main(resource_type):
    # Define paths
    base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")
    excel_file_path = os.path.join(base_path, 'Documents', f'{resource_type}_commands.xlsx')
    results_dir = os.path.join(base_path, 'Results')

    # Ensure the 'Results' folder exists
    if not os.path.exists(results_dir):
        os.makedirs(results_dir)

    # Load JSON configuration
    config = load_config()
    resource_columns = config.get(resource_type, {})

    # Check if the Excel file exists
    if not os.path.exists(excel_file_path):
        print(f"Error: Excel file '{excel_file_path}' does not exist.")
        sys.exit(1)

    # Generate commands for all sheets
    all_commands = []
    for command_type, fields in resource_columns.items():
        # Read the sheet
        data_frame = read_file(excel_file_path, sheet_name=command_type)
        if data_frame is None:
            continue

        # Extract mandatory field names from JSON
        mandatory_fields = [field["name"] for field in fields.get("mandatory", [])]

        # Generate commands for the sheet
        commands = generate_commands(data_frame, resource_type, command_type, mandatory_fields)
        all_commands.extend(commands)

        # Add a line break after commands from each sheet
        all_commands.append('')

    # Write all commands to a single file (overwrite if it exists)
    output_file_path = os.path.join(results_dir, f'{resource_type.lower()}_commands.txt')
    with open(output_file_path, 'w') as file:
        for command in all_commands:
            file.write(command + '\n')

    print(f"Commands written to: {output_file_path}")

    # Return the output file path for the GUI to handle
    return output_file_path

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python command_generator.py <resource_type>")
        sys.exit(1)

    resource_type = sys.argv[1]  # e.g., 'CIFS', 'NFS', 'FileSystem'
    main(resource_type)