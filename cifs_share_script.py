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

def generate_commands(data_frame, command_type):
    """
    Generates commands from a DataFrame based on the command type.

    :param data_frame: DataFrame containing the data.
    :param command_type: Type of command (e.g., 'Create', 'Change', 'Show').
    :return: List of generated commands.
    """
    commands = []
    for index, row in data_frame.iterrows():
        try:
            # Skip rows where mandatory fields are missing
            if command_type == 'Create' and (pd.isna(row['name']) or row['name'] == '' or pd.isna(row['local_path']) or row['local_path'] == ''):
                print(f"Warning: Missing mandatory data in row {index + 1} of sheet '{command_type}'. Skipping.")
                continue
            elif command_type == 'Change' and (pd.isna(row['share_id']) or row['share_id'] == '') and (pd.isna(row['share_name']) or row['share_name'] == ''):
                print(f"Warning: Missing mandatory data in row {index + 1} of sheet '{command_type}'. Skipping.")
                continue
            elif command_type == 'Show' and (pd.isna(row['share_id']) or row['share_id'] == '') and (pd.isna(row['file_system_id']) or row['file_system_id'] == '') and (pd.isna(row['share_name']) or row['share_name'] == ''):
                print(f"Warning: Missing mandatory data in row {index + 1} of sheet '{command_type}'. Skipping.")
                continue

            # Start building the command
            if command_type == 'Create':
                # Validate and format 'local_path'
                local_path = row['local_path']
                if not local_path.startswith('/'):
                    local_path = '/' + local_path

                command = f"{command_type.lower()} share cifs name={row['name']} local_path={local_path}"

                # Add additional parameters
                for param, value in row.items():
                    if pd.isna(value) or value == '' or value is None or param in ['name', 'local_path']:
                        continue  # Skip empty or mandatory fields
                    command += f" {param}={value}"

            elif command_type == 'Change':
                # Mandatory fields: share_id or share_name
                if not pd.isna(row['share_id']) and row['share_id'] != '':
                    command = f"{command_type.lower()} share cifs share_id={row['share_id']}"
                elif not pd.isna(row['share_name']) and row['share_name'] != '':
                    command = f"{command_type.lower()} share cifs share_name={row['share_name']}"
                else:
                    continue  # Skip if both are missing

                # Add additional parameters
                for param, value in row.items():
                    if pd.isna(value) or value == '' or value is None or param in ['share_id', 'share_name']:
                        continue  # Skip empty or mandatory fields
                    command += f" {param}={value}"

            elif command_type == 'Show':
                # Mandatory fields: share_id, file_system_id, or share_name
                if not pd.isna(row['share_id']) and row['share_id'] != '':
                    command = f"{command_type.lower()} share cifs share_id={row['share_id']}"
                elif not pd.isna(row['file_system_id']) and row['file_system_id'] != '':
                    command = f"{command_type.lower()} share cifs file_system_id={row['file_system_id']}"
                elif not pd.isna(row['share_name']) and row['share_name'] != '':
                    command = f"{command_type.lower()} share cifs share_name={row['share_name']}"
                else:
                    continue  # Skip if all are missing

                # Add additional parameters
                for param, value in row.items():
                    if pd.isna(value) or value == '' or value is None or param in ['share_id', 'file_system_id', 'share_name']:
                        continue  # Skip empty or mandatory fields
                    command += f" {param}={value}"

            commands.append(command)
        except KeyError as e:
            print(f"Warning: Missing column '{e}' in row {index + 1} of sheet '{command_type}'. Skipping.")
            continue

    return commands

def main():
    # Define paths
    base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")
    cifs_file_path = os.path.join(base_path, 'Documents', 'CIFS_commands.xlsx')
    results_dir = os.path.join(base_path, 'Results')

    # Ensure the 'Results' folder exists
    if not os.path.exists(results_dir):
        os.makedirs(results_dir)

    # Load JSON configuration
    config = load_config()
    cifs_columns = config.get('CIFS', {})

    # Check if the Excel file exists
    if not os.path.exists(cifs_file_path):
        print(f"Error: Excel file '{cifs_file_path}' does not exist.")
        sys.exit(1)

    # Generate commands for each sheet
    all_commands = []
    for command_type in cifs_columns:
        # Read the sheet
        data_frame = read_file(cifs_file_path, sheet_name=command_type)
        if data_frame is None:
            continue

        # Generate commands for the sheet
        commands = generate_commands(data_frame, command_type)
        all_commands.extend(commands)

        # Add a line break after commands from each sheet
        all_commands.append('')

    # Write all commands to a single file (overwrite if it exists)
    output_file_path = os.path.join(results_dir, 'cifs_shares_commands.txt')
    with open(output_file_path, 'w') as file:
        for command in all_commands:
            file.write(command + '\n')

    print(f"Commands written to: {output_file_path}")

    # Return the output file path for the GUI to handle
    return output_file_path

if __name__ == "__main__":
    main()