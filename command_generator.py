import os
import sys
import pandas as pd
import re
from utils import load_config, read_file, get_data_file_path

def process_audit_items(audit_items_input):
    """
    Process the audit_items input to handle multiple formats:
    - value1, value2, value3, ...
    - value1,value2,value3
    - value1 value2 value3
    """
    if pd.isna(audit_items_input) or audit_items_input == '':
        return None

    # Remove any leading/trailing whitespace
    audit_items_input = audit_items_input.strip()
    
    # Replace spaces and commas with a single comma
    processed_value = re.sub(r"[\s,]+", ",", audit_items_input)
    
    return processed_value

def generate_commands(data_frame, resource_type, command_type, mandatory_fields):
    commands = []

    if command_type == "Show":
        if data_frame.empty:
            if resource_type == "FileSystem":
                commands.append(f"{command_type.lower()} file_system general")
            else:
                commands.append(f"{command_type.lower()} share {resource_type.lower()}")
            return commands

    for index, row in data_frame.iterrows():
        try:
            if command_type != "Show":
                missing_mandatory = [field for field in mandatory_fields if pd.isna(row.get(field)) or row.get(field) == '']
                if missing_mandatory:
                    print(f"Warning: Missing mandatory data in row {index + 1} of sheet '{command_type}'. Missing fields: {missing_mandatory}. Skipping.")
                    continue

            if resource_type == "FileSystem":
                command = f"{command_type.lower()} file_system general"
            else:
                command = f"{command_type.lower()} share {resource_type.lower()}"

            if command_type != "Show":
                for field in mandatory_fields:
                    value = row.get(field)
                    if field == 'local_path' and not str(value).startswith('/'):
                        command += f" {field}=/{str(value).lstrip('/')}"
                    elif field == 'audit_items':
                        processed_value = process_audit_items(value)
                        if processed_value:
                            command += f" {field}={processed_value}"
                    else:
                        command += f" {field}={value}"

            for param, value in row.items():
                if param in mandatory_fields or pd.isna(value) or value == '' or value is None:
                    continue
                if param == 'audit_items':
                    processed_value = process_audit_items(value)
                    if processed_value:
                        command += f" {param}={processed_value}"
                else:
                    command += f" {param}={value}"

            commands.append(command)
        except KeyError as e:
            print(f"Warning: Missing column '{e}' in row {index + 1} of sheet '{command_type}'. Skipping.")
            continue
        except Exception as e:
            print(f"Error processing row {index + 1}: {e}")
            continue

    return commands

def main(resource_type):
    base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")
    excel_file_path = os.path.join(base_path, 'Documents', f'{resource_type}_commands.xlsx')
    results_dir = os.path.join(base_path, 'Results')

    if not os.path.exists(results_dir):
        os.makedirs(results_dir)

    config = load_config()
    resource_columns = config.get(resource_type, {})

    if not os.path.exists(excel_file_path):
        print(f"Error: Excel file '{excel_file_path}' does not exist.")
        sys.exit(1)

    all_commands = []
    for command_type, fields in resource_columns.items():
        data_frame = read_file(excel_file_path, sheet_name=command_type)
        if data_frame is None:
            continue

        mandatory_fields = [field["name"] for field in fields.get("mandatory", [])]
        commands = generate_commands(data_frame, resource_type, command_type, mandatory_fields)
        all_commands.extend(commands)
        all_commands.append('')

    output_file_path = os.path.join(results_dir, f'{resource_type.lower()}_commands.txt')
    with open(output_file_path, 'w') as file:
        for command in all_commands:
            file.write(command + '\n')

    print(f"Commands written to: {output_file_path}")
    return output_file_path  # Return the path to the output file

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python command_generator.py <resource_type>")
        sys.exit(1)

    resource_type = sys.argv[1]
    main(resource_type)