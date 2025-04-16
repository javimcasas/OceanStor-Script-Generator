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

    # Adding logic for 'create namespace' command
    if resource_type == "Namespace" and command_type == "Create":
        for index, row in data_frame.iterrows():
            try:
                # Ensure all mandatory fields for 'create namespace' are present
                missing_mandatory = [field for field in mandatory_fields if pd.isna(row.get(field)) or row.get(field) == '']
                if missing_mandatory:
                    print(f"Warning: Missing mandatory data in row {index + 1} of sheet '{command_type}'. Missing fields: {missing_mandatory}. Skipping.")
                    continue

                # Start building the 'create namespace' command
                command = f"create namespace general"

                # Loop through each mandatory field and add to command
                for field in mandatory_fields:
                    value = row.get(field)
                    if field == 'name':
                        command += f" name={value}"
                    elif field == 'storage_pool_id':
                        command += f" storage_pool_id={value}"
                    elif field == 'account_id':
                        command += f" account_id={value}"
                    elif field == 'dir_split_policy' or field == 'dir_split_bitwidth' or field == 'root_split_bitwidth' or field == 'is_trash_enable' or field == 'is_trash_visible' or field == 'interval_trash' or field == 'checkpoint_trash' or field == 'recycle_op_permission' or field == 'enable_encrypt' or field == 'atime_update_mode' or field == 'rdc' or field == 'is_show_snap_dir' or field == 'dentry_table_type' or field == 'acl_policy_type' or field == 'crypt_alg' or field == 'forbidden_dpc' or field == 'concat_enable' or field == 'obs_hide_dir' or field == 'io_block_size' or field == 'unix_permissions' or field == 'case_sensitive' or field == 'aggre_policy_type' or field == 'dir_frag_policy' or field == 'attr_enhance_state' or field == 'access_is_owner' or field == 'std_compress_front_switch' or field == 'std_compress_tier_switch' or field == 'std_compress_dpc_switch' or field == 'std_compress_front_hot_pool_switch' or field == 'std_compress_tier_hot_pool_switch' or field == 'std_compress_front_mode' or field == 'std_compress_tier_mode' or field == 'snapshot_difference_switch' or field == 'glacier_read_switch' or field == 'glacier_read_rsv_time' or field == 'glacier_read_trig_restore' or field == 'application_type' or field == 'dtree_mig_switch' or field == 'tier_hot_cap_limit' or field == 'tier_warm_cap_limit' or field == 'tier_cold_cap_limit':
                        # Optional fields
                        if value:  # Only add if the value is not empty
                            command += f" {field}={value}"

                commands.append(command)
            except KeyError as e:
                print(f"Warning: Missing column '{e}' in row {index + 1} of sheet '{command_type}'. Skipping.")
                continue
            except Exception as e:
                print(f"Error processing row {index + 1}: {e}")
                continue

    else:
        # Existing logic for other resource types
        for index, row in data_frame.iterrows():
            try:
                if command_type != "Show":
                    missing_mandatory = [field for field in mandatory_fields if pd.isna(row.get(field)) or row.get(field) == '']
                    if missing_mandatory:
                        print(f"Warning: Missing mandatory data in row {index + 1} of sheet '{command_type}'. Missing fields: {missing_mandatory}. Skipping.")
                        continue

                if resource_type == "FileSystem":
                    command = f"{command_type.lower()} file_system general"
                elif resource_type == "Quota":
                    if command_type == "FS Quota":
                        command = "create quota file_system"
                    elif command_type == "Dtree Quota":
                        command = "create quota dtree"
                    elif command_type == "Change":
                        command = "change quota general"
                    elif command_type == "Show":
                        command = "show quota general"
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
                        elif field == 'file_system_name' or field == 'file_system_id':
                            command += f" {field}={value}"
                        elif field == 'quota_type':
                            command += f" {field}={value}"
                        elif field == 'space_hard_quota' or field == 'space_soft_quota' or field == 'file_hard_quota' or field == 'file_soft_quota':
                            command += f" {field}={value}"
                        elif field == 'quota_id':
                            command += f" {field}={value}"
                        else:
                            command += f" {field}={value}"

                for param, value in row.items():
                    if param in mandatory_fields or pd.isna(value) or value == '' or value is None:
                        continue
                    if param == 'audit_items':
                        processed_value = process_audit_items(value)
                        if processed_value:
                            command += f" {param}={processed_value}"
                    elif param == 'user_group_type' or param == 'user_name' or param == 'group_name' or param == 'domain_type':
                        command += f" {param}={value}"
                    elif param == 'space_hard_quota' or param == 'space_soft_quota' or param == 'file_hard_quota' or param == 'file_soft_quota':
                        command += f" {param}={value}"
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

def main(resource_type, device_type):
    base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")
    excel_file_path = os.path.join(base_path, 'Documents', f'{device_type}_{resource_type}_commands.xlsx')
    results_dir = os.path.join(base_path, 'Results')

    if not os.path.exists(results_dir):
        os.makedirs(results_dir)

    config = load_config(device_type)
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

    output_file_path = os.path.join(results_dir, f'{device_type.lower()}_{resource_type.lower()}_commands.txt')
    with open(output_file_path, 'w') as f:
        for command in all_commands:
            f.write(command + '\n')

    print(f"Commands written to: {output_file_path}")
    return output_file_path  # Return the path to the output file

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python command_generator.py <resource_type>")
        sys.exit(1)

    resource_type = sys.argv[1]
    main(resource_type)
