import os
import pandas as pd

def create_excel_if_not_exists(script_type, file_path):
    """
    Checks if the given Excel file exists. If it doesn't, creates the file
    with the appropriate columns for the selected script type.
    
    :param script_type: The type of script ('CIFS' or 'NFS').
    :param file_path: The path to the Excel file.
    """
    if not os.path.exists(file_path):
        # Define the column names based on the script type
        if script_type == 'CIFS':
            columns = [
                'name', 'local_path', 'continue_available_enabled', 'oplock_enabled', 
                'nofiy_enabled', 'file_system_id', 'file_system_name', 'offline_file_mode', 
                'smb2_ca_enable', 'ip_control_enabled', 'abe_enabled', 'audit_items', 
                'file_filter_enable', 'apply_default_acl', 'share_type', 
                'show_previous_versions_enabled', 'show_snapshot_enabled', 'browse_enable', 
                'readdir_timeout', 'lease_enable', 'dir_unmask', 'file_unmask'
            ]
        elif script_type == 'NFS':
            columns = [
                'local_path', 'file_system_id', 'file_system_name', 'charset', 
                'lock_type', 'alias', 'audit_items', 'show_snapshot_enabled', 
                'description'
            ]
        else:
            raise ValueError("Invalid script type")

        # Create a DataFrame with the specified columns and no data
        df = pd.DataFrame(columns=columns)

        # Save the DataFrame to an Excel file
        os.makedirs(os.path.dirname(file_path), exist_ok=True)  # Ensure the directory exists
        df.to_excel(file_path, index=False)
        print(f"Excel file created: {file_path}")
    else:
        print(f"Excel file already exists: {file_path}")
