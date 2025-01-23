import os
import sys
import pandas as pd
from readData import readFile

def main():
    # Este es el lugar donde ejecutas tu código
    print("Script CIFS/NFS ejecutado correctamente.")

    # Define the 'Results' folder path
    results_dir = 'Results'

    # Use sys._MEIPASS to get the correct path if running in a packaged executable
    if getattr(sys, 'frozen', False):  # Check if we are in a PyInstaller bundle
        base_path = sys._MEIPASS  # Use the temporary folder created by PyInstaller
    else:
        base_path = os.path.abspath(".")  # Use the current working directory

    # Ahora se accede correctamente a los archivos dentro de los subdirectorios
    cifs_file_path = os.path.join(base_path, 'Documents', 'CifsShares.xlsx')
    data_file = readFile(cifs_file_path)

    # Ensure the data_file is a DataFrame
    if data_file is None:
        sys.exit("Execution Stopped: Error reading data_file")

    # If readFile returns a list, convert it to a DataFrame
    if isinstance(data_file, list):
        data_file = pd.DataFrame(data_file, columns=[ 
            'name', 'local_path', 'continue_available_enabled', 'oplock_enabled', 'notify_enabled', 
            'file_system_id', 'file_system_name', 'offline_file_mode', 'smb2_ca_enable', 'ip_control_enabled',
            'abe_enabled', 'audit_items', 'file_filter_enable', 'apply_default_acl', 'share_type',
            'show_previous_versions_enabled', 'show_snapshot_enabled', 'browse_enable', 'readdir_timeout',
            'lease_enable', 'dir_umask', 'file_umask'
        ])

    # Check if the 'Results' folder exists, if not, create it
    if not os.path.exists(results_dir):
        os.makedirs(results_dir)

    # Full path of the results file
    file_path = os.path.join(results_dir, 'cifs_shares_commands.txt')

    # Create and write to the results file
    with open(file_path, 'w') as archive:
        # Iterate through each row of data in the DataFrame
        for index, row in data_file.iterrows():
            # Check if the first two columns (share name and local path) have data
            if pd.isna(row['name']) or row['name'] == '' or pd.isna(row['local_path']) or row['local_path'] == '':
                print(f"Error: Missing mandatory data in row {index + 1}. 'name' or 'local_path' is empty.")
                sys.exit("Execution Stopped: Missing mandatory data.")

            # Validate if 'local_path' starts with '/' and add it if missing
            local_path = row['local_path']
            if not local_path.startswith('/'):
                local_path = '/' + local_path  # Add '/' at the beginning if it's missing

            # Start building the base command with the first two columns (mandatory)
            command = f"create share cifs name={row['name']} local_path={local_path}"

            # Iterate over each column (from 2 onward) and add to command if it's not empty or NaN
            for param, value in row.items():
                if pd.isna(value) or value == '' or value is None:
                    continue  # Skip if the value is NaN or empty

                # Skip the first two columns ('name' and 'local_path')
                if param in ['name', 'local_path']:
                    continue

                # Add the parameter to the command
                command += f" {param}={value}"

            # Write the command to the file
            archive.write(command + '\n')

    print(f"Commands written to: {file_path}")

if __name__ == "__main__":
    main()
