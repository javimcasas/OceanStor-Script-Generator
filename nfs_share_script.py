import os, sys
import pandas as pd
from readData import readFile

def main():
    # Aquí va el código principal que debe ejecutarse al correr este script
    print("Script CIFS/NFS ejecutado correctamente.")

if __name__ == "__main__":
    main()

# Define the 'Results' folder path
results_dir = 'Results'

# Define XLSX document to read from
data_file = readFile('Documents\\NfsShares.xlsx')

# Ensure the data_file is a DataFrame
if data_file is None:
    sys.exit("Execution Stopped: Error reading data_file")

# If readFile returns a list, convert it to a DataFrame
if isinstance(data_file, list):
    data_file = pd.DataFrame(data_file, columns=[
        'local_path', 'file_system_id', 'file_system_name', 'charset', 'lock_type', 'alias', 'audit_items', 
        'show_snapshot_enabled', 'description'
    ])

# Check if the 'Results' folder exists, if not, create it
if not os.path.exists(results_dir):
    os.makedirs(results_dir)

# Full path of the results file
file_path = os.path.join(results_dir, 'nfs_shares_commands.txt')

# Create and write to the results file
with open(file_path, 'w') as archive:
    # Iterate through each row of data in the DataFrame
    for index, row in data_file.iterrows():
        # Check if the 'local_path' column has data
        if pd.isna(row['local_path']) or row['local_path'] == '':
            print(f"Error: Missing mandatory data in row {index + 1}. 'local_path' is empty.")
            sys.exit("Execution Stopped: Missing mandatory data.")
        
        # Validate if 'local_path' starts with '/' and add it if missing
        local_path = row['local_path']
        if not local_path.startswith('/'):
            local_path = '/' + local_path  # Add '/' at the beginning if it's missing
        
        # Start building the base command with the mandatory 'local_path' column
        command = f"create share nfs local_path={local_path}"
        
        # Iterate over each column (from 1 onward) and add to command if it's not empty or NaN
        for param, value in row.items():
            if pd.isna(value) or value == '' or value is None:
                continue  # Skip if the value is NaN or empty

            # Skip the first column ('local_path')
            if param == 'local_path':
                continue

            # Add the parameter to the command
            command += f" {param}={value}"

        # Write the command to the file
        archive.write(command + '\n')

print(f"Commands written to: {file_path}")
