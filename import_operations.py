import pandas as pd
import os
import sys
from tkinter import messagebox
from log_operations import create_import_log
from import_commands import (
    import_vstore, 
    import_filesystem, 
    import_cifs_share, 
    import_nfs_share, 
    import_cifs_permission, 
    import_nfs_permission
)


TARGET_SHEETS = [
    'Vstore', 
    'Filesystem', 
    'CIFS_Share', 
    'NFS_Share',
    'CIFS_Share_Permission'
]

def process_imported_data(file_path):
    """Process the imported Excel data by handling each target sheet separately."""
    results_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Imported_Results")
    status = "Success"
    
    try:
        if not os.path.exists(results_dir):
            os.makedirs(results_dir)

        xls = pd.ExcelFile(file_path)
        
        # First process vstores to get their IDs and names
        vstores = {}
        if 'Vstore' in xls.sheet_names:
            vstore_df = xls.parse('Vstore').dropna(how='all')
            for _, row in vstore_df.iterrows():
                if 'Vstore ID' in row and 'Vstore' in row:
                    vstores[row['Vstore']] = row['Vstore ID']
        
        for sheet_name in TARGET_SHEETS:
            if sheet_name in xls.sheet_names:
                try:
                    df = xls.parse(sheet_name).dropna(how='all')
                    if df.empty:
                        continue
                        
                    process_sheet_rows(sheet_name, df, results_dir, vstores)
                    
                except Exception as e:
                    print(f"Error processing {sheet_name}: {str(e)}")
                    status = "Failed"
        
        # Create log file after processing
        log_file = create_import_log(file_path, results_dir, status)
        if log_file:
            print(f"Log file created: {log_file}")
            
    except Exception as e:
        status = "Failed"
        messagebox.showerror("Error", f"Error processing file: {str(e)}")
        # Try to create log even if processing failed
        create_import_log(file_path, results_dir, status)
        raise

def process_sheet_rows(sheet_name, df, output_dir, vstores):
    """Process each row in a sheet and generate commands with vstore context."""
    if sheet_name == 'NFS_Share':
        # Create two separate files for NFS shares and permissions
        share_file = os.path.join(output_dir, "NFS_Share_commands.txt")
        perm_file = os.path.join(output_dir, "NFS_Share_Permission_commands.txt")
        
        current_vstore = None
        
        with open(share_file, 'w') as sf, open(perm_file, 'w') as pf:
            for index, row in df.iterrows():
                try:
                    # Get vstore info for this row
                    row_vstore = None
                    if 'Vstore' in row and not pd.isna(row['Vstore']):
                        row_vstore = row['Vstore']
                    elif 'Vstore ID' in row and not pd.isna(row['Vstore ID']):
                        row_vstore = next((name for name, id_ in vstores.items() if id_ == row['Vstore ID']), None)
                    
                    # Add vstore change command if needed
                    if row_vstore and row_vstore != current_vstore:
                        if current_vstore is not None:
                            sf.write("\n")
                            pf.write("\n")
                        vstore_cmd = f"change vstore view name={row_vstore}\n"
                        sf.write(vstore_cmd)
                        pf.write(vstore_cmd)
                        current_vstore = row_vstore
                    
                    # Generate share command
                    share_cmd = import_nfs_share.generate_nfs_share_command(row)
                    if share_cmd:
                        sf.write(f"{share_cmd}\n")
                    
                    # Generate permission command if permission data exists
                    perm_cmd = import_nfs_permission.generate_nfs_permission_command(row)
                    if perm_cmd:
                        pf.write(f"{perm_cmd}\n")
                        
                except Exception as e:
                    print(f"Error in row {index} of {sheet_name}: {str(e)}")
    else:
        # Original processing for other sheet types
        output_file = os.path.join(output_dir, f"{sheet_name}_commands.txt")
        current_vstore = None
        
        with open(output_file, 'w') as f:
            for index, row in df.iterrows():
                try:
                    # Get vstore info for this row
                    row_vstore = None
                    if 'Vstore' in row and not pd.isna(row['Vstore']):
                        row_vstore = row['Vstore']
                    elif 'Vstore ID' in row and not pd.isna(row['Vstore ID']):
                        row_vstore = next((name for name, id_ in vstores.items() if id_ == row['Vstore ID']), None)
                    
                    # Add vstore change command if needed
                    if sheet_name != 'Vstore' and row_vstore and row_vstore != current_vstore:
                        if current_vstore is not None:
                            f.write("\n")
                        f.write(f"change vstore view name={row_vstore}\n")
                        current_vstore = row_vstore
                    
                    cmd = generate_command(sheet_name, row)
                    if cmd:
                        f.write(f"{cmd}\n")
                        
                except Exception as e:
                    print(f"Error in row {index} of {sheet_name}: {str(e)}")

def generate_command(sheet_name, row):
    """Generate specific command based on sheet type and row data."""
    row = row.dropna()  # Remove empty values
    
    if sheet_name == 'Vstore':
        return import_vstore.generate_vstore_command(row)
    elif sheet_name == 'Filesystem':
        return import_filesystem.generate_filesystem_command(row)
    elif sheet_name == 'CIFS_Share':
        return import_cifs_share.generate_cifs_share_command(row)
    elif sheet_name == 'NFS_Share':
        return import_nfs_share.generate_nfs_share_command(row)
    elif sheet_name == 'CIFS_Share_Permission':
        return import_cifs_permission.generate_cifs_permission_command(row)
    return None