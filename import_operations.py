import pandas as pd
import os
from tkinter import messagebox
from log_operations import create_import_log

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
                    share_cmd = generate_nfs_share_command(row)
                    if share_cmd:
                        sf.write(f"{share_cmd}\n")
                    
                    # Generate permission command if permission data exists
                    perm_cmd = generate_nfs_permission_command(row)
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
        return generate_vstore_command(row)
    elif sheet_name == 'Filesystem':
        return generate_filesystem_command(row)
    elif sheet_name == 'CIFS_Share':
        return generate_cifs_share_command(row)
    elif sheet_name == 'NFS_Share':
        return generate_nfs_share_command(row)
    elif sheet_name == 'CIFS_Share_Permission':
        return generate_cifs_permission_command(row)
    return None

def generate_vstore_command(row):
    """Generate vstore creation command."""
    name = row.get('Vstore', row.get('Name'))
    if not name:
        return None
    
    params = []
    if 'nas_capacity_quota' in row:
        params.append(f"nas_capacity_quota={row['nas_capacity_quota']}")
    if 'description' in row:
        params.append(f"description={row['description']}")
    
    param_str = " " + " ".join(params) if params else ""
    return f"create vstore general name={name}{param_str}"

def generate_filesystem_command(row):
    """Generate filesystem creation command."""
    name = row.get('Filesystem Name', row.get('Name'))
    if not name:
        return None
    
    required_params = []
    if 'pool_name' in row:
        required_params.append(f"pool_name={row['pool_name']}")
    elif 'pool_id' in row:
        required_params.append(f"pool_id={row['pool_id']}")
    
    optional_params = []
    for param in ['initial_distribute_policy', 'capacity', 'file_system_id', 
                 'alloc_type', 'owner_controller', 'compression_enabled',
                 'dedup_enabled', 'block_size', 'application_scenario',
                 'capacity_threshold', 'description']:
        if param in row:
            optional_params.append(f"{param}={row[param]}")
    
    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    return f"create file_system general name={name}{param_str}"

def generate_cifs_share_command(row):
    """Generate CIFS share creation command."""
    name = row.get('Share Name', row.get('Name'))
    local_path = row.get('Local Path')
    if not name or not local_path:
        return None
    
    required_params = [f"local_path={local_path}"]
    
    optional_params = []
    for param in ['continue_available_enabled', 'oplock_enabled', 'notify_enabled',
                 'file_system_id', 'file_system_name', 'offline_file_mode',
                 'smb2_ca_enabled', 'abe_enabled', 'apply_default_acl',
                 'show_snapshot_enabled', 'browse_enable']:
        if param in row:
            value = format_boolean(row[param])
            if value is not None:
                optional_params.append(f"{param}={value}")
    
    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    return f"create share cifs name={name}{param_str}"

def generate_nfs_share_command(row):
    """Generate NFS share creation command using the provided column names."""
    local_path = row.get('Local Path')
    if not local_path:
        return None
    
    required_params = [f"local_path={local_path}"]
    
    optional_params = []
    for param in ['file_system_id', 'alias', 'CharSet', 'Lock Type',
                 'show_snapshot_enabled', 'Share Description']:
        if param in row and not pd.isna(row[param]):
            value = format_boolean(row[param]) if param.endswith('_enabled') else row[param]
            optional_params.append(f"{param.replace(' ', '_').lower()}={value}")
    
    param_str = " " + " ".join(optional_params) if optional_params else ""
    return f"create share nfs{param_str}"

def generate_cifs_permission_command(row):
    """Generate CIFS share permission command with properly formatted permission types."""
    access_name = row.get('Access Name')
    share_id = row.get('Share ID')
    share_name = row.get('Share Name')
    permission_type = row.get('Permission Type', row.get('Access Type'))
    
    if not access_name or not (share_id or share_name):
        return None
    
    # Map permission types to the required format
    permission_mapping = {
        'read only': 'read_only',
        'read write': 'read_write',
        'no access': 'no_access',
        'all control': 'all_control',
        '1': 'read_only',
        '2': 'read_write',
        '3': 'no_access',
        '4': 'all_control'
    }
    
    # Format the permission type
    formatted_permission = None
    if permission_type:
        # Convert to lowercase and strip whitespace for matching
        perm_key = str(permission_type).lower().strip()
        formatted_permission = permission_mapping.get(perm_key, perm_key)
    
    if not formatted_permission:
        return None
    
    share_ref = f"share_id={share_id}" if share_id else f"share_name={share_name}"
    
    params = []
    if 'domain_type' in row:
        params.append(f"domain_type={row['domain_type']}")
    
    param_str = " " + " ".join(params) if params else ""
    return f"create share_permission cifs access_name={access_name} {share_ref} permission_type={formatted_permission}{param_str}"

def generate_nfs_permission_command(row):
    """Generate NFS share permission command using the provided column names."""
    access_name = row.get('Access Name')
    share_id = row.get('Share ID')
    
    if not access_name or not share_id:
        return None
    
    # Handle permission parameters with column name mapping
    param_mapping = {
        'Access Type': 'access_type',
        'Sync Enabled': 'sync_enabled',
        'All Squash Enabled': 'all_squash_enabled',
        'Root Squash Enabled': 'root_squash_enabled',
        'Secure Enabled': 'secure_enabled',
        'Share Permission Charset': 'charset',
        'Anonymous User ID': 'anonymous_user_id',
        'V4 Acl Preserve': 'v4_acl_preserve',
        'Ntfs Unix Security Ops': 'ntfs_unix_security_ops'
    }
    
    optional_params = []
    for excel_col, param in param_mapping.items():
        if excel_col in row and not pd.isna(row[excel_col]):
            value = format_boolean(row[excel_col]) if 'Enabled' in excel_col else row[excel_col]
            optional_params.append(f"{param}={value}")
    
    # Set defaults for required squash parameters if not specified
    if 'All Squash Enabled' not in row:
        optional_params.append("all_squash_enabled=no")
    if 'Root Squash Enabled' not in row:
        optional_params.append("root_squash_enabled=no")
    
    param_str = " " + " ".join(optional_params) if optional_params else ""
    return f"create share_permission nfs access_name={access_name} share_id={share_id}{param_str}"

def format_boolean(value):
    """Convert various boolean representations to yes/no."""
    if pd.isna(value):
        return None
    value = str(value).lower().strip()
    if value in ['enable', 'enabled', 'yes', 'true', '1']:
        return "yes"
    elif value in ['disable', 'disabled', 'no', 'false', '0']:
        return "no"
    return None