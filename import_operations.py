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
    """Generate filesystem creation command with all supported parameters."""
    name = row.get('Filesystem Name')
    if not name:
        return None
    
    # Map Excel column names to command parameters
    param_mapping = {
        'pool_name': 'pool_name',
        'pool_id': 'pool_id',
        'initial_distribute_policy': 'initial_distribute_policy',
        'Capacity': 'capacity',
        'Filesystem ID': 'file_system_id',
        'Type': 'alloc_type',
        'owner_controller': 'owner_controller',
        'io_priority': 'io_priority',
        'compression_enabled': 'compression_enabled',
        'compression_method': 'compression_method',
        'dedup_enabled': 'dedup_enabled',
        'intelligent_dedup_enabled': 'intelligent_dedup_enabled',
        'bytecomparison_enabled': 'bytecomparison_enabled',
        'dedup_metadata_sample_ratio': 'dedup_metadata_sample_ratio',
        'checksum_enabled': 'checksum_enabled',
        'Atime Enabled': 'atime_enabled',
        'atime_update_mode': 'atime_update_mode',
        'Show Snapshot Directory Enabled': 'show_enabled',
        'Auto Delete Snapshot Enabled': 'auto_delete_snapshot_enabled',
        'Timing Snapshot Enabled': 'timing_snapshot_enabled',
        'Block Size': 'block_size',
        'Application Scenario': 'application_scenario',
        'Capacity Threshold(%)': 'capacity_threshold',
        'Snapshot Reserve(%)': 'snapshot_reserve',
        'Timing Snapshot Max Number': 'timing_snapshot_max_number',
        'Sub Type': 'sub_type',
        'worm_type': 'worm_type',
        'auto_delete_enabled': 'auto_delete_enabled',
        'auto_lock_enabled': 'auto_lock_enabled',
        'default_protect_period': 'default_protect_period',
        'default_protect_period_unit': 'default_protect_period_unit',
        'min_protect_period': 'min_protect_period',
        'min_protect_period_unit': 'min_protect_period_unit',
        'max_protect_period': 'max_protect_period',
        'max_protect_period_unit': 'max_protect_period_unit',
        'auto_lock_time': 'auto_lock_time',
        'auto_lock_time_unit': 'auto_lock_time_unit',
        'smart_cache_state': 'smart_cache_state',
        'space_self_adjusting_mode': 'space_self_adjusting_mode',
        'autosize_enable': 'autosize_enable',
        'auto_shrink_threshold_percent': 'auto_shrink_threshold_percent',
        'auto_grow_threshold_percent': 'auto_grow_threshold_percent',
        'min_autosize': 'min_autosize',
        'max_autosize': 'max_autosize',
        'autosize_increment': 'autosize_increment',
        'space_recycle_mode': 'space_recycle_mode',
        'Description': 'description',
        'Alternate Data Streams Enabled': 'alternate_data_streams_enabled',
        'ssd_capacity_upper_limit_of_user_data': 'ssd_capacity_upper_limit_of_user_data',
        'Long Filename Enabled': 'long_filename_enabled',
        'Security Style': 'security_style',
        'unix_permissions': 'unix_permissions',
        'workload_type_id': 'workload_type_id',
        'fs_layer_distribution_algorithm': 'fs_layer_distribution_algorithm',
        'is_auditlog_fs': 'is_auditlog_fs',
        'is_worm_auditlog_fs': 'is_worm_auditlog_fs',
        'hyper_cdp_schedule_name': 'hyper_cdp_schedule_name',
        'audit_items': 'audit_items',
        'quota_switch': 'quota_switch',
        'nas_locking_policy': 'nas_locking_policy',
        'fs_capacity_switch': 'fs_capacity_switch',
        'character_set': 'character_set',
        'VAAI_switch': 'VAAI_switch',
        'dir_placement_enabled': 'dir_placement_enabled',
        'snapdiff_switch': 'snapdiff_switch'
    }

    # Handle required parameters - default to pool_id=1 if neither pool_name nor pool_id is specified
    required_params = []
    if 'pool_name' in row and pd.notna(row['pool_name']):
        required_params.append(f"pool_name={row['pool_name']}")
    elif 'pool_id' in row and pd.notna(row['pool_id']):
        required_params.append(f"pool_id={row['pool_id']}")
    else:
        required_params.append("pool_id=1")  # Default pool if none specified

    # Handle optional parameters
    optional_params = []
    for excel_col, param in param_mapping.items():
        if excel_col in row and pd.notna(row[excel_col]):
            value = row[excel_col]
            
            # Special handling for boolean values
            if isinstance(value, str) and value.lower() in ['enable', 'enabled', 'yes', 'true']:
                value = 'yes'
            elif isinstance(value, str) and value.lower() in ['disable', 'disabled', 'no', 'false']:
                value = 'no'
            
            # Special handling for specific parameters
            if param == 'alloc_type' and isinstance(value, str):
                value = value.lower()  # Convert alloc_type to lowercase
                
            elif param == 'block_size' and isinstance(value, str):
                # Remove decimal part and any duplicate units from block_size
                if '.' in value:
                    value = value.split('.')[0] + ''.join([c for c in value[value.find('.'):] if not c.isdigit()])
                    value = value.replace('.', '')
                # Remove any duplicate units (like KBKB)
                units = ['KB', 'MB', 'GB', 'TB']
                for unit in units:
                    if unit in value and value.endswith(unit):
                        value = value.replace(unit, '') + unit
                
            elif param == 'application_scenario' and isinstance(value, str):
                # Only include if value is database or virtual_machine (case-insensitive)
                normalized_value = value.lower().replace(' ', '_')
                if normalized_value in ['database', 'virtual_machine']:
                    value = normalized_value
                else:
                    continue  # Skip this parameter
                    
            elif param in ['capacity_threshold', 'snapshot_reserve'] and isinstance(value, str):
                # Remove % sign if present
                value = value.replace('%', '')
                
            elif param == 'sub_type' and isinstance(value, str):
                value = value.lower()  # Convert sub_type to lowercase
            
            optional_params.append(f"{param}={value}")

    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    return f"create file_system general name={name}{param_str}"

def generate_cifs_share_command(row):
    """Generate CIFS share creation command with proper yes/no values for enabled parameters."""
    name = row.get('Share Name', row.get('Name'))
    local_path = row.get('Local Path')
    if not name or not local_path:
        return None
    
    # Required parameters
    required_params = [
        f"name={str(name).lower()}",
        f"local_path={str(local_path).lower()}"
    ]
    
    # Map Excel column names to command parameters
    param_mapping = {
        'Share Type': 'share_type',
        'Oplock Enabled': 'oplock_enabled',
        'Notify Enabled': 'notify_enabled',
        'Continue Available Enabled': 'continue_available_enabled',
        'Offline File Mode': 'offline_file_mode',
        'Smb2 CA Enabled': 'smb2_ca_enabled',
        'IP Access Control': 'ip_control_enabled',
        'ABE Enabled': 'abe_enabled',
        'Audit Items': 'audit_items',
        'File Extenson Filter': 'file_filter_enable',
        'Apply Default ACL': 'apply_default_acl',
        'Show Previous Versions Enabled': 'show_previous_versions_enabled',
        'Show Snapshot Enabled': 'show_snapshot_enabled',
        'Browse Enabled': 'browse_enabled',
        'Readdir Timeout(s)': 'readdir_timeout'
    }
    
    # List of parameters that should be converted to yes/no
    yes_no_params = [
        'oplock_enabled', 'notify_enabled', 'continue_available_enabled',
        'smb2_ca_enabled', 'ip_control_enabled', 'abe_enabled',
        'file_filter_enable', 'apply_default_acl', 
        'show_previous_versions_enabled', 'show_snapshot_enabled',
        'browse_enable'
    ]
    
    # Handle optional parameters
    optional_params = []
    for excel_col, param in param_mapping.items():
        if excel_col in row and pd.notna(row[excel_col]):
            value = row[excel_col]
            
            # Convert to string and lowercase
            value = str(value).lower()
            
            # Special handling for yes/no parameters
            if param in yes_no_params:
                if value in ['enable', 'enabled', 'yes', 'true', '1']:
                    value = 'yes'
                elif value in ['disable', 'disabled', 'no', 'false', '0']:
                    value = 'no'
                else:
                    continue  # Skip if value isn't recognized
            
            # Special handling for offline_file_mode
            if param == 'offline_file_mode':
                value = value.replace(' ', '_')  # manual -> manual, auto -> auto
            
            optional_params.append(f"{param}={value}")
    
    # Combine all parameters
    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    
    return f"create share cifs{param_str}"

processed_paths = set()
def generate_nfs_share_command(row):
    """Generate NFS share creation command using the provided column names."""
    local_path = row.get('Local Path')
    if not local_path:
        return None
    
    # Check if this local_path has already been processed
    if local_path in processed_paths:
        return None  # Skip the creation of duplicate share
    
    # Mark this local_path as processed
    processed_paths.add(local_path)
    
    # Required parameter
    required_params = [f"local_path={local_path}"]
    
    # File system reference (either ID or name)
    fs_ref = None
    if 'Filesystem ID' in row and not pd.isna(row['Filesystem ID']):
        fs_ref = f"file_system_id={row['Filesystem ID']}"
    elif 'file_system_name' in row and not pd.isna(row['file_system_name']):
        fs_ref = f"file_system_name={row['file_system_name']}"
    
    if fs_ref:
        required_params.append(fs_ref)
    
    # Optional parameters with proper formatting
    optional_params = []
    param_mapping = {
        'CharSet': 'charset',
        'Lock Type': 'lock_type',
        'Alias': 'alias',
        'Audit Items': 'audit_items',
        'show_snapshot_enabled': 'show_snapshot_enabled',
        'fh_byte_alignment_switch': 'fh_byte_alignment_switch'
    }
    
    for excel_col, param in param_mapping.items():
        if excel_col in row and not pd.isna(row[excel_col]):
            value = row[excel_col]
            # Handle boolean values
            if excel_col.endswith('_enabled') or excel_col.endswith('_switch'):
                value = format_boolean(value)
                if value is None:
                    continue
            optional_params.append(f"{param}={value}")
    
    # Combine all parameters
    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    
    return f"create share nfs{param_str}"

def generate_cifs_permission_command(row):
    """Generate CIFS share permission command with comprehensive parameter handling."""
    access_name = row.get('Access Name')
    share_id = row.get('Share ID')
    share_name = row.get('Share Name')
    permission_type = row.get('Permission Type', row.get('Access Type'))
    
    if not access_name or not (share_id or share_name):
        return None
    
    # Required parameters
    share_ref = f"share_name={share_name}" if share_name else f"share_id={share_id}"
    required_params = [
        f"access_name={access_name}",
        share_ref
    ]
    
    # Permission type mapping (Excel values to command values)
    permission_mapping = {
        'read only': 'read_only',
        'read-only': 'read_only',
        'read write': 'read_write',
        'read-write': 'read_write',
        'read and write': 'read_write',
        'no access': 'no_access',
        'no-access': 'no_access',
        'all control': 'all_control',
        'all-control': 'all_control',
        'full control': 'all_control',
        'full-control': 'all_control',
        '1': 'read_only',
        '2': 'read_write',
        '3': 'no_access',
        '4': 'all_control'
    }
    
    # Handle permission type
    optional_params = []
    if permission_type and not pd.isna(permission_type):
        perm_value = str(permission_type).lower().strip()
        formatted_permission = permission_mapping.get(perm_value, perm_value)
        optional_params.append(f"permission_type={formatted_permission}")
    else:
        return None  # Permission type is required
    
    # Handle other optional parameters
    if 'Domain Type' in row and not pd.isna(row['Domain Type']):
        domain_type = str(row['Domain Type']).lower()
        optional_params.append(f"domain_type={domain_type}")
    
    # Handle additional parameters that might be present
    additional_params = {
        'sync_enabled': 'Sync Enabled',
        'inherit_enabled': 'Inherit Enabled',
        'inherit_owner': 'Inherit Owner',
        'inherit_group': 'Inherit Group',
        'inherit_dacl': 'Inherit DACL',
        'inherit_sacl': 'Inherit SACL'
    }
    
    for param, excel_col in additional_params.items():
        if excel_col in row and not pd.isna(row[excel_col]):
            value = format_boolean(row[excel_col])
            if value is not None:
                optional_params.append(f"{param}={value}")
    
    # Combine all parameters
    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    
    return f"create share_permission cifs{param_str}"

def generate_nfs_permission_command(row):
    """Generate NFS share permission command using all available parameters from NFS_Share sheet."""
    access_name = row.get('Access Name')
    share_id = row.get('Share ID')
    share_name = row.get('Alias')
    
    if not access_name or not share_id or not share_name:
        return None
    
    # Required parameters
    required_params = [
        f"access_name={access_name}",
        #f"share_id={share_id}",
        f"share_name={share_name}"
    ]
    
    # Permission type mapping (Excel values to command values)
    permission_mapping = {
        'read only': 'read_only',
        'read-only': 'read_only',
        'read write': 'read_write',
        'read-write': 'read_write',
        'read and write': 'read_write',
        'no permission': 'no_permission',
        'no-permission': 'no_permission',
        '1': 'read_only',
        '2': 'read_write',
        '5': 'no_permission'
    }
    
    # Parameter mapping from Excel to command
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
    
    # Handle permission type separately
    optional_params = []
    if 'Access Type' in row and not pd.isna(row['Access Type']):
        perm_value = str(row['Access Type']).lower().strip()
        mapped_perm = permission_mapping.get(perm_value, perm_value)
        optional_params.append(f"access_type={mapped_perm}")
    
    # Handle other parameters
    for excel_col, param in param_mapping.items():
        if excel_col == 'Access Type':
            continue  # Already handled
        
        if excel_col in row and not pd.isna(row[excel_col]):
            value = row[excel_col]
            
            # Convert boolean values to yes/no
            if 'Enabled' in excel_col:
                value = format_boolean(value)
                if value is None:
                    continue
            elif excel_col == 'V4 Acl Preserve':
                value = 'true' if str(value).lower() in ['true', 'yes', 'enable'] else 'false'
            
            # Convert to string and lowercase if not a numeric value
            if not str(value).isdigit() and excel_col != 'Anonymous User ID':
                value = str(value).lower()
            
            optional_params.append(f"{param}={value}")
    
    # Set defaults for required squash parameters if not specified
    if 'All Squash Enabled' not in row:
        optional_params.append("all_squash_enabled=no")
    if 'Root Squash Enabled' not in row:
        optional_params.append("root_squash_enabled=no")
    
    # Combine all parameters
    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    
    return f"create share_permission nfs{param_str}"

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