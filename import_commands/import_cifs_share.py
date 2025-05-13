import pandas as pd

def append_y_command_if_browse_enabled(optional_params):
    """
    Appends a 'y' command to the result if 'browse_enabled' is 'yes'.
    This function can be commented out if not needed.
    """
    for param in optional_params:
        if param.startswith('browse_enabled=yes'):
            return "\n y"
    return ""

def generate_cifs_share_command(row):
    """Generate CIFS share creation command with proper yes/no values for enabled parameters."""
    name = row.get('Share Name', row.get('Name'))
    local_path = row.get('Local Path')
    if not name or not local_path:
        return None

    required_params = [
        f"name={str(name).lower()}",
        f"local_path={str(local_path).lower()}"
    ]

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
        'File Extension Filter': 'file_filter_enable',
        'Apply Default ACL': 'apply_default_acl',
        'Show Previous Versions Enabled': 'show_previous_versions_enabled',
        'Show Snapshot Enabled': 'show_snapshot_enabled',
        'Browse Enabled': 'browse_enabled',
        'Readdir Timeout(s)': 'readdir_timeout'
    }

    yes_no_params = [
        'oplock_enabled', 'notify_enabled', 'continue_available_enabled',
        'smb2_ca_enabled', 'ip_control_enabled', 'abe_enabled',
        'file_filter_enable', 'apply_default_acl',
        'show_previous_versions_enabled', 'show_snapshot_enabled',
        'browse_enabled'
    ]

    optional_params = []
    for excel_col, param in param_mapping.items():
        if excel_col in row and pd.notna(row[excel_col]):
            value = row[excel_col]
            value = str(value).lower()

            if param in yes_no_params:
                if value in ['enable', 'enabled', 'yes', 'true', '1']:
                    value = 'yes'
                elif value in ['disable', 'disabled', 'no', 'false', '0']:
                    value = 'no'
                else:
                    continue

            if param == 'offline_file_mode':
                value = value.replace(' ', '_')

            optional_params.append(f"{param}={value}")

    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""

    # Call the function to append 'y' if browse_enabled is 'yes'
    y_command = append_y_command_if_browse_enabled(optional_params)

    return f"create share cifs{param_str}{y_command}"
