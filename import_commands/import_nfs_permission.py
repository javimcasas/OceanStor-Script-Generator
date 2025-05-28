import pandas as pd
from .import_utils import format_boolean

def generate_nfs_permission_command(row):
    """Generate NFS share permission command using all available parameters from NFS_Share sheet."""
    access_name = row.get('Access Name')
    share_id = row.get('Share ID')
    share_name = row.get('Local Path')
    
    if not access_name or not share_id or not share_name:
        return None
    
    required_params = [
        f"access_name={access_name}",
        f"share_name={share_name}"
    ]
    
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
    if 'Access Type' in row and not pd.isna(row['Access Type']):
        perm_value = str(row['Access Type']).lower().strip()
        mapped_perm = permission_mapping.get(perm_value, perm_value)
        optional_params.append(f"access_type={mapped_perm}")
    
    for excel_col, param in param_mapping.items():
        if excel_col == 'Access Type':
            continue
        
        if excel_col in row and not pd.isna(row[excel_col]):
            value = row[excel_col]
            
            if 'Enabled' in excel_col:
                value = format_boolean(value)
                if value is None:
                    continue
            elif excel_col == 'V4 Acl Preserve':
                value = 'true' if str(value).lower() in ['true', 'yes', 'enable'] else 'false'
            
            if not str(value).isdigit() and excel_col != 'Anonymous User ID':
                value = str(value).lower()
            
            optional_params.append(f"{param}={value}")
    
    if 'All Squash Enabled' not in row:
        optional_params.append("all_squash_enabled=no")
    if 'Root Squash Enabled' not in row:
        optional_params.append("root_squash_enabled=no")
    
    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    
    return f"create share_permission nfs{param_str}"