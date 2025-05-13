import pandas as pd
from .import_utils import format_boolean

def generate_cifs_permission_command(row):
    """Generate CIFS share permission command with comprehensive parameter handling."""
    access_name = row.get('Access Name')
    share_id = row.get('Share ID')
    share_name = row.get('Share Name')
    permission_type = row.get('Permission Type', row.get('Access Type'))
    
    if not access_name or not (share_id or share_name):
        return None
    
    share_ref = f"share_name={share_name}" if share_name else f"share_id={share_id}"
    required_params = [
        f"access_name={access_name}",
        share_ref
    ]
    
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
    
    optional_params = []
    if permission_type and not pd.isna(permission_type):
        perm_value = str(permission_type).lower().strip()
        formatted_permission = permission_mapping.get(perm_value, perm_value)
        optional_params.append(f"permission_type={formatted_permission}")
    else:
        return None
    
    if 'Domain Type' in row and not pd.isna(row['Domain Type']):
        domain_type = str(row['Domain Type']).lower()
        optional_params.append(f"domain_type={domain_type}")
    
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
    
    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    
    return f"create share_permission cifs{param_str}"
