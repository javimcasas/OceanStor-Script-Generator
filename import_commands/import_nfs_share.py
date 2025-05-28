import pandas as pd
from .import_utils import format_boolean, fix_description

processed_paths = set()
def generate_nfs_share_command(row):
    """Generate NFS share creation command using the provided column names."""
    local_path = row.get('Local Path')
    if not local_path:
        return None

    if local_path.endswith('/'):
        local_path = local_path[:-1]
    
    if local_path in processed_paths:
        return None
    
    processed_paths.add(local_path)
    
    required_params = [f"local_path={local_path}"]
    
    fs_ref = None
    if 'Filesystem ID' in row and not pd.isna(row['Filesystem ID']):
        fs_ref = f"file_system_id={row['Filesystem ID']}"
    elif 'file_system_name' in row and not pd.isna(row['file_system_name']):
        fs_ref = f"file_system_name={row['file_system_name']}"
    
    if fs_ref:
        required_params.append(fs_ref)
    
    optional_params = []
    param_mapping = {
        'CharSet': 'charset',
        'Lock Type': 'lock_type',
        'Alias': 'alias',
        'Audit Items': 'audit_items',
        'show_snapshot_enabled': 'show_snapshot_enabled',
        'fh_byte_alignment_switch': 'fh_byte_alignment_switch',
        'Share Description': 'description'
    }
    
    for excel_col, param in param_mapping.items():
        if excel_col in row and not pd.isna(row[excel_col]):
            value = row[excel_col]
            
            if param == 'description':
                value = fix_description(value)
                
            if excel_col.endswith('_enabled') or excel_col.endswith('_switch'):
                value = format_boolean(value)
                if value is None:
                    continue
            optional_params.append(f"{param}={value}")
    
    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    
    return f"create share nfs{param_str}"