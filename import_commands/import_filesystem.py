import pandas as pd

from .import_utils import fix_capacity, fix_description

def include_hypercdp_schedule(command_str):
    """Add hyper_cdp_schedule_name=NONE to the end of the command if not already present."""
    if "hyper_cdp_schedule_name=" not in command_str:
        if command_str.strip().endswith('"'):
            # Handle case where command ends with a quote
            return command_str[:-1] + ' hyper_cdp_schedule_name=NONE"'
        return command_str + " hyper_cdp_schedule_name=NONE"
    return command_str

def generate_filesystem_command(row):
    """Generate filesystem creation command with all supported parameters."""
    name = row.get('Filesystem Name')
    if not name:
        return None
    
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

    required_params = []
    if 'pool_name' in row and pd.notna(row['pool_name']):
        required_params.append(f"pool_name={row['pool_name']}")
    elif 'pool_id' in row and pd.notna(row['pool_id']):
        required_params.append(f"pool_id={row['pool_id']}")
    else:
        required_params.append("pool_id=1")

    optional_params = []
    for excel_col, param in param_mapping.items():
        if excel_col in row and pd.notna(row[excel_col]):
            value = row[excel_col]
            
            # Apply specific fixes based on parameter type
            if param == 'capacity':
                value = fix_capacity(value)
            elif param == 'description':
                value = fix_description(value)
            
            if isinstance(value, str) and value.lower() in ['enable', 'enabled', 'yes', 'true']:
                value = 'yes'
            elif isinstance(value, str) and value.lower() in ['disable', 'disabled', 'no', 'false']:
                value = 'no'
            
            if param == 'alloc_type' and isinstance(value, str):
                value = value.lower()
            elif param == 'block_size' and isinstance(value, str):
                if '.' in value:
                    value = value.split('.')[0] + ''.join([c for c in value[value.find('.'):] if not c.isdigit()])
                    value = value.replace('.', '')
                units = ['KB', 'MB', 'GB', 'TB']
                for unit in units:
                    if unit in value and value.endswith(unit):
                        value = value.replace(unit, '') + unit
            elif param == 'application_scenario' and isinstance(value, str):
                normalized_value = value.lower().replace(' ', '_')
                if normalized_value in ['database', 'virtual_machine']:
                    value = normalized_value
                else:
                    continue
            elif param in ['capacity_threshold', 'snapshot_reserve'] and isinstance(value, str):
                value = value.replace('%', '')
            elif param == 'sub_type' and isinstance(value, str):
                value = value.lower()
            
            optional_params.append(f"{param}={value}")

    all_params = required_params + optional_params
    param_str = " " + " ".join(all_params) if all_params else ""
    
    command = f"create file_system general name={name}{param_str}"
    
    # Add hyper_cdp_schedule_name=NONE if not already present
    command = include_hypercdp_schedule(command)
    
    return command