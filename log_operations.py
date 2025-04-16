import os
import datetime
import pandas as pd
from tkinter import messagebox

def create_logs_directory():
    """Create Logs directory if it doesn't exist."""
    logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logs")
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)
    return logs_dir

def generate_log_filename():
    """Generate timestamp-based log filename."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return f"import_log_{timestamp}.txt"

def count_commands_in_file(file_path, command_prefix, is_nfs=False, is_permission=False):
    """Count how many times a command appears in a file."""
    if not os.path.exists(file_path):
        return 0
    
    count = 0
    try:
        encodings = ['utf-8', 'latin-1', 'cp1252']
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    for line in f:
                        stripped = line.strip()
                        if stripped.startswith(command_prefix):
                            # For permissions, count all matching commands
                            if is_permission:
                                count += 1
                            # For NFS shares, check for local_path
                            elif is_nfs and 'local_path=' in stripped:
                                count += 1
                            # For non-NFS, check for name=
                            elif not is_nfs and 'name=' in stripped:
                                count += 1
                break
            except UnicodeDecodeError:
                continue
    except Exception as e:
        print(f"Warning: Error counting commands in {file_path}: {str(e)}")
        with open(file_path, 'r', errors='ignore') as f:
            for line in f:
                if line.strip().startswith(command_prefix):
                    count += 1
    return count

def count_excel_lines(xls, sheet_name):
    """Count non-empty lines in an Excel sheet with special handling."""
    # Special case: For NFS Share Permissions, always count all rows from NFS_Share sheet
    if sheet_name == 'NFS_Share_Permission':
        actual_sheet = 'NFS_Share'
    else:
        actual_sheet = sheet_name
    
    # Try different variations of the sheet name
    possible_names = {
        'NFS_Share': ['NFS_Share', 'NFS Shares', 'NFS Share'],
        'Filesystem': ['Filesystem', 'FileSystem', 'File Systems', 'Filesystems'],
        'CIFS_Share': ['CIFS_Share', 'CIFS Shares', 'CIFS Share']
    }
    
    # Get the actual sheet name from possible variations
    if actual_sheet in possible_names:
        for name in possible_names[actual_sheet]:
            if name in xls.sheet_names:
                actual_sheet = name
                break
    
    if actual_sheet not in xls.sheet_names:
        return 0
    
    df = xls.parse(actual_sheet)
    
    # For NFS Share sheet, count unique Local Paths (for shares) or all rows (for permissions)
    if actual_sheet == 'NFS_Share':
        if sheet_name == 'NFS_Share_Permission':
            # For permissions, count all rows
            return len(df.dropna(how='all'))
        elif 'Local Path' in df.columns:
            # For shares, count unique Local Paths
            return df['Local Path'].dropna().nunique()
    
    # Regular counting for other sheets
    return len(df.dropna(how='all'))

def create_import_log(file_path, results_dir, status):
    """Create a detailed log of the import operation."""
    try:
        logs_dir = create_logs_directory()
        log_file = os.path.join(logs_dir, generate_log_filename())
        
        xls = pd.ExcelFile(file_path)
        
        arrow = '→'
        try:
            with open(log_file, 'w', encoding='utf-8') as f:
                _write_log_content(f, xls, results_dir, status, arrow)
        except UnicodeEncodeError:
            arrow = '->'
            with open(log_file, 'w', encoding='utf-8', errors='replace') as f:
                _write_log_content(f, xls, results_dir, status, arrow)
        
        print(f"Log file created: {log_file}")
        return log_file
    except Exception as e:
        messagebox.showerror("Logging Error", f"Failed to create log file: {str(e)}")
        return None

def _write_log_content(file_obj, xls, results_dir, status, arrow):
    """Helper function to write log content."""
    file_obj.write(f"EXECUTION STATUS: {status}\n")
    file_obj.write("=" * 40 + "\n\n")
    
    command_stats = [
        ('Vstore', 'create vstore', 'Vstore_commands.txt', False, False),
        ('Filesystem', 'create file_system general', 'Filesystem_commands.txt', False, False),
        ('CIFS_Share', 'create share cifs', 'CIFS_Share_commands.txt', False, False),
        ('NFS_Share', 'create share nfs', 'NFS_Share_commands.txt', True, False),  # is_nfs=True
        ('CIFS_Share_Permission', 'create share_permission cifs', 'CIFS_Share_Permission_commands.txt', False, False),
        ('NFS_Share_Permission', 'create share_permission nfs', 'NFS_Share_Permission_commands.txt', True, True)  # is_nfs=True, is_permission=True
    ]
    
    # First parse the NFS_Share sheet to get both counts
    nfs_sheet_name = None
    nfs_sheet_variations = ['NFS_Share', 'NFS Shares', 'NFS Share']
    for name in nfs_sheet_variations:
        if name in xls.sheet_names:
            nfs_sheet_name = name
            break
    
    nfs_unique_count = 0
    nfs_total_count = 0
    if nfs_sheet_name:
        df = xls.parse(nfs_sheet_name)
        nfs_total_count = len(df.dropna(how='all'))
        if 'Local Path' in df.columns:
            nfs_unique_count = df['Local Path'].dropna().nunique()
    
    # Get counts for all other sheets
    sheet_counts = {sheet: count_excel_lines(xls, sheet) for sheet in xls.sheet_names}
    
    for item in command_stats:
        sheet_name, command_prefix, output_file, is_nfs, is_permission = item
        
        output_path = os.path.join(results_dir, output_file)
        commands_created = count_commands_in_file(
            output_path, command_prefix, is_nfs, is_permission
        )
        
        # Special handling for NFS sheets
        if sheet_name == 'NFS_Share':
            excel_lines = nfs_unique_count
        elif sheet_name == 'NFS_Share_Permission':
            excel_lines = nfs_total_count
        else:
            excel_lines = sheet_counts.get(sheet_name, 0)
            # Handle alternative sheet names
            if excel_lines == 0:
                alt_names = {
                    'Filesystem': ['FileSystem', 'File Systems', 'Filesystems'],
                    'CIFS_Share': ['CIFS Shares', 'CIFS Share']
                }
                if sheet_name in alt_names:
                    for name in alt_names[sheet_name]:
                        if name in sheet_counts:
                            excel_lines = sheet_counts[name]
                            break
        
        file_obj.write(f"{sheet_name}:\n")
        file_obj.write(f"  {arrow} Created: {commands_created}\n")
        file_obj.write(f"  {arrow} Excel Lines: {excel_lines}\n\n")
    
    file_obj.write("\n" + "=" * 40 + "\n")
    file_obj.write(f"Log generated at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")