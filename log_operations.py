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

def count_commands_in_file(file_path, command_prefix, is_nfs=False):
    """Count how many times a command appears in a file."""
    if not os.path.exists(file_path):
        return 0
    
    count = 0
    try:
        # Try UTF-8 first, then fall back to other encodings
        encodings = ['utf-8', 'latin-1', 'cp1252']
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    for line in f:
                        stripped = line.strip()
                        if stripped.startswith(command_prefix):
                            # For NFS shares, check for local_path instead of name
                            if is_nfs and 'local_path=' in stripped:
                                count += 1
                            elif not is_nfs and 'name=' in stripped:
                                count += 1
                break  # If we got here, encoding worked
            except UnicodeDecodeError:
                continue
    except Exception as e:
        print(f"Warning: Error counting commands in {file_path}: {str(e)}")
        # Fallback to simple counting
        with open(file_path, 'r', errors='ignore') as f:
            for line in f:
                if line.strip().startswith(command_prefix):
                    count += 1
    return count

def count_excel_lines(xls, sheet_name):
    """Count non-empty lines in an Excel sheet with special handling."""
    # Try different variations of the sheet name
    possible_names = {
        'NFS_Share_Permission': ['NFS_Share_Permission', 'NFS Share Permissions', 'NFS Permissions'],
        'NFS_Share': ['NFS_Share', 'NFS Shares', 'NFS Share'],
        # Add other sheets if needed
    }
    
    if sheet_name == 'NFS_Share_Permission':
        sheet_name = 'NFS_Share'
    
    # Get the actual sheet name from possible variations
    actual_sheet = sheet_name
    if sheet_name in possible_names:
        for name in possible_names[sheet_name]:
            if name in xls.sheet_names:
                actual_sheet = name
                break
    
    if actual_sheet not in xls.sheet_names:
        return 0
    
    df = xls.parse(actual_sheet)
    
    # Special handling for NFS_Share - count unique Local Paths
    if sheet_name == 'NFS_Share' and 'Local Path' in df.columns:
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
        ('Vstore', 'create vstore', 'Vstore_commands.txt', False),
        ('Filesystem', 'create file_system general', 'Filesystem_commands.txt', False),
        ('CIFS_Share', 'create share cifs', 'CIFS_Share_commands.txt', False),
        ('NFS_Share', 'create share nfs', 'NFS_Share_commands.txt', True),  # Mark as NFS
        ('CIFS_Share_Permission', 'create share_permission cifs', 'CIFS_Share_Permission_commands.txt', False),
        ('NFS_Share_Permission', 'create share_permission nfs', 'NFS_Share_Permission_commands.txt', True)  # Mark as NFS
    ]
    
    sheet_counts = {sheet: count_excel_lines(xls, sheet) for sheet in xls.sheet_names}
    
    for item in command_stats:
        sheet_name, command_prefix, output_file, is_nfs = item
        
        output_path = os.path.join(results_dir, output_file)
        commands_created = count_commands_in_file(output_path, command_prefix, is_nfs)
        
        excel_lines = sheet_counts.get(sheet_name, 0)
        
        # Handle alternative sheet names
        if excel_lines == 0:
            alt_names = {
                'Filesystem': ['FileSystem', 'File Systems', 'Filesystems'],
                'NFS_Share_Permission': ['NFS Share Permissions', 'NFS Permissions'],
                'NFS_Share': ['NFS Shares', 'NFS Share']
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