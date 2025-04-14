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

def count_commands_in_file(file_path, command_prefix):
    """Count how many times a command appears in a file."""
    if not os.path.exists(file_path):
        return 0
    
    count = 0
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip().startswith(command_prefix):
                    count += 1
    except:
        # Fallback if UTF-8 fails
        with open(file_path, 'r') as f:
            for line in f:
                if line.strip().startswith(command_prefix):
                    count += 1
    return count

def count_excel_lines(xls, sheet_name):
    """Count non-empty lines in an Excel sheet."""
    if sheet_name not in xls.sheet_names:
        return 0
    df = xls.parse(sheet_name)
    return len(df.dropna(how='all'))

def create_import_log(file_path, results_dir, status):
    """
    Create a detailed log of the import operation.
    
    Args:
        file_path (str): Path to the imported Excel file
        results_dir (str): Path to the Imported_Results directory
        status (str): 'Success' or 'Failed'
    """
    try:
        logs_dir = create_logs_directory()
        log_file = os.path.join(logs_dir, generate_log_filename())
        
        xls = pd.ExcelFile(file_path)
        
        # Replace Unicode arrow with ASCII alternative if encoding fails
        arrow = '→'
        try:
            # First try with UTF-8 encoding
            with open(log_file, 'w', encoding='utf-8') as f:
                _write_log_content(f, xls, results_dir, status, arrow)
        except UnicodeEncodeError:
            # Fallback to ASCII with replacement character
            arrow = '->'
            with open(log_file, 'w', encoding='utf-8', errors='replace') as f:
                _write_log_content(f, xls, results_dir, status, arrow)
        
        return log_file
    except Exception as e:
        messagebox.showerror("Logging Error", f"Failed to create log file: {str(e)}")
        return None

def _write_log_content(file_obj, xls, results_dir, status, arrow):
    """Helper function to write log content with proper encoding handling."""
    # Write header with improved formatting
    file_obj.write(f"EXECUTION STATUS: {status}\n")
    file_obj.write("=" * 40 + "\n\n")
    
    # Process each command type with better spacing
    command_stats = [
        ('Vstore', 'create vstore', 'Vstore_commands.txt'),
        ('Filesystem', 'create file_system', 'Filesystem_commands.txt'),
        ('CIFS_Share', 'create share cifs', 'CIFS_Share_commands.txt'),
        ('NFS_Share', 'create share nfs', 'NFS_Share_commands.txt'),
        ('CIFS_Share_Permission', 'create share_permission cifs', 'CIFS_Share_Permission_commands.txt'),
        ('NFS_Share_Permission', 'create share_permission nfs', 'NFS_Share_Permission_commands.txt')
    ]
    
    # Special handling for NFS shares and permissions from same sheet
    nfs_sheet_lines = count_excel_lines(xls, 'NFS_Share') if 'NFS_Share' in xls.sheet_names else 0
    
    for item in command_stats:
        sheet_name, command_prefix, output_file = item
        
        output_path = os.path.join(results_dir, output_file)
        commands_created = count_commands_in_file(output_path, command_prefix)
        
        # For NFS permissions, use the same sheet count as NFS shares
        excel_lines = nfs_sheet_lines if 'NFS_Share' in sheet_name else count_excel_lines(xls, sheet_name)
        
        file_obj.write(f"{sheet_name}:\n")
        file_obj.write(f"  {arrow} Created: {commands_created}\n")
        
        # Only show Excel lines count for the main sheet
        if not sheet_name.endswith('_Permission'):
            file_obj.write(f"  {arrow} Excel Lines: {excel_lines}\n\n")
        else:
            file_obj.write("\n")
    
    # Improved timestamp formatting
    file_obj.write("\n" + "=" * 40 + "\n")
    file_obj.write(f"Log generated at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")