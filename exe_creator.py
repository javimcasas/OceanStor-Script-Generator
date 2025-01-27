import subprocess
import os
import sys
import shutil

def create_executable():
    """Create the executable using PyInstaller."""
    # Path to the main script
    script = "script_selector.py"

    # Additional files to include in the executable
    add_data = [
        "commands_config.json;.",  # Include the JSON configuration file
        "help_links.json;.",       # Include the JSON help links file
        "command_generator.py;.",  # Include the command generation file
        "utils.py;.",              # Include utility functions
        "file_operations.py;.",    # Include file operations module
        "gui_helpers.py;.",        # Include GUI helpers module
    ]

    # Build the --add-data option
    add_data_option = " ".join([f'--add-data "{item}"' for item in add_data])

    # Output directory for the executable
    output_dir = "dist"

    # Clean up the output directory if it already exists
    if os.path.exists(output_dir):
        print(f"Deleting existing '{output_dir}' directory...")
        shutil.rmtree(output_dir)

    # Command to generate the executable
    pyinstaller_command = (
        f'pyinstaller --onefile --windowed --clean {add_data_option} {script}'
    )

    # Execute the command to generate the executable
    print("Generating the executable...")
    subprocess.check_call(pyinstaller_command, shell=True)

    print("Executable generated successfully.")

def main():
    """Main function to create the executable."""
    create_executable()

if __name__ == "__main__":
    main()