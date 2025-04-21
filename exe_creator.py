import subprocess
import os
import shutil

def create_executable():
    """Create the executable using PyInstaller."""
    # Path to the main script
    script = "main.py"

    # Additional files to include in the executable
    add_data = [
        ("command_generator.py", "."),
        ("file_operations.py", "."),
        ("gui_helpers.py", "."),
        ("import_operations.py", "."),
        ("log_operations.py", "."),
        ("oceanstor_dorado_commands.json", "."),
        ("oceanstor_pacific_commands.json", "."),
        ("utils.py", "."),
        ("Documents", "Documents"),  # Include the Documents directory
    ]

    # Path to the icon file (update this to your actual icon path)
    icon_path = os.path.join("Icons", "exe_icon.ico")

    # Output directory for the executable
    output_dir = "dist"
    executable_name = "Script Generator"

    # Clean up the output directory if it already exists
    if os.path.exists(output_dir):
        print(f"Deleting existing '{output_dir}' directory...")
        shutil.rmtree(output_dir)

    # Build the PyInstaller command
    pyinstaller_command = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--clean",
        f"--icon={icon_path}",
        f"--name={executable_name}",
        "--add-binary", f"{icon_path};Icons",  # Include the icon in the package
    ]

    # Add --add-data options
    for source, dest in add_data:
        if os.path.isdir(source):
            # For directories, we need to include all files
            for root, _, files in os.walk(source):
                for file in files:
                    src_path = os.path.join(root, file)
                    dest_path = os.path.relpath(root, start=os.path.dirname(source))
                    pyinstaller_command.append(f"--add-data={src_path}{os.pathsep}{os.path.join(dest, dest_path)}")
        else:
            pyinstaller_command.append(f"--add-data={source}{os.pathsep}{dest}")

    # Add the main script
    pyinstaller_command.append(script)

    # Execute the command to generate the executable
    print("Generating the executable...")
    print("Command:", " ".join(pyinstaller_command))
    subprocess.check_call(pyinstaller_command)

    print(f"Executable '{executable_name}.exe' generated successfully in {output_dir} directory.")

def main():
    """Main function to create the executable."""
    create_executable()

if __name__ == "__main__":
    main()