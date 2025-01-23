import subprocess
import os
import sys

def install_pyinstaller():
    """Verifica si PyInstaller está instalado, y si no, lo instala."""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "show", "pyinstaller"])
        print("PyInstaller ya está instalado.")
    except subprocess.CalledProcessError:
        print("PyInstaller no está instalado. Instalando...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

def create_executable():
    """Crea el ejecutable usando PyInstaller."""
    # Ruta a tu script principal
    script = 'script_selector.py'

    # Archivos adicionales que se deben incluir en el ejecutable
    add_data = [
        "nfs_share_script.py;.",
        "cifs_share_script.py;.",
        "readData.py;.",
        "Documents\\CIFSShares.xlsx;Documents",
        "Documents\\NFSShares.xlsx;Documents",
    ]

    # Construimos la opción de --add-data
    add_data_option = ' '.join([f'--add-data "{item}"' for item in add_data])

    # Ruta de salida para el ejecutable
    output_dir = "dist"

    # Verifica si ya existe una carpeta dist y la elimina si es necesario
    if os.path.exists(output_dir):
        print(f"Eliminando la carpeta '{output_dir}' existente...")
        for root, dirs, files in os.walk(output_dir, topdown=False):
            for filename in files:
                os.remove(os.path.join(root, filename))
            for dirname in dirs:
                os.rmdir(os.path.join(root, dirname))
        os.rmdir(output_dir)

    # Comando para generar el ejecutable
    pyinstaller_command = f'pyinstaller --onefile --windowed {add_data_option} {script}'

    # Ejecuta el comando para generar el ejecutable
    print("Generando el ejecutable...")
    subprocess.check_call(pyinstaller_command, shell=True)

    print("El ejecutable ha sido generado correctamente.")

def main():
    """Función principal para instalar PyInstaller y crear el ejecutable."""
    # Uncomment the following line if you need to ensure PyInstaller is installed
    # install_pyinstaller()
    create_executable()

if __name__ == "__main__":
    main()
