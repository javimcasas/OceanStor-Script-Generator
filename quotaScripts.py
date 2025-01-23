import os, sys
from readData import readFile

# Definir la ruta de la carpeta 'Results'
carpeta_results = 'Results'

# Verificar si la carpeta 'Results' existe, si no, crearla
if not os.path.exists(carpeta_results):
    os.makedirs(carpeta_results)

# Ruta completa del archivo
archivo_path = os.path.join(carpeta_results, 'hola_mundo.txt')

# Crear y escribir en el archivo
with open(archivo_path, 'w') as archivo:
    archivo.write('Hola mundo')

print(f"El archivo ha sido creado en: {archivo_path}")
