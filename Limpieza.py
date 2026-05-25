import os
import time
from pathlib import Path

# Carpeta donde están los archivos Excel
CARPETA = r"C:/Users/Janus I5/Desktop/easygestion/descargas"  # Cambia esta ruta si es necesario

# Tiempo máximo permitido (en horas)
HORAS_MAXIMAS = 24

# Convertir horas a segundos
TIEMPO_LIMITE = HORAS_MAXIMAS * 60 * 60

# Tiempo actual
ahora = time.time()

# Buscar archivos Excel
for archivo in Path(CARPETA).glob("*.xlsx"):

    # Obtener última modificación
    fecha_modificacion = archivo.stat().st_mtime

    # Calcular antigüedad
    antiguedad = ahora - fecha_modificacion

    # Si supera el tiempo límite → eliminar
    if antiguedad > TIEMPO_LIMITE:
        try:
            os.remove(archivo)
            print(f"Archivo eliminado: {archivo.name}")
        except Exception as e:
            print(f"Error eliminando {archivo.name}: {e}")

print("Limpieza finalizada.")