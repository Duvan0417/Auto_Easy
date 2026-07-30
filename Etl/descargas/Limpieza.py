import os
import time
from pathlib import Path

# Raíz del proyecto (carpeta que contiene Etl/), independiente del cwd desde
# donde se lance el proceso (importante para Docker/servicios/systemd).
BASE_DIR = Path(__file__).resolve().parent.parent.parent


def limpiar_excels_antiguos(
    carpeta=None,
    horas_maximas=12
):
    """
    Elimina archivos .xlsx antiguos de una carpeta.

    Args:
        carpeta (str): Ruta de la carpeta. Si no se pasa, se resuelve desde
            la env var EASYSALES_DOWNLOAD_DIR, o si no está definida, relativa
            a la raíz del proyecto (permite apuntar a un volumen montado en
            Docker sin tocar el código).
        horas_maximas (int): Tiempo máximo permitido en horas.
    """
    carpeta = carpeta or os.environ.get(
        "EASYSALES_DOWNLOAD_DIR", str(BASE_DIR / "descargas")
    )

    # Convertir horas a segundos
    tiempo_limite = horas_maximas * 60 * 60

    # Tiempo actual
    ahora = time.time()

    # Verificar si la carpeta existe
    if not os.path.exists(carpeta):
        return {
            "status": "error",
            "mensaje": f"La carpeta no existe: {carpeta}"
        }

    archivos_eliminados = []
    errores = []

    # Buscar archivos Excel
    for archivo in Path(carpeta).glob("*.xlsx"):

        # Obtener última modificación
        fecha_modificacion = archivo.stat().st_mtime

        # Calcular antigüedad
        antiguedad = ahora - fecha_modificacion

        # Eliminar si supera el límite
        if antiguedad > tiempo_limite:
            try:
                os.remove(archivo)

                archivos_eliminados.append(archivo.name)

                print(f"Archivo eliminado: {archivo.name}")

            except Exception as e:
                errores.append({
                    "archivo": archivo.name,
                    "error": str(e)
                })

                print(f"Error eliminando {archivo.name}: {e}")

    return {
        "status": "ok",
        "archivos_eliminados": archivos_eliminados,
        "errores": errores,
        "mensaje": "Limpieza finalizada."
    }