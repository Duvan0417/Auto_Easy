import pandas as pd
import os
from pathlib import Path

# Carpeta de origen/destino de los archivos DIZFRANCO. Por defecto se resuelve
# bajo el Desktop del usuario actual (portable entre cuentas/equipos); se puede
# apuntar a otra ubicación (p. ej. un volumen montado en Docker) con la env var.
CARPETA_DIZFRANCO = os.environ.get(
    "EASYGESTION_DIZFRANCO_DIR",
    str(Path.home() / "Desktop" / "INFORMACION DIARIA" / "DIZFRANCO"),
)

rutas = [
    {
        "input": os.path.join(CARPETA_DIZFRANCO, "Ventas.xlsx"),
        "output": os.path.join(CARPETA_DIZFRANCO, "Ventas.csv")
    },
    {
        "input": os.path.join(CARPETA_DIZFRANCO, "Devoluciones.xlsx"),
        "output": os.path.join(CARPETA_DIZFRANCO, "Devoluciones.csv")
    }
]

for ruta in rutas:
    input_excel = ruta["input"]
    output_csv = ruta["output"]

    # Verificar que el archivo existe
    if not os.path.exists(input_excel):
        print(f" Error: No se encontró el archivo '{input_excel}'")
        continue  # Saltamos al siguiente archivo en lugar de detener todo

    # Leer el archivo Excel (primera hoja)
    print(f" Leyendo '{input_excel}' ...")
    df = pd.read_excel(input_excel, sheet_name=0, dtype=str)

    # Limpiar comas en todas las celdas de tipo texto
    df_clean = df.applymap(lambda x: x.replace(',', ' ') if isinstance(x, str) else x)

    # Guardar como CSV con delimitador coma y codificación UTF-8 con BOM
    df_clean.to_csv(output_csv, index=False, sep=',', encoding='utf-8-sig')

    print(f" Archivo CSV generado: '{output_csv}'")
    print(f" Registros procesados: {len(df_clean)}\n")

print(" Proceso finalizado para todos los archivos.")