import pandas as pd
import os

# Ruta donde se guardó el archivo de texto
ruta_info = "C:/Users/Janus I5/Desktop/easygestion/descargas/ultimo_excel.txt"

with open(ruta_info, "r", encoding="utf-8") as f:
    archivo_excel = f.read().strip()

# Verificar que existe
if os.path.exists(archivo_excel):
    df = pd.read_excel(archivo_excel)
    print(df.head())
else:
    print("El archivo Excel no se encontró.")
    