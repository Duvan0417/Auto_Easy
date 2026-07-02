import pandas as pd
import os

ruta = "C:/Users/Janus I5/Desktop/TotalControl prueba.xlsx"

df= pd.read_excel(ruta)
df.to_csv("C:/Users/Janus I5/Desktop/TotalControl prueba.csv", index=False, sep=',', encoding='utf-8-sig')

print("Archivo CSV generado.")