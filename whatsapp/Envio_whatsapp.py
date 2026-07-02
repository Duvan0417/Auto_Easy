import os
import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import tempfile
import shutil


WASAPI_EMAIL = "valen.hidalgo.munoz@gmail.com"
WASAPI_PASSWORD = "Dizfranco2026"   
WASAPI_BASE = "https://api.wasapi.io"  
# La URL de login puede variar. Asumimos /v1/auth/login
LOGIN_URL = f"{WASAPI_BASE}/v1/auth/login"
SEND_DOCUMENT_URL = f"{WASAPI_BASE}/v1/messages/document"
GROUPS_URL = f"{WASAPI_BASE}/v1/groups"

# ------------------------------------------------------------
# Autenticación
# ------------------------------------------------------------
def login_wasapi():
    """Devuelve el token de acceso (bearer)"""
    payload = {
        "email": WASAPI_EMAIL,
        "password": WASAPI_PASSWORD
    }
    resp = requests.post(LOGIN_URL, json=payload)
    if resp.status_code != 200:
        raise Exception(f"Error login wasapi: {resp.text}")
    data = resp.json()
    # Ajusta según la respuesta real (puede ser "access_token", "token", etc.)
    token = data.get("token") or data.get("access_token")
    if not token:
        raise Exception("No se encontró token en la respuesta")
    return token

# ------------------------------------------------------------
# Obtener lista de grupos (útil para mapear)
# ------------------------------------------------------------
def get_groups(token):
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(GROUPS_URL, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Error obteniendo grupos: {resp.text}")
    return resp.json()  # Lista de grupos con id, name, etc.

# ------------------------------------------------------------
# Enviar un archivo a un grupo
# ------------------------------------------------------------
def send_document_to_group(token, group_id, file_path, caption=""):
    headers = {"Authorization": f"Bearer {token}"}
    with open(file_path, "rb") as f:
        files = {
            "file": (os.path.basename(file_path), f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        }
        data = {
            "groupId": group_id,
            "caption": caption
        }
        resp = requests.post(SEND_DOCUMENT_URL, headers=headers, data=data, files=files)
    if resp.status_code not in (200, 201):
        raise Exception(f"Error enviando documento: {resp.text}")
    return resp.json()

# ------------------------------------------------------------
# Extraer una hoja como un Excel independiente (solo valores)
# ------------------------------------------------------------
def extract_sheet_as_excel(original_excel_path, sheet_name, output_dir=None):
    """
    Lee la hoja especificada, convierte las fórmulas a valores,
    y guarda un nuevo archivo Excel que contiene únicamente esa hoja.
    Retorna la ruta del archivo creado.
    """
    # Leer con pandas (ya convierte fórmulas a valores)
    df = pd.read_excel(original_excel_path, sheet_name=sheet_name, dtype=str)
    # También queremos conservar el formato básico (anchos de columna, etc. - opcional)
    # Para simplificar, guardamos solo los datos.
    if output_dir is None:
        output_dir = tempfile.gettempdir()
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"{sheet_name}.xlsx")
    
    # Guardar usando openpyxl para mejor control
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output_path

# ------------------------------------------------------------
# Enviar todas las hojas según mapeo
# ------------------------------------------------------------
def send_all_sheets(original_excel_path, sheet_group_mapping):
    """
    sheet_group_mapping: dict { "nombre_hoja": "groupId", ... }
    """
    token = login_wasapi()
    results = {}
    for sheet_name, group_id in sheet_group_mapping.items():
        try:
            # Extraer hoja a archivo temporal
            temp_file = extract_sheet_as_excel(original_excel_path, sheet_name)
            # Enviar
            resp = send_document_to_group(token, group_id, temp_file, caption=f"Reporte {sheet_name} - {pd.Timestamp.now().strftime('%d/%m/%Y')}")
            results[sheet_name] = {"status": "ok", "response": resp}
            # Limpiar archivo temporal
            os.unlink(temp_file)
        except Exception as e:
            results[sheet_name] = {"status": "error", "error": str(e)}
    return results

# ------------------------------------------------------------
# Ejemplo de uso directo (para pruebas)
# ------------------------------------------------------------
if __name__ == "__main__":
    # Primero, lista tus grupos para conocer sus IDs
    token = login_wasapi()
    grupos = get_groups(token)
    print("Grupos disponibles:")
    for g in grupos:
        print(f"  {g['name']} -> {g['id']}")
    
    # Luego define tu mapeo (debes ajustarlo)
    mapping = {
        "Sanin": "ID_DEL_GRUPO_SANIN",
        "Casa Luker": "ID_DEL_GRUPO_CASA_LUKER",
        "Colombina": "ID_DEL_GRUPO_COLOMBINA",
        "Tecnoquimicas": "ID_DEL_GRUPO_TECNOQUIMICAS",
        "Clientes Especiales": "ID_DEL_GRUPO_CLIENTES_ESPECIALES",
        "Brinsa": "ID_DEL_GRUPO_BRINSA",
        "Colombina Helados": "ID_DEL_GRUPO_COLOMBINA_HELADOS",
        "Alimentos Polar": "ID_DEL_GRUPO_ALIMENTOS_POLAR",
        # Opcional: Dashboard General, Detalle Asesores, Ranking Asesores
    }
    
    # Ruta de tu informe actualizado (ajústala)
    informe_path = "C:/Users/Janus I5/Desktop/easygestion/Plantilla/Ventas_Marzo_Informe_Automatico.xlsx"
    
    resultado = send_all_sheets(informe_path, mapping)
    print(resultado)