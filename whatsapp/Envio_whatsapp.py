import os
import time
import pythoncom
import win32com.client as win32
import pywhatkit as kit
from datetime import datetime
import pandas as pd
import inspect
from PIL import Image
import re

# ================= CONFIGURACIÓN =================
ARCHIVO_EXCEL = r"C:/Users/Janus I5/Desktop/easygestion/Plantilla/Ventas_Marzo_Informe_Automatico.xlsx"

# Diccionario que asigna cada hoja a un destino (puede ser número de teléfono individual o ID de grupo)
# Ejemplo: "Dashboard General": "+573207590982" (individual) o "Mi Grupo": "1234567890-123456@g.us" (grupo)
'''DESTINOS_POR_HOJA = {
    "Dashboard General": "+573207590982",  # Cambia por el destino real
    "Detalle Asesores": "+573211112222",
    "Sanin": "+573213334444",
    "Casa Luker": "+573215556666",
    "Colombina": "+573217778888",
    "Tecnoquimicas": "+573219990000",
    "Brinsa": "+573221112222",
    "Colombina Helados": "+573223334444",
    "Alimentos Polar": "+573225556666"
}'''

DESTINOS_POR_HOJA = {
    "Dashboard General": "Cd7jmsC2Ksg4ERDIS1KilR",  # Cambia por el destino real
    "Sanin": "Ls1TVXkC7Gw8NJ7ltI0djv",
}
# Si alguna hoja no está en el diccionario, se usará este destino por defecto (opcional)
DESTINO_POR_DEFECTO = None  # o "+573207590982"

ENVIO_PROGRAMADO = False
ANIO, MES, DIA, HORA, MINUTO = 2026, 7, 1, 11, 15
# =================================================

# ---------- FUNCIÓN DE ENVÍO (con destino como parámetro) ----------
def enviar_imagen_whatsapp(imagen_path, destino, caption=""):
    """
    Envía una imagen por WhatsApp al destino (número individual o ID de grupo).
    Comprime la imagen si pesa más de 2MB.
    """
    try:
        # COMPRESIÓN si pesa más de 2 MB
        tamaño_bytes = os.path.getsize(imagen_path)
        if tamaño_bytes > 2 * 1024 * 1024:
            print(f"   Imagen pesa {tamaño_bytes/(1024*1024):.1f} MB. Comprimiendo...")
            with Image.open(imagen_path) as img:
                if img.width > 1200:
                    img.thumbnail((1200, 1200))
                img.save(imagen_path, "PNG", optimize=True)
            nuevo_tamaño = os.path.getsize(imagen_path)
            print(f"   Nueva imagen: {nuevo_tamaño/(1024*1024):.1f} MB")

        # Inspección de parámetros para compatibilidad
        sig = inspect.signature(kit.sendwhats_image)
        params = list(sig.parameters.keys())

        # Envío con wait_time=60 y cierre de pestaña
        if 'phone_no' in params:
            kit.sendwhats_image(phone_no=destino, img_path=imagen_path,
                                caption=caption, wait_time=60, tab_close=True)
        elif 'phone_number' in params:
            kit.sendwhats_image(phone_number=destino, img_path=imagen_path,
                                caption=caption, wait_time=60, tab_close=True)
        else:
            kit.sendwhats_image(destino, imagen_path, caption=caption,
                                wait_time=60, tab_close=True)

        print(f" Imagen enviada correctamente a {destino}.")
        return True

    except Exception as e:
        print(f" Error detallado al enviar imagen: {type(e).__name__} - {e}")
        return False

# ---------- FUNCIONES AUXILIARES (captura, validación) ----------
def capturar_hoja_como_imagen(ruta_excel, nombre_hoja, ruta_imagen_salida_absoluta):
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    try:
        wb = excel.Workbooks.Open(ruta_excel)
        ws = wb.Sheets(nombre_hoja)
        ws.UsedRange.CopyPicture(Format=2)
        chart = wb.Charts.Add()
        chart.Paste()
        chart.Export(ruta_imagen_salida_absoluta, "PNG")
        chart.Delete()
        wb.Close(SaveChanges=False)
        print(f" Captura guardada: {os.path.basename(ruta_imagen_salida_absoluta)}")
    except Exception as e:
        print(f" Error capturando '{nombre_hoja}': {e}")
        raise
    finally:
        excel.Quit()
        excel = None
        pythoncom.CoUninitialize()
        time.sleep(3)
    return ruta_imagen_salida_absoluta

def validar_hojas(archivo, lista_hojas):
    try:
        xls = pd.ExcelFile(archivo)
        hojas_existentes = xls.sheet_names
        hojas_validas = []
        for hoja in lista_hojas:
            if hoja in hojas_existentes:
                hojas_validas.append(hoja)
            else:
                print(f" ADVERTENCIA: La hoja '{hoja}' NO existe en el Excel. Se omitirá.")
        return hojas_validas
    except Exception as e:
        print(f" Error al leer el Excel: {e}")
        return []

# ---------- FUNCIÓN PRINCIPAL QUE PROCESA TODAS LAS HOJAS ----------
def enviar_imagenes():
    if not os.path.exists(ARCHIVO_EXCEL):
        print(f" No se encontró el archivo: {ARCHIVO_EXCEL}")
        return

    # No necesitamos la lista fija, usamos las claves del diccionario
    hojas_a_procesar = list(DESTINOS_POR_HOJA.keys())
    hojas_validas = validar_hojas(ARCHIVO_EXCEL, hojas_a_procesar)
    if not hojas_validas:
        print(" No hay hojas válidas para enviar.")
        return
    print(f" Hojas a procesar: {hojas_validas}")

    # Prueba de conectividad (opcional, pero puede saltar si ya confías)
    print(" Realizando prueba de conectividad con WhatsApp...")
    try:
        # Envía un mensaje de prueba al primer destino (o a un destino fijo)
        primer_destino = next(iter(DESTINOS_POR_HOJA.values()))
        kit.sendwhatmsg(primer_destino, "Prueba de conectividad", 
                        datetime.now().hour, datetime.now().minute + 2)
        print(" Prueba de conectividad exitosa.")
    except Exception as e:
        print(f" Error en prueba de conectividad: {e}")
        # Podemos continuar de todas formas
        # return  # descomentar si quieres detenerte

    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    if not directorio_actual:
        directorio_actual = os.getcwd()

    for hoja in hojas_validas:
        # Obtener destino para esta hoja
        destino = DESTINOS_POR_HOJA.get(hoja, DESTINO_POR_DEFECTO)
        if not destino:
            print(f"  No hay destino definido para '{hoja}'. Saltando...")
            continue

        try:
            # Limpiar nombre para archivo
            nombre_imagen = re.sub(r'[^a-zA-Z0-9_]', '_', hoja) + ".png"
            img_path_abs = os.path.join(directorio_actual, nombre_imagen)

            # Capturar hoja como imagen
            capturar_hoja_como_imagen(ARCHIVO_EXCEL, hoja, img_path_abs)

            time.sleep(3)
            if not os.path.exists(img_path_abs):
                print(f" La imagen no se creó: {img_path_abs}")
                continue

            # Enviar imagen con reintentos (máximo 3)
            max_intentos = 3
            for intento in range(1, max_intentos + 1):
                print(f"   Intento {intento} de {max_intentos} para enviar '{nombre_imagen}' a {destino}...")
                if enviar_imagen_whatsapp(img_path_abs, destino, f" {hoja} - Julio 2026"):
                    break
                if intento < max_intentos:
                    print("   Esperando 30 segundos antes de reintentar...")
                    time.sleep(30)
            else:
                print(f" Fallaron todos los intentos para esta imagen.")
                continue

            time.sleep(30)  # Pausa entre hojas

        except Exception as e:
            print(f" Error CRÍTICO con '{hoja}': {e}")
            time.sleep(10)

# ---------- PUNTO DE ENTRADA ----------
if __name__ == "__main__":
    if ENVIO_PROGRAMADO:
        ahora = datetime.now()
        objetivo = datetime(ANIO, MES, DIA, HORA, MINUTO)
        if objetivo <= ahora:
            print(" La hora programada ya pasó. Enviando inmediatamente...")
            enviar_imagenes()
        else:
            segundos_espera = (objetivo - ahora).total_seconds()
            print(f" Esperando {segundos_espera/60:.0f} minutos hasta las {HORA:02d}:{MINUTO:02d}")
            time.sleep(segundos_espera)
            enviar_imagenes()
    else:
        enviar_imagenes()
    print(" Proceso finalizado.")