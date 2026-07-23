import os
import time
import pythoncom
import win32com.client as win32
import pywhatkit as kit
from datetime import datetime
import pandas as pd
import inspect
from PIL import Image

# ================= CONFIGURACIÓN =================
ARCHIVO_EXCEL = r"C:/Users/Janus I5/Desktop/easygestion/Plantilla/Ventas_Marzo_Informe_Automatico.xlsx"
NUMERO_DESTINO = "+573207590982"
HOJAS_A_ENVIAR = ["Dashboard General", "Detalle Asesores", "Sanin", "Casa Luker", 
                  "Colombina", "Tecnoquimicas", "Brinsa", "Colombina Helados", "Alimentos Polar"]

ENVIO_PROGRAMADO = False
ANIO, MES, DIA, HORA, MINUTO = 2026, 7, 1, 11, 15
# =================================================

# ---------- NUEVA FUNCIÓN DE ENVÍO (con compresión y timeout ampliado) ----------
def enviar_imagen_whatsapp(imagen_path, caption=""):
    """
    Envía una imagen por WhatsApp comprimiéndola si pesa más de 2MB,
    con wait_time=60 segundos y cierre automático de pestaña.
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
            kit.sendwhats_image(phone_no=NUMERO_DESTINO, img_path=imagen_path,
                                caption=caption, wait_time=60, tab_close=True)
        elif 'phone_number' in params:
            kit.sendwhats_image(phone_number=NUMERO_DESTINO, img_path=imagen_path,
                                caption=caption, wait_time=60, tab_close=True)
        else:
            kit.sendwhats_image(NUMERO_DESTINO, imagen_path, caption=caption,
                                wait_time=60, tab_close=True)

        print(f" Imagen enviada correctamente.")
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

    hojas_validas = validar_hojas(ARCHIVO_EXCEL, HOJAS_A_ENVIAR)
    if not hojas_validas:
        print(" No hay hojas válidas para enviar.")
        return
    print(f" Hojas a procesar: {hojas_validas}")

    # Prueba de conectividad (opcional, pero puede saltar si ya confías)
    print(" Realizando prueba de conectividad con WhatsApp...")
    try:
        kit.sendwhatmsg(NUMERO_DESTINO, "Prueba de conectividad", 
                        datetime.now().hour, datetime.now().minute + 2)
        print(" Prueba de conectividad exitosa.")
    except Exception as e:
        print(f" Error en prueba de conectividad: {e}")
        return

    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    if not directorio_actual:
        directorio_actual = os.getcwd()

    for hoja in hojas_validas:
        try:
            # Limpiar nombre para archivo (elimina cualquier carácter raro)
            import re
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
                print(f"   Intento {intento} de {max_intentos} para enviar '{nombre_imagen}'...")
                if enviar_imagen_whatsapp(img_path_abs, f" {hoja} - Julio 2026"):
                    break
                if intento < max_intentos:
                    print("   Esperando 30 segundos antes de reintentar...")
                    time.sleep(30)
            else:
                print(f" Fallaron todos los intentos para esta imagen.")
                continue

            time.sleep(30)  # Pausa entre hojas (aumentada a 30s)

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