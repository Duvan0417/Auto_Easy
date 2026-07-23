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

# Diccionario: hoja -> destino (número o ID de grupo)
DESTINOS_POR_HOJA = {
    "Dashboard General": "KsrzDWeRDpLGK0TTJiJItV",       # Ejemplo: número individual
    "Detalle Asesores": "KsrzDWeRDpLGK0TTJiJItV",
    "Sanin": "IO17qwaVukn1DscvMz4aUF",
    "Casa Luker": "C40Txf4yFsh1S35rvcqQwR",
    "Colombina": "JyTY9GhkmRt7hAbme3lsam",
    "Tecnoquimicas": "Krgd8fPVbTXJjZXAtvji2I",
    "Brinsa": "EPYMNhYTeDwBNo68Mi8eLa",
    "Colombina Helados": "+573207590982",
    "Alimentos Polar": "IUPJUOEsz6C6F9p9tpke9d"
    # Si usas grupos, pon el ID en formato "1234567890-123456@g.us"
}

DESTINO_POR_DEFECTO = None
ENVIO_PROGRAMADO = False
ANIO, MES, DIA, HORA, MINUTO = 2026, 7, 1, 11, 15

# Factor de escala real para la resolución de captura (aplicado al VECTOR, no a la imagen ya rasterizada)
# 2.0-3.0 suele ser suficiente para texto muy nítido. Valores muy altos pueden hacer
# que Excel tarde más o falle al exportar; si eso pasa, baja el valor.
ESCALA_CAPTURA = {
    "Detalle Asesores": 3.0
}

# Configuración de compresión en el envío (None = sin redimensionar)
CONFIG_COMPRESION = {
    "Detalle Asesores": {"max_width": None, "max_height": None}  # Sin compresión
}
# =================================================

def enviar_imagen_whatsapp(imagen_path, destino, caption="", max_width=1200, max_height=1200):
    """
    Envía una imagen por WhatsApp.
    Si max_width y max_height son None, no se redimensiona.
    """
    try:
        tamaño_bytes = os.path.getsize(imagen_path)
        # Solo redimensionar si se especifican límites y el archivo pesa >2MB
        if max_width is not None and max_height is not None and tamaño_bytes > 2 * 1024 * 1024:
            print(f"   Imagen pesa {tamaño_bytes/(1024*1024):.1f} MB. Redimensionando...")
            with Image.open(imagen_path) as img:
                ancho_orig, alto_orig = img.size
                nuevo_ancho, nuevo_alto = ancho_orig, alto_orig

                if max_width and ancho_orig > max_width:
                    factor = max_width / ancho_orig
                    nuevo_ancho = max_width
                    nuevo_alto = int(alto_orig * factor)

                if max_height and nuevo_alto > max_height:
                    factor = max_height / nuevo_alto
                    nuevo_alto = max_height
                    nuevo_ancho = int(nuevo_ancho * factor)

                if nuevo_ancho != ancho_orig or nuevo_alto != alto_orig:
                    img = img.resize((nuevo_ancho, nuevo_alto), Image.Resampling.LANCZOS)

                img.save(imagen_path, "PNG", compress_level=0)
            print(f"   Nueva imagen: {os.path.getsize(imagen_path)/(1024*1024):.1f} MB")
        else:
            if tamaño_bytes > 2 * 1024 * 1024:
                print(f"   Imagen pesa {tamaño_bytes/(1024*1024):.1f} MB (sin redimensionar).")
                with Image.open(imagen_path) as img:
                    img.save(imagen_path, "PNG", compress_level=0)

        sig = inspect.signature(kit.sendwhats_image)
        params = list(sig.parameters.keys())
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
        print(f" Error al enviar imagen: {type(e).__name__} - {e}")
        return False


def capturar_hoja_como_imagen(ruta_excel, nombre_hoja, ruta_imagen_salida_absoluta, escala=1.0):
    """
    Captura una hoja de Excel como imagen PNG en ALTA RESOLUCIÓN REAL.

    Claves del cambio respecto a la versión anterior:
    1. Se copia con Format=xlPicture (vectorial/EMF) en vez de xlBitmap.
       Un bitmap ya viene "congelado" a la resolución de pantalla; un vector
       se puede volver a dibujar más grande sin perder nitidez.
    2. Se pega en un ChartObject EMBEBIDO (no una hoja de gráfico completa),
       lo que permite fijar Width/Height en puntos de forma explícita.
       Al multiplicar esas dimensiones por `escala`, Excel vuelve a renderizar
       el contenido a ese tamaño mayor -> más píxeles reales, no interpolados.
    """
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    chart_obj = None
    try:
        excel.Visible = False
        wb = excel.Workbooks.Open(ruta_excel)
        ws = wb.Sheets(nombre_hoja)
        excel.ActiveWindow.Zoom = 100

        used_range = ws.UsedRange
        ancho_pts = used_range.Width
        alto_pts = used_range.Height

        # xlScreen = 1, xlPicture = -4147 (vectorial, NO bitmap)
        used_range.CopyPicture(Appearance=1, Format=-4147)

        # ChartObject embebido con tamaño explícito ampliado por la escala
        chart_obj = ws.ChartObjects().Add(
            0, 0,
            ancho_pts * escala,
            alto_pts * escala
        )
        chart = chart_obj.Chart
        chart.Paste()

        # IMPORTANTE: al pegar, la imagen entra a su tamaño ORIGINAL (no al
        # tamaño del lienzo que acabamos de ampliar). Si no la reescalamos,
        # queda una tabla pequeña dentro de un canvas grande en blanco.
        # Tomamos la forma recién pegada y la estiramos para llenar todo
        # el lienzo; al ser vectorial, se redibuja nítida en el tamaño nuevo.
        forma = chart.Shapes(chart.Shapes.Count)
        forma.LockAspectRatio = False
        forma.Left = 0
        forma.Top = 0
        forma.Width = ancho_pts * escala
        forma.Height = alto_pts * escala

        chart.Export(ruta_imagen_salida_absoluta, "PNG")

        wb.Close(SaveChanges=False)
        print(f" Captura guardada en alta resolución (escala {escala}x): "
              f"{os.path.basename(ruta_imagen_salida_absoluta)}")
    except Exception as e:
        print(f" Error capturando '{nombre_hoja}': {e}")
        raise
    finally:
        try:
            if chart_obj is not None:
                chart_obj.Delete()
        except Exception:
            pass
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
                print(f" ADVERTENCIA: La hoja '{hoja}' NO existe. Se omitirá.")
        return hojas_validas
    except Exception as e:
        print(f" Error al leer el Excel: {e}")
        return []


def enviar_imagenes():
    if not os.path.exists(ARCHIVO_EXCEL):
        print(f" No se encontró el archivo: {ARCHIVO_EXCEL}")
        return

    hojas_a_procesar = list(DESTINOS_POR_HOJA.keys())
    hojas_validas = validar_hojas(ARCHIVO_EXCEL, hojas_a_procesar)
    if not hojas_validas:
        print(" No hay hojas válidas para enviar.")
        return
    print(f" Hojas a procesar: {hojas_validas}")

    print(" Realizando prueba de conectividad con WhatsApp...")
    try:
        primer_destino = next(iter(DESTINOS_POR_HOJA.values()))
        if primer_destino.startswith('+'):
            kit.sendwhatmsg(primer_destino, "Prueba de conectividad",
                            datetime.now().hour, datetime.now().minute + 2)
            print(" Prueba de conectividad exitosa.")
        else:
            print(" El destino no es un número de teléfono, omitiendo prueba.")
    except Exception as e:
        print(f" Error en prueba de conectividad (se continúa igual): {e}")

    directorio_actual = os.path.dirname(os.path.abspath(__file__)) or os.getcwd()

    for hoja in hojas_validas:
        destino = DESTINOS_POR_HOJA.get(hoja, DESTINO_POR_DEFECTO)
        if not destino:
            print(f"  No hay destino para '{hoja}'. Saltando...")
            continue

        try:
            nombre_imagen = re.sub(r'[^a-zA-Z0-9_]', '_', hoja) + ".png"
            img_path_abs = os.path.join(directorio_actual, nombre_imagen)

            escala = ESCALA_CAPTURA.get(hoja, 1.0)

            # Captura ya en alta resolución real (vectorial + escala aplicada por Excel)
            capturar_hoja_como_imagen(ARCHIVO_EXCEL, hoja, img_path_abs, escala=escala)
            time.sleep(3)

            if not os.path.exists(img_path_abs):
                print(f" La imagen no se creó: {img_path_abs}")
                continue

            config = CONFIG_COMPRESION.get(hoja, {})
            max_width = config.get("max_width", 1200)
            max_height = config.get("max_height", 1200)

            for intento in range(1, 4):
                print(f"   Intento {intento} para enviar '{nombre_imagen}' a {destino}...")
                if enviar_imagen_whatsapp(img_path_abs, destino, f" {hoja} - Julio 2026",
                                          max_width=max_width, max_height=max_height):
                    break
                if intento < 3:
                    print("   Esperando 30s antes de reintentar...")
                    time.sleep(30)
            else:
                print(f" Fallaron todos los intentos para '{hoja}'.")
                continue

            time.sleep(30)

        except Exception as e:
            print(f" Error CRÍTICO con '{hoja}': {e}")
            time.sleep(10)


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