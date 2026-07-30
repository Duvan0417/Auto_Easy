import os
import time
import pythoncom
import win32com.client as win32
import pywhatkit as kit
from datetime import datetime
from pathlib import Path
import pandas as pd
import inspect
from PIL import Image
import re

# ================= CONFIGURACIÓN =================
# Raíz del proyecto (carpeta que contiene whatsapp/), independiente del cwd
# desde donde se lance el proceso. Override con EASYGESTION_INFORME_XLSX si
# el archivo vive en otra ubicación (p. ej. un volumen montado en Docker).
BASE_DIR = Path(__file__).resolve().parents[2]
ARCHIVO_EXCEL = os.environ.get(
    "EASYGESTION_INFORME_XLSX",
    str(BASE_DIR / "Plantilla" / "Ventas_Marzo_Informe_Automatico.xlsx"),
)
CAPTURAS_DIR = os.environ.get(
    "EASYGESTION_CAPTURAS_DIR",
    str(BASE_DIR / "whatsapp" / "Capturas_Excel"),
)

# Grupo temporal de pruebas: https://chat.whatsapp.com/Lnqlkg5QOLyH584xiT3e17
# El "destino" de pywhatkit es el código de invitación (lo que va tras chat.whatsapp.com/).
#GRUPO_PRUEBA = "Lnqlkg5QOLyH584xiT3e17"

# Diccionario: hoja -> destino (número o ID de grupo)
# Por ahora todas las hojas (excepto Clientes Especiales y Ranking Asesores,
# que aún no tienen grupo asignado) van a GRUPO_PRUEBA. Los destinos originales
# quedan comentados para restaurarlos cuando se reactiven los grupos individuales.
DESTINOS_POR_HOJA = {
    "Dashboard General": "KsrzDWeRDpLGK0TTJiJItV",     # antes: "KsrzDWeRDpLGK0TTJiJItV"
    "Detalle Asesores": "KsrzDWeRDpLGK0TTJiJItV",      # antes: "KsrzDWeRDpLGK0TTJiJItV"
    "Sanin": "IO17qwaVukn1DscvMz4aUF",                 # antes: "IO17qwaVukn1DscvMz4aUF"
    "Casa Luker": "C40Txf4yFsh1S35rvcqQwR",            # antes: "C40Txf4yFsh1S35rvcqQwR"
    "Colombina": "JyTY9GhkmRt7hAbme3lsam",             # antes: "JyTY9GhkmRt7hAbme3lsam"
    "Tecnoquimicas": "Krgd8fPVbTXJjZXAtvji2I",         # antes: "Krgd8fPVbTXJjZXAtvji2I"
    "Brinsa": "EPYMNhYTeDwBNo68Mi8eLa",                # antes: "EPYMNhYTeDwBNo68Mi8eLa"
    "Colombina Helados": "Lnqlkg5QOLyH584xiT3e17",     # antes: "+573207590982"
    "Alimentos Polar": "IUPJUOEsz6C6F9p9tpke9d",       # antes: "IUPJUOEsz6C6F9p9tpke9d"
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

# Configuración de compresión en el envío, por hoja.
# calidad_jpeg: 1-95 recomprime como JPEG a esa calidad (reduce peso). None = deja la
# imagen intacta (mismo formato/píxeles, sin recodificar). max_width/max_height limitan
# dimensiones antes de comprimir; None = sin límite.
# Por defecto NO se recomprime: para capturas de tablas/gráficos (colores planos, texto
# nítido) el PNG nativo de Excel ya pesa menos que cualquier JPEG probado (85/70/50 salen
# más pesados; recién a calidad ~30 gana tamaño, pero con texto visiblemente borroso).
# "Detalle Asesores" se deja intacta a propósito: necesita máxima nitidez.
CALIDAD_JPEG_DEFECTO = None
CONFIG_COMPRESION = {
    "Detalle Asesores": {"max_width": None, "max_height": None, "calidad_jpeg": None},
}
# =================================================

def enviar_imagen_whatsapp(imagen_path, destino, max_width=1200, max_height=1200,
                           calidad_jpeg=CALIDAD_JPEG_DEFECTO):
    """
    Envía una imagen por WhatsApp.
    - calidad_jpeg=1-95: se recomprime como JPEG a esa calidad antes de enviar.
    - calidad_jpeg=None y max_width/max_height definidos: si pesa más de 2MB se
      comprime con PIL.thumbnail (mismo formato PNG, conserva relación de aspecto).
      Es la lógica de Envio_whatsapp.py, ya probada sin fallos en el grupo de pruebas.
    - calidad_jpeg=None y max_width=max_height=None: se envía tal cual, sin tocarla
      (caso "Detalle Asesores": necesita máxima nitidez, ver CONFIG_COMPRESION).
    """
    try:
        ruta_envio = imagen_path
        tamaño_original = os.path.getsize(imagen_path)

        if calidad_jpeg is not None:
            with Image.open(imagen_path) as img:
                ancho, alto = img.size
                if max_width and ancho > max_width:
                    alto = int(alto * (max_width / ancho))
                    ancho = max_width
                if max_height and alto > max_height:
                    ancho = int(ancho * (max_height / alto))
                    alto = max_height
                if (ancho, alto) != img.size:
                    img = img.resize((ancho, alto), Image.Resampling.LANCZOS)

                ruta_envio = os.path.splitext(imagen_path)[0] + "_wa.jpg"
                img.convert("RGB").save(ruta_envio, "JPEG", quality=calidad_jpeg, optimize=True)
            print(f"   Recomprimida a JPEG calidad {calidad_jpeg}: "
                  f"{tamaño_original/1024:.0f} KB -> {os.path.getsize(ruta_envio)/1024:.0f} KB")

        elif (max_width or max_height) and tamaño_original > 2 * 1024 * 1024:
            print(f"   Imagen pesa {tamaño_original/(1024*1024):.1f} MB. Comprimiendo...")
            with Image.open(imagen_path) as img:
                limite = (max_width or max_height, max_height or max_width)
                if img.width > limite[0] or img.height > limite[1]:
                    img.thumbnail(limite)
                img.save(imagen_path, "PNG", optimize=True)
            print(f"   Nueva imagen: {os.path.getsize(imagen_path)/(1024*1024):.1f} MB")

        else:
            print(f"   Enviando sin recomprimir ({tamaño_original/1024/1024:.1f} MB): "
                  f"{os.path.basename(imagen_path)}")

        sig = inspect.signature(kit.sendwhats_image)
        params = list(sig.parameters.keys())
        if 'phone_no' in params:
            kit.sendwhats_image(phone_no=destino, img_path=ruta_envio,
                                caption="", wait_time=20, tab_close=True)
        elif 'phone_number' in params:
            kit.sendwhats_image(phone_number=destino, img_path=ruta_envio,
                                caption="", wait_time=20, tab_close=True)
        else:
            kit.sendwhats_image(destino, ruta_envio, caption="",
                                wait_time=20, tab_close=True)
        print(f" Imagen enviada correctamente a {destino}.")
        return True
    except Exception as e:
        print(f" Error al enviar imagen: {type(e).__name__} - {e}")
        return False


def capturar_hoja_como_imagen(ruta_excel, nombre_hoja, ruta_imagen_salida_absoluta, escala=1.0):

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    # Sin esto, cualquier prompt de Excel (guardar cambios, actualizar vínculos, etc.)
    # se queda esperando una respuesta humana que nunca llega (Visible=False) y el
    # proceso huérfano bloquea el portapapeles del sistema para TODAS las capturas
    # posteriores, no solo la que falló.
    excel.DisplayAlerts = False
    chart_obj = None
    wb = None
    try:
        excel.Visible = False
        wb = excel.Workbooks.Open(ruta_excel, UpdateLinks=0)
        ws = wb.Sheets(nombre_hoja)
        ws.Activate()  # CopyPicture con apariencia de pantalla necesita la hoja activa
        excel.ActiveWindow.Zoom = 100

        used_range = ws.UsedRange
        ancho_pts = used_range.Width
        alto_pts = used_range.Height

        # CopyPicture llena el portapapeles de forma asíncrona: si Paste() se dispara
        # antes de que el contenido esté listo, pega 0 formas (Shapes.Count == 0) y
        # Shapes(Count) revienta con "índice fuera de límites". Se reintenta con una
        # espera creciente para darle tiempo al portapapeles antes de fallar de verdad.
        MAX_INTENTOS_PEGADO = 4
        forma = None
        for intento_pegado in range(1, MAX_INTENTOS_PEGADO + 1):
            # xlScreen = 1, xlPicture = -4147 (vectorial, NO bitmap)
            used_range.CopyPicture(Appearance=1, Format=-4147)
            time.sleep(0.4 * intento_pegado)

            chart_obj = ws.ChartObjects().Add(
                0, 0,
                ancho_pts * escala,
                alto_pts * escala
            )
            chart = chart_obj.Chart
            chart.Paste()

            if chart.Shapes.Count > 0:
                forma = chart.Shapes(chart.Shapes.Count)
                break

            print(f"   Pegado vacío (intento {intento_pegado}/{MAX_INTENTOS_PEGADO}) "
                  f"para '{nombre_hoja}', reintentando...")
            chart_obj.Delete()
            chart_obj = None

        if forma is None:
            raise RuntimeError(
                f"No se pudo pegar la captura de '{nombre_hoja}' tras "
                f"{MAX_INTENTOS_PEGADO} intentos (portapapeles vacío)."
            )

        forma.LockAspectRatio = False
        forma.Left = 0
        forma.Top = 0
        forma.Width = ancho_pts * escala
        forma.Height = alto_pts * escala

        chart.Export(ruta_imagen_salida_absoluta, "PNG")
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
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
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

    directorio_capturas = CAPTURAS_DIR
    os.makedirs(directorio_capturas, exist_ok=True)

    for hoja in hojas_validas:
        destino = DESTINOS_POR_HOJA.get(hoja, DESTINO_POR_DEFECTO)
        if not destino:
            print(f"  No hay destino para '{hoja}'. Saltando...")
            continue

        try:
            nombre_imagen = re.sub(r'[^a-zA-Z0-9_]', '_', hoja) + ".png"
            img_path_abs = os.path.join(directorio_capturas, nombre_imagen)

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
            calidad_jpeg = config.get("calidad_jpeg", CALIDAD_JPEG_DEFECTO)

            for intento in range(1, 4):
                print(f"   Intento {intento} para enviar '{nombre_imagen}' a {destino}...")
                if enviar_imagen_whatsapp(img_path_abs, destino,
                                          max_width=max_width, max_height=max_height,
                                          calidad_jpeg=calidad_jpeg):
                    break
                if intento < 3:
                    print("   Esperando 15s antes de reintentar...")
                    time.sleep(15)
            else:
                print(f" Fallaron todos los intentos para '{hoja}'.")
                continue

            time.sleep(15)

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
            print(f" Esperando {segundos_espera/20:.0f} minutos hasta las {HORA:02d}:{MINUTO:02d}")
            time.sleep(segundos_espera)
            enviar_imagenes()
    else:
        enviar_imagenes()
    print(" Proceso finalizado.")