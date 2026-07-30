import os
import random
import time
from pathlib import Path

from playwright.sync_api import (
    Locator,
    Page,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)
from playwright_stealth import Stealth

# ================= CONFIGURACIÓN =================
USERNAME = os.environ.get("EASYSALES_USERNAME", "web.sistemas.dzf")
PASSWORD = os.environ.get("EASYSALES_PASSWORD", "Easynet123")
BASE_URL = os.environ.get("EASYSALES_BASE_URL", "https://easysales.com.co")
HEADLESS = os.environ.get("EASYSALES_HEADLESS", "false").lower() == "true"
PROXY = os.environ.get("EASYSALES_PROXY")  # ej: "http://host:puerto"

DEFAULT_TIMEOUT_MS = 20_000
DOWNLOAD_TIMEOUT_MS = 60_000
MAX_RETRIES = 2

# Raíz del proyecto (carpeta que contiene Bot/), independiente del cwd desde
# donde se lance el proceso (importante para Docker/servicios/systemd).
BASE_DIR = Path(__file__).resolve().parent.parent

# Carpeta donde se guardarán las descargas (override con env var para Docker,
# p. ej. montar un volumen en /data/descargas)
DOWNLOAD_DIR = Path(os.environ.get("EASYSALES_DOWNLOAD_DIR", BASE_DIR / "descargas"))
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

# Archivo donde se guardará la ruta del último Excel descargado
LAST_EXCEL_FILE = DOWNLOAD_DIR / "ultimo_excel.txt"

ERROR_SCREENSHOT = Path(
    os.environ.get("EASYSALES_ERROR_SCREENSHOT", BASE_DIR / "error_screenshot.png")
)

_STEALTH = Stealth(
    navigator_languages_override=("es-ES", "es"),
    navigator_vendor_override="Google Inc.",
    navigator_platform_override="Win32",
    webgl_vendor_override="Intel Inc.",
    webgl_renderer_override="Intel Iris OpenGL Engine",
)


def human_type(locator: Locator, text: str, delay_ms: int = 60) -> None:
    """Escribe carácter a carácter con una demora variable, simulando tecleo humano."""
    locator.click()
    locator.fill("")
    locator.press_sequentially(text, delay=delay_ms + random.randint(-20, 20))


def create_browser(playwright):
    launch_args = ["--disable-blink-features=AutomationControlled"]
    launch_kwargs = {"headless": HEADLESS, "args": launch_args}
    if PROXY:
        launch_kwargs["proxy"] = {"server": PROXY}

    browser = playwright.chromium.launch(**launch_kwargs)
    context = browser.new_context(
        viewport={"width": 1550, "height": 830},
        locale="es-ES",
        accept_downloads=True,
    )
    _STEALTH.apply_stealth_sync(context)
    context.set_default_timeout(DEFAULT_TIMEOUT_MS)

    page = context.new_page()
    return browser, context, page


def run_flow(page: Page) -> bool:
    try:
        print("1. Abriendo URL...")
        page.goto(f"{BASE_URL}/easygestioneasysalesv3/")

        print("2. Ingresando usuario...")
        username_field = page.locator(".username")
        human_type(username_field, USERNAME)

        print("3. Ingresando contraseña...")
        password_field = page.locator("#userPassword")
        human_type(password_field, PASSWORD)

        print("4. Haciendo login...")
        page.locator(".btn").click()

        print("5. Clic en icono Ventas...")
        page.locator(".icon-Ventas-01").click()

        print("6. Submenú Ventas y pedidos...")
        page.locator("#accordion-body-5 > .ng-star-inserted:nth-child(1) .text").click()

        print("7. Clic en filtro...")
        page.locator(".filters-box:nth-child(4) > .form-control").click()

        print("8. Checkbox thead...")
        page.locator("thead .width_no_cliente .easynet-check_icon").click()

        print("9. Botón siguiente página...")
        page.locator(".mr-auto").click()

        print("10-14. Seleccionando 5 registros...")
        registros_selectores = [
            ".ng-star-inserted:nth-child(1) > .width_no_cliente .easynet-check_icon",
            ".ng-star-inserted:nth-child(2) > .width_no_cliente .easynet-check_icon",
            ".ng-star-inserted:nth-child(3) > .width_no_cliente .easynet-check_icon",
            ".ng-star-inserted:nth-child(4) .easynet-check_icon",
            ".ng-star-inserted > .width_no_cliente > .active > .easynet-check_icon",
        ]
        for idx, selector in enumerate(registros_selectores, 1):
            try:
                page.locator(selector).click()
                print(f"   Registro {idx} seleccionado")
            except PlaywrightTimeoutError as e:
                print(f"Error seleccionando {selector}: {e}")

        print("15. Aceptar diálogo...")
        page.locator(".mat-dialog-actions > .modal-button-acept").click()

        print("16. Búsqueda definitiva...")
        # El diálogo "Seleccionar usuarios" deja un botón .btn_search_definitive propio
        # en el DOM mientras se cierra; se acota a <main> para evitar el duplicado.
        page.locator("main .btn_search_definitive").click()

        print("17. Abriendo menú de exportación...")
        page.locator(".mat-menu-trigger").hover()

        print("18. Exportar a Excel - clic en icono...")
        excel_icon = page.locator(".icon-file-excel")
        excel_icon.hover()
        excel_icon.click()

        print("19. Seleccionando opción de reporte...")
        menu_item = page.locator(".cdk-focused > span")
        menu_item.hover()

        print("20. Descargando reporte...")
        with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as download_info:
            menu_item.click()
        download = download_info.value

        downloaded_file = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
        download.save_as(downloaded_file)
        print(f"Archivo descargado: {downloaded_file}")

        with open(LAST_EXCEL_FILE, "w", encoding="utf-8") as f:
            f.write(downloaded_file)
        print(f"Ruta guardada en: {LAST_EXCEL_FILE}")

        print("Flujo completado exitosamente.")
        return True

    except Exception as e:
        print(f"ERROR en el paso actual: {e}")
        try:
            page.screenshot(path=ERROR_SCREENSHOT, full_page=True)
            print(f"Captura de pantalla guardada como '{ERROR_SCREENSHOT}'")
        except Exception as screenshot_error:
            print(f"No se pudo guardar la captura de pantalla: {screenshot_error}")
        return False


def main():
    """
    Función principal que ejecuta el flujo de descarga de Excel.
    Retorna un diccionario con el resultado.
    """
    with sync_playwright() as playwright:
        for attempt in range(1, MAX_RETRIES + 1):
            print(f"\n--- Intento {attempt} de {MAX_RETRIES} ---")
            browser, context, page = create_browser(playwright)
            try:
                if run_flow(page):
                    return {
                        "status": "success",
                        "message": "AutoEasy completado correctamente",
                        "archivo_excel": str(LAST_EXCEL_FILE),
                    }
                print(f"Intento {attempt} fallido. Cerrando navegador...")
            finally:
                context.close()
                browser.close()

            if attempt < MAX_RETRIES:
                time.sleep(random.uniform(3, 6))

    return {
        "status": "error",
        "message": "Se agotaron los reintentos. El proceso falló.",
    }


if __name__ == "__main__":
    resultado = main()
    print(resultado)
