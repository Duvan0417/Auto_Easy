import time
import random
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import undetected_chromedriver as uc
from selenium_stealth import stealth

# ================= CONFIGURACIÓN =================
USERNAME = "web.sistemas.dzf"
PASSWORD = "Easynet123"
BASE_URL = "https://easysales.com.co"
PROXY = None

# =========== FUNCIONES DE COMPORTAMIENTO HUMANO ===========
def human_delay(min_sec=0.3, max_sec=1.2):
    time.sleep(random.uniform(min_sec, max_sec))

def human_type(element, text, delay_between_keys=0.05, random_variation=0.03):
    element.clear()
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(delay_between_keys - random_variation,
                                   delay_between_keys + random_variation))

def human_click(driver, element):
    actions = ActionChains(driver)
    actions.move_to_element_with_offset(element, random.randint(-5, 5), random.randint(-5, 5))
    actions.pause(random.uniform(0.1, 0.3))
    actions.click()
    actions.perform()

def scroll_to_element(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
    human_delay(0.4, 0.9)

# =========== FUNCIÓN PARA CREAR DRIVER (NUEVO CADA VEZ) ===========
def create_driver():
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    if PROXY:
        options.add_argument(f'--proxy-server={PROXY}')
    
    driver = uc.Chrome(options=options, version_main=148)
    stealth(driver,
            languages=["es-ES", "es"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True)
    return driver

# =========== EJECUCIÓN DEL FLUJO ===========
def run_flow(driver, wait, actions):
    try:
        print("1. Abriendo URL...")
        driver.get(BASE_URL + "/easygestioneasysalesv3/")
        human_delay(1.5, 2.5)

        print("2. Configurando ventana...")
        driver.set_window_size(1550, 830)
        human_delay(0.5, 1.0)

        print("3. Ingresando usuario...")
        username_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".username")))
        scroll_to_element(driver, username_field)
        human_click(driver, username_field)
        human_type(username_field, USERNAME, delay_between_keys=0.06, random_variation=0.02)
        human_delay(0.5, 1.0)

        print("4. Ingresando contraseña...")
        password_field = driver.find_element(By.ID, "userPassword")
        scroll_to_element(driver, password_field)
        human_click(driver, password_field)
        human_type(password_field, PASSWORD)
        human_delay(0.5, 1.0)

        print("5. Haciendo login...")
        login_btn = driver.find_element(By.CSS_SELECTOR, ".btn")
        scroll_to_element(driver, login_btn)
        human_click(driver, login_btn)
        human_delay(2.0, 3.5)

        print("6. Clic en icono Ventas...")
        ventas_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".icon-Ventas-01")))
        scroll_to_element(driver, ventas_icon)
        human_click(driver, ventas_icon)
        human_delay(0.8, 1.5)

        print("7. Submenú Ventas y pedidos...")
        ventas_pedidos = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#accordion-body-5 > .ng-star-inserted:nth-child(1) .text")))
        scroll_to_element(driver, ventas_pedidos)
        human_click(driver, ventas_pedidos)
        human_delay(1.0, 2.0)

        print("8. Clic en filtro...")
        filtro = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".filters-box:nth-child(4) > .form-control")))
        scroll_to_element(driver, filtro)
        human_click(driver, filtro)
        human_delay(0.5, 1.0)

        print("9. Checkbox thead...")
        check_thead = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "thead .width_no_cliente .easynet-check_icon")))
        scroll_to_element(driver, check_thead)
        human_click(driver, check_thead)
        human_delay(0.5, 1.0)

        print("10. Botón siguiente página...")
        next_page_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".mr-auto")))
        scroll_to_element(driver, next_page_btn)
        human_click(driver, next_page_btn)
        human_delay(0.8, 1.5)

        print("11-15. Seleccionando 5 registros...")
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ng-star-inserted .width_no_cliente .easynet-check_icon")))
        registros_selectores = [
            ".ng-star-inserted:nth-child(1) > .width_no_cliente .easynet-check_icon",
            ".ng-star-inserted:nth-child(2) > .width_no_cliente .easynet-check_icon",
            ".ng-star-inserted:nth-child(3) > .width_no_cliente .easynet-check_icon",
            ".ng-star-inserted:nth-child(4) .easynet-check_icon",
            ".ng-star-inserted > .width_no_cliente > .active > .easynet-check_icon"
        ]
        for idx, selector in enumerate(registros_selectores, 1):
            try:
                registro = driver.find_element(By.CSS_SELECTOR, selector)
                scroll_to_element(driver, registro)
                human_click(driver, registro)
                print(f"   Registro {idx} seleccionado")
                human_delay(0.3, 0.7)
            except Exception as e:
                print(f"Error seleccionando {selector}: {e}")

        print("16. Aceptar diálogo...")
        aceptar_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".mat-dialog-actions > .modal-button-acept")))
        scroll_to_element(driver, aceptar_btn)
        aceptar_btn.click()
        print("   Diálogo aceptado")
        human_delay(1.5, 2.5)

        print("17. Búsqueda definitiva...")
        buscar_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn_search_definitive")))
        scroll_to_element(driver, buscar_btn)
        buscar_btn.click()
        human_delay(1.5, 2.5)

        print("18. MouseOver/MouseOut en menú trigger...")
        menu_trigger = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".mat-menu-trigger")))
        actions.move_to_element(menu_trigger).perform()
        human_delay(0.3, 0.7)
        actions.move_by_offset(0, 0).perform()
        human_delay(0.3, 0.7)

        print("19. Exportar a Excel - clic en icono...")
        excel_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".icon-file-excel")))
        excel_icon.click()
        human_delay(0.5, 1.0)

        print("20. MouseOver en icono Excel...")
        actions.move_to_element(excel_icon).perform()
        human_delay(0.3, 0.7)

        print("21. MouseOut (0,0)...")
        actions.move_by_offset(0, 0).perform()
        human_delay(0.3, 0.7)

        print("22. Esperando elemento .cdk-focused > span...")
        menu_item = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".cdk-focused > span")))
        actions.move_to_element(menu_item).perform()
        human_delay(0.4, 0.8)

        print("23. Clic en opción de reporte...")
        menu_item.click()
        human_delay(0.5, 1.0)

        print("24. MouseOut final (0,0)...")
        actions.move_by_offset(0, 0).perform()

        print("Esperando 10 segundos para completar descarga...")
        time.sleep(10)
        print("Flujo completado exitosamente.")
        return True

    except Exception as e:
        print(f"ERROR en el paso actual: {e}")
        driver.save_screenshot("error_screenshot.png")
        print("Captura de pantalla guardada como 'error_screenshot.png'")
        return False

# =========== REINTENTOS CON NUEVO DRIVER ===========
MAX_RETRIES = 2
driver = None
for attempt in range(MAX_RETRIES):
    print(f"\n--- Intento {attempt+1} de {MAX_RETRIES} ---")
    driver = create_driver()
    wait = WebDriverWait(driver, 20)
    actions = ActionChains(driver)
    
    if run_flow(driver, wait, actions):
        break
    else:
        print(f"Intento {attempt+1} fallido. Cerrando navegador...")
        driver.quit()
        time.sleep(random.uniform(3, 6))
else:
    print("Se agotaron los reintentos. El proceso falló.")
    if driver:
        driver.quit()