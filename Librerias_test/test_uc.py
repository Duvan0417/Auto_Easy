import undetected_chromedriver as uc

options = uc.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--no-sandbox")
options.add_argument("--disable-blink-features=AutomationControlled")

# 👇 Especifica tu versión de Chrome (148)
driver = uc.Chrome(options=options, version_main=148)

driver.get("https://www.google.com")
print("✅ Driver funcionando correctamente")
input("Presiona Enter para cerrar...")
driver.quit()