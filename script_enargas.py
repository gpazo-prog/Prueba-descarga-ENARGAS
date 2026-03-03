import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def wait_for_downloads(download_dir, expected_count, timeout=60):
    for _ in range(timeout):
        files = os.listdir(download_dir)
        if any(f.endswith('.crdownload') or f.endswith('.tmp') for f in files):
            time.sleep(2)
            continue
        xls_files = [f for f in files if f.endswith('.xls')]
        if len(xls_files) >= expected_count:
            time.sleep(2)
            return True
        time.sleep(1)
    return False

def main():
    download_dir = os.path.abspath("descargas_gnc")
    os.makedirs(download_dir, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    
    # STEALTH MODE: Ocultar automatización
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    # Ocultar navigator.webdriver
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })

    wait = WebDriverWait(driver, 30)
    try:
        driver.get("https://www.enargas.gov.ar/secciones/gas-natural-comprimido/estadisticas.php")
        Select(wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc")))).select_by_value("5;2")
        time.sleep(2)
        Select(wait.until(EC.presence_of_element_located((By.ID, "periodo")))).select_by_value("2026")
        
        nombres_cuadros = {"1": "Conversiones", "2": "Desmontajes", "3": "Revisiones", "4": "Modificaciones", "5": "Revisiones Cil.", "6": "Cilindro CRPC"}
        archivos_ok = 0
        for val, name in nombres_cuadros.items():
            sel = wait.until(EC.presence_of_element_located((By.XPATH, "//select[option[@value='1']]")))
            Select(sel).select_by_value(val)
            time.sleep(2)
            btn = wait.until(EC.element_to_be_clickable((By.ID, "btn-ver-xls")))
            driver.execute_script("arguments[0].click();", btn)
            if wait_for_downloads(download_dir, archivos_ok + 1):
                archivos_ok += 1
                print(f"Descargado: {name}")
    finally:
        driver.quit()
