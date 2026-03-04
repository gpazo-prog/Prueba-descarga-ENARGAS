import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

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
    print(f"Directorio de descargas: {download_dir}")

    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    
    # USER-AGENT REAL
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_settings.automatic_downloads": 1,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    print("Iniciando Chrome con User-Agent de Windows...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    wait = WebDriverWait(driver, 40)
    archivos_descargados = 0
    
    def preparar_pagina():
        url = "https://www.enargas.gov.ar/secciones/gas-natural-comprimido/estadisticas.php"
        print(f"Preparando página: {url}")
        driver.get(url)
        time.sleep(5)
        
        print("Seleccionando Tipo de Consulta...")
        tipo = wait.until(EC.element_to_be_clickable((By.ID, "tipo-consulta-gnc")))
        Select(tipo).select_by_value("5;2")
        time.sleep(4)

        print("Seleccionando Año 2026...")
        periodo = wait.until(EC.element_to_be_clickable((By.ID, "periodo")))
        Select(periodo).select_by_value("2026")
        time.sleep(4)

    try:
        preparar_pagina()
        nombres_cuadros = {"1": "Conversiones", "2": "Desmontajes", "3": "Revisiones", "4": "Modificaciones", "5": "Revisiones Cil.", "6": "Cilindro CRPC"}
        
        for valor_cuadro, nombre_cuadro in nombres_cuadros.items():
            print(f"\n--- Iniciando: {nombre_cuadro} (ID: {valor_cuadro}) ---")
            
            try:
                try:
                    sel_element = wait.until(EC.presence_of_element_located((By.XPATH, "//select[option[@value='1']]")))
                    Select(sel_element).select_by_value(valor_cuadro)
                except:
                    print("DOM no responde, refrescando página...")
                    preparar_pagina()
                    sel_element = wait.until(EC.presence_of_element_located((By.XPATH, "//select[option[@value='1']]")))
                    Select(sel_element).select_by_value(valor_cuadro)
                
                time.sleep(3)
                
                btn_xls = wait.until(EC.element_to_be_clickable((By.ID, "btn-ver-xls")))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_xls)
                time.sleep(2)

                print(f"Clic en descarga para {nombre_cuadro}...")
                actions = ActionChains(driver)
                actions.move_to_element(btn_xls).click().perform()
                
                if wait_for_downloads(download_dir, archivos_descargados + 1, timeout=50):
                    archivos_descargados += 1
                    print(f"ÉXITO: {nombre_cuadro} guardado.")
                else:
                    print(f"FALLO: El servidor no inició la descarga de {nombre_cuadro}.")
                    driver.save_screenshot(f"error_descarga_{valor_cuadro}.png")
                    preparar_pagina()

            except Exception as e:
                print(f"Error procesando {nombre_cuadro}: {str(e)}")
                driver.save_screenshot(f"fallo_excepcion_{valor_cuadro}.png")
                preparar_pagina()

    except Exception as e:
        print(f"ERROR GENERAL: {str(e)}")
        driver.save_screenshot("error_general.png")
    finally:
        driver.quit()
        print(f"\nResumen: {archivos_descargados}/6 archivos descargados.")
        if archivos_descargados < 6:
            exit(1)

if __name__ == '__main__':
    main()
