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
    print(f"Directorio de descargas: {download_dir}")

    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    # Preferencias de descarga ultra-agresivas
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False, # A veces bloquea archivos .xls
        "profile.default_content_settings.popups": 0,
        "profile.default_content_settings.automatic_downloads": 1,
        "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--allow-running-insecure-content")

    print("Iniciando Chrome...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    # Ocultar navigator.webdriver
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })

    wait = WebDriverWait(driver, 30)
    archivos_descargados = 0
    
    try:
        url = "https://www.enargas.gov.ar/secciones/gas-natural-comprimido/estadisticas.php"
        print(f"Navegando a {url}...")
        driver.get(url)
        
        driver.save_screenshot("inicio_pagina.png")

        print("Seleccionando Tipo de Consulta...")
        tipo_select = wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc")))
        Select(tipo_select).select_by_value("5;2")
        time.sleep(4) # Espera generosa para AJAX

        print("Seleccionando Periodo...")
        periodo_select = wait.until(EC.presence_of_element_located((By.ID, "periodo")))
        Select(periodo_select).select_by_value("2026")
        time.sleep(4) # Espera generosa para AJAX

        nombres_cuadros = {"1": "Conversiones", "2": "Desmontajes", "3": "Revisiones", "4": "Modificaciones", "5": "Revisiones Cil.", "6": "Cilindro CRPC"}
        
        for valor_cuadro, nombre_cuadro in nombres_cuadros.items():
            print(f"--- Cuadro {valor_cuadro}: {nombre_cuadro} ---")
            try:
                # Localizar el select de cuadros
                select_cuadro_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//select[option[@value='1']]")))
                Select(select_cuadro_element).select_by_value(valor_cuadro)
                print(f"Seleccionado cuadro {valor_cuadro} en el menú.")
                time.sleep(3) # Tiempo para que el JS del sitio actualice el botón

                # Esperar a que el botón sea clickeable y no esté oscurecido
                btn_xls = wait.until(EC.element_to_be_clickable((By.ID, "btn-ver-xls")))
                
                # Scroll al botón para asegurar visibilidad
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_xls)
                time.sleep(1)

                print("Haciendo clic en el botón de descarga...")
                driver.execute_script("arguments[0].click();", btn_xls)
                
                if wait_for_downloads(download_dir, archivos_descargados + 1, timeout=45):
                    archivos_descargados += 1
                    print(f"EXITO: {nombre_cuadro} descargado correctamente.")
                else:
                    print(f"ERROR: El archivo de {nombre_cuadro} no apareció en la carpeta.")
                    driver.save_screenshot(f"error_descarga_{valor_cuadro}.png")
                
                time.sleep(2) # Pausa entre descargas

            except Exception as e:
                print(f"Fallo en cuadro {nombre_cuadro}: {str(e)}")
                driver.save_screenshot(f"fallo_cuadro_{valor_cuadro}.png")
                # Intentar continuar con el siguiente cuadro
                continue

    except Exception as e:
        print(f"ERROR CRITICO: {str(e)}")
        driver.save_screenshot("error_critico.png")
    finally:
        driver.quit()
        print(f"Fin del proceso. Total: {archivos_descargados}")
        if archivos_descargados == 0:
            exit(1)

if __name__ == '__main__':
    main()
