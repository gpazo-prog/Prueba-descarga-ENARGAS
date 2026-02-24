import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def wait_for_downloads(download_dir, expected_count, timeout=60):
    """Espera hasta que la cantidad de archivos .xls en la carpeta sea la esperada."""
    print("Esperando a que aparezca y finalice el archivo...")
    for _ in range(timeout):
        files = os.listdir(download_dir)
        # Si hay un archivo temporal descargándose, seguimos esperando
        if any(f.endswith('.crdownload') or f.endswith('.tmp') for f in files):
            time.sleep(1)
            continue
            
        # Contamos cuántos archivos .xls ya están listos
        xls_files = [f for f in files if f.endswith('.xls')]
        if len(xls_files) >= expected_count:
            print(f"Descarga completada. Total archivos actuales: {len(xls_files)}")
            return True
            
        time.sleep(1)
        
    print("Tiempo de espera agotado para la descarga.")
    return False

def main():
    download_dir = os.path.join(os.getcwd(), "descargas_gnc")
    os.makedirs(download_dir, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    print("Iniciando navegador...")
    driver = webdriver.Chrome(options=chrome_options)
    
    # ¡CRUCIAL PARA HEADLESS! Fuerza el permiso de descargas a nivel interno
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": download_dir
    })
    
    wait = WebDriverWait(driver, 20)

    try:
        url = "https://www.enargas.gov.ar/secciones/gas-natural-comprimido/estadisticas.php"
        print(f"Accediendo a: {url}")
        driver.get(url)

        print("Seleccionando Tipo de Estadística...")
        tipo_estadistica = Select(wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc"))))
        tipo_estadistica.select_by_value("5;2")
        time.sleep(3)

        print("Seleccionando Periodo...")
        periodo = Select(wait.until(EC.presence_of_element_located((By.ID, "periodo"))))
        periodo.select_by_value("2026")
        time.sleep(3)

        nombres_cuadros = {
            "1": "Conversiones de vehículos",
            "2": "Desmontajes de equipos en vehículos",
            "3": "Revisiones periódicas de vehículos",
            "4": "Modificaciones de equipos en vehículos",
            "5": "Revisiones de Cilindros",
            "6": "Cilindro de GNC revisiones CRPC"
        }

        # Llevamos la cuenta de cuántos archivos deberíamos tener
        archivos_esperados = 0

        for valor_cuadro, nombre_cuadro in nombres_cuadros.items():
            print(f"\n--- Procesando: {nombre_cuadro} ---")
            archivos_esperados += 1
            
            # Buscamos el elemento select del cuadro en cada iteración para evitar errores de elementos "caducados"
            # Usamos un XPath robusto: "El select que contiene una opción con el texto 'Conversiones'"
            xpath_cuadro = "//select[option[contains(text(), 'Conversiones')]]"
            select_cuadro_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_cuadro)))
            
            # Hacemos scroll por si el elemento quedó fuera de pantalla
            driver.execute_script("arguments[0].scrollIntoView();", select_cuadro_element)
            time.sleep(1)
            
            cuadro = Select(select_cuadro_element)
            cuadro.select_by_value(valor_cuadro)
            time.sleep(2)

            print("Haciendo clic en 'Ver .xls'...")
            btn_xls = wait.until(EC.element_to_be_clickable((By.ID, "btn-ver-xls")))
            driver.execute_script("arguments[0].click();", btn_xls)

            # Esperamos a que la carpeta tenga 'archivos_esperados' cantidad de archivos .xls
            exito = wait_for_downloads(download_dir, archivos_esperados)
            
            if exito:
                print(f"Pausa de seguridad de 10 segundos...")
                time.sleep(10)
            else:
                print("Hubo un problema con la descarga actual, intentando continuar...")

        print("\n¡Proceso finalizado!")
        print("Archivos en la carpeta:")
        for file in os.listdir(download_dir):
            print(f"- {file}")

    except Exception as e:
        print(f"Ocurrió un error en el flujo principal:")
        print(e)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
