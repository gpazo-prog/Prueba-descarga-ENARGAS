import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def wait_for_downloads(download_dir, expected_count, timeout=90):
    print(f"Esperando a que el archivo llegue a la carpeta (Timeout: {timeout}s)...")
    for i in range(timeout):
        files = os.listdir(download_dir)
        # Si hay un archivo temporal descargándose, seguimos esperando
        if any(f.endswith('.crdownload') or f.endswith('.tmp') for f in files):
            time.sleep(1)
            continue
            
        # Contamos cuántos archivos .xls ya están listos
        xls_files = [f for f in files if f.endswith('.xls')]
        if len(xls_files) >= expected_count:
            print(f"¡Descarga completada! Total archivos actuales: {len(xls_files)}")
            return True
            
        # Aviso en consola cada 15 segundos para saber que no se tildó
        if i % 15 == 0 and i > 0:
            print(f"Aún esperando descarga... ({i} segundos)")
            
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
    chrome_options.add_argument("--disable-gpu")
    
    # CLAVE 1: Disfrazamos al bot como si fuera un usuario real usando Windows y Chrome normal
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    print("Iniciando navegador...")
    driver = webdriver.Chrome(options=chrome_options)
    
    # CLAVE 2: Usamos el comando 'Browser' en lugar de 'Page' para forzar descargas en GitHub
    driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": download_dir,
        "eventsEnabled": True
    })
    
    wait = WebDriverWait(driver, 25)

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

        archivos_esperados = 0

        for valor_cuadro, nombre_cuadro in nombres_cuadros.items():
            print(f"\n--- Procesando: {nombre_cuadro} ---")
            archivos_esperados += 1
            
            # CLAVE 3: Bloque Try/Except. Si una descarga falla, no rompe el resto del script.
            try:
                xpath_cuadro = "//select[option[contains(text(), 'Conversiones')]]"
                select_cuadro_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_cuadro)))
                
                # Scroll para asegurar que el elemento sea interactuable
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_cuadro_element)
                time.sleep(1)
                
                cuadro = Select(select_cuadro_element)
                cuadro.select_by_value(valor_cuadro)
                time.sleep(2)

                print("Haciendo clic en 'Ver .xls'...")
                # Usamos presence en lugar de clickable por si hay algún popup invisible
                btn_xls = wait.until(EC.presence_of_element_located((By.ID, "btn-ver-xls")))
                driver.execute_script("arguments[0].click();", btn_xls)

                # Aumentamos el timeout a 90 segundos por si el servidor de Enargas está lento
                exito = wait_for_downloads(download_dir, archivos_esperados, timeout=90)
                
                if exito:
                    print(f"Pausa de seguridad de 10 segundos...")
                    time.sleep(10)
                else:
                    print("Advertencia: No se descargó el archivo a tiempo. Restando del contador...")
                    archivos_esperados -= 1

            except Exception as e_cuadro:
                print(f"Error procesando '{nombre_cuadro}': {e_cuadro}")
                archivos_esperados -= 1
                
                # Si ocurrió un error grave, refrescamos la página para resetear el estado
                print("Refrescando la página para intentar con el siguiente cuadro...")
                driver.get(url)
                time.sleep(3)
                Select(wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc")))).select_by_value("5;2")
                time.sleep(3)
                Select(wait.until(EC.presence_of_element_located((By.ID, "periodo")))).select_by_value("2026")
                time.sleep(3)

        print("\n¡Proceso finalizado!")
        print("Archivos encontrados en la carpeta de GitHub:")
        for file in os.listdir(download_dir):
            print(f"- {file}")

    except Exception as e:
        print(f"Ocurrió un error en el flujo principal:")
        print(e)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
