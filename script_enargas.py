import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def wait_for_downloads(download_dir, timeout=60):
    """Espera hasta que no haya archivos temporales de descarga (.crdownload)"""
    print("Esperando a que finalice la descarga...")
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        is_downloading = any(fname.endswith('.crdownload') or fname.endswith('.tmp') for fname in os.listdir(download_dir))
        if not is_downloading:
            # Damos un segundo extra para asegurar que el archivo se renombró correctamente a .xls
            time.sleep(1)
            print("Descarga completada.")
            return True
        seconds += 1
    print("Tiempo de espera agotado para la descarga.")
    return False

def main():
    # 1. Configurar la carpeta de descargas
    download_dir = os.path.join(os.getcwd(), "descargas_gnc")
    os.makedirs(download_dir, exist_ok=True)

    # 2. Configurar Chrome en modo Headless (sin interfaz gráfica, necesario para GitHub)
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    
    # Preferencias para forzar la descarga en la carpeta deseada sin preguntar
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    print("Iniciando navegador...")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 15)

    try:
        url = "https://www.enargas.gov.ar/secciones/gas-natural-comprimido/estadisticas.php"
        print(f"Accediendo a: {url}")
        driver.get(url)

        # 3. Seleccionar "Tipo de estadistica" -> '5;2'
        print("Seleccionando Tipo de Estadística...")
        tipo_estadistica = Select(wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc"))))
        tipo_estadistica.select_by_value("5;2")
        time.sleep(3) # Esperar a que la página actualice los otros menús mediante AJAX

        # 4. Seleccionar "Periodo" -> '2026'
        print("Seleccionando Periodo...")
        periodo = Select(wait.until(EC.presence_of_element_located((By.ID, "periodo"))))
        periodo.select_by_value("2026")
        time.sleep(3)

        # 5. Iterar sobre los cuadros (Opciones del 1 al 6)
        nombres_cuadros = {
            "1": "Conversiones de vehículos",
            "2": "Desmontajes de equipos en vehículos",
            "3": "Revisiones periódicas de vehículos",
            "4": "Modificaciones de equipos en vehículos",
            "5": "Revisiones de Cilindros",
            "6": "Cilindro de GNC revisiones CRPC"
        }

        for valor_cuadro, nombre_cuadro in nombres_cuadros.items():
            print(f"\nProcesando: {nombre_cuadro}")
            
            # Seleccionar el cuadro correspondiente. (Buscamos el select visible que maneja los cuadros)
            # En la estructura de Enargas, el select suele tener name="cuadro" o id="cuadro"
            cuadro = Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[option[@value='1' and contains(., 'Conversiones')]]"))))
            cuadro.select_by_value(valor_cuadro)
            
            # Breve pausa para que se procese la selección
            time.sleep(2)

            # 6. Hacer clic en "Ver .xls"
            print("Haciendo clic en 'Ver .xls'...")
            btn_xls = wait.until(EC.element_to_be_clickable((By.ID, "btn-ver-xls")))
            # A veces el botón se obstruye, usar Javascript click es más seguro en headless
            driver.execute_script("arguments[0].click();", btn_xls)

            # 7. Esperar a que termine la descarga y pausa de seguridad
            wait_for_downloads(download_dir)
            
            print(f"Pausa de seguridad de 10 segundos antes del siguiente archivo...")
            time.sleep(10)  # Evita que el servidor te bloquee por descargas excesivamente rápidas

        print("\n¡Todas las descargas han finalizado exitosamente!")
        print("Archivos descargados:")
        for file in os.listdir(download_dir):
            print(f"- {file}")

    except Exception as e:
        print(f"Ocurrió un error: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
