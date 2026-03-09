import os
import time
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def main():
    download_dir = os.path.abspath("descargas_gnc")
    os.makedirs(download_dir, exist_ok=True)
    
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    print("Iniciando Selenium para obtener sesión...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    wait = WebDriverWait(driver, 40)
    archivos_ok = 0
    
    try:
        url = "https://www.enargas.gov.ar/secciones/gas-natural-comprimido/estadisticas.php"
        driver.get(url)
        time.sleep(5)

        print("Seleccionando filtros iniciales...")
        Select(wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc")))).select_by_value("5;2")
        time.sleep(3)
        Select(wait.until(EC.presence_of_element_located((By.ID, "periodo")))).select_by_value("2026")
        time.sleep(3)

        # Extraer sesión
        session_cookies = {c['name']: c['value'] for c in driver.get_cookies()}
        user_agent = driver.execute_script("return navigator.userAgent")
        
        nombres_cuadros = {"1": "Conversiones", "2": "Desmontajes", "3": "Revisiones", "4": "Modificaciones", "5": "Revisiones Cil.", "6": "Cilindro CRPC"}
        
        for valor_cuadro, nombre_cuadro in nombres_cuadros.items():
            print(f"Descargando {nombre_cuadro} vía Requests...")
            
            payload = {
                "tipo-consulta-gnc": "5;2",
                "periodo": "2026",
                "cuadro": valor_cuadro,
                "btn-ver-xls": "Ver XLS"
            }
            
            headers = {
                "User-Agent": user_agent,
                "Referer": url,
                "Origin": "https://www.enargas.gov.ar"
            }

            try:
                response = requests.post(url, data=payload, cookies=session_cookies, headers=headers, timeout=30)
                if response.status_code == 200 and len(response.content) > 1000:
                    file_path = os.path.join(download_dir, f"{nombre_cuadro}.xls")
                    with open(file_path, "wb") as f:
                        f.write(response.content)
                    archivos_ok += 1
                    print(f"✓ {nombre_cuadro} descargado.")
                else:
                    print(f"✗ Error en respuesta para {nombre_cuadro}")
            except Exception as e:
                print(f"✗ Error en petición: {e}")
            
            time.sleep(2)

    finally:
        driver.quit()
        print(f"\nResumen Final: {archivos_ok}/6")
        if archivos_ok < 6:
            exit(1)

if __name__ == '__main__':
    main()
