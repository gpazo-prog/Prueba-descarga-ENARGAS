import os
import time
import base64
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

    print("Iniciando Navegador...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    wait = WebDriverWait(driver, 40)
    archivos_ok = 0
    
    try:
        url = "https://www.enargas.gov.ar/secciones/gas-natural-comprimido/estadisticas.php"
        driver.get(url)
        time.sleep(5)

        print("Configurando filtros...")
        Select(wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc")))).select_by_value("5;2")
        time.sleep(2)
        Select(wait.until(EC.presence_of_element_located((By.ID, "periodo")))).select_by_value("2026")
        time.sleep(2)

        nombres_cuadros = {"1": "Conversiones", "2": "Desmontajes", "3": "Revisiones", "4": "Modificaciones", "5": "Revisiones Cil.", "6": "Cilindro CRPC"}
        
        # SCRIPT DE JS PARA CAPTURAR EL EXCEL EN MEMORIA
        js_download_script = """
        var callback = arguments[arguments.length - 1];
        var formData = new FormData();
        formData.append('tipo-consulta-gnc', '5;2');
        formData.append('periodo', '2026');
        formData.append('cuadro', arguments[0]);
        formData.append('btn-ver-xls', 'Ver XLS');

        fetch(window.location.href, {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) throw new Error('Network response was not ok');
            return response.blob();
        })
        .then(blob => {
            var reader = new FileReader();
            reader.onloadend = function() {
                callback(reader.result.split(',')[1]);
            };
            reader.readAsDataURL(blob);
        })
        .catch(err => callback("ERROR: " + err));
        """

        for valor_cuadro, nombre_cuadro in nombres_cuadros.items():
            print(f"Extrayendo {nombre_cuadro}...")
            
            # Ejecutamos el script asíncrono y esperamos el base64
            base64_data = driver.execute_async_script(js_download_script, valor_cuadro)
            
            if base64_data.startswith("ERROR"):
                print(f"✗ Fallo en JS: {base64_data}")
                continue

            excel_bytes = base64.b64decode(base64_data)
            
            if len(excel_bytes) < 5000:
                print(f"✗ El servidor devolvió HTML en lugar de Excel para {nombre_cuadro}")
                continue

            file_path = os.path.join(download_dir, f"{nombre_cuadro}.xls")
            with open(file_path, "wb") as f:
                f.write(excel_bytes)
            
            archivos_ok += 1
            print(f"✓ {nombre_cuadro} guardado ({len(excel_bytes)} bytes).")
            time.sleep(1)

    finally:
        driver.quit()
        print(f"\nResumen Final: {archivos_ok}/6")
        if archivos_ok < 6:
            exit(1)

if __name__ == '__main__':
    main()
