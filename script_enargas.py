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
        url = "https://www.enargas.gob.ar/secciones/gas-natural-comprimido/estadisticas.php"
        driver.get(url)
        time.sleep(5)
        
        # Guardar captura inicial para diagnóstico
        driver.save_screenshot("estado_inicial.png")

        print("Configurando filtros...")
        Select(wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc")))).select_by_value("5;2")
        time.sleep(2)
        Select(wait.until(EC.presence_of_element_located((By.ID, "periodo")))).select_by_value("2026")
        time.sleep(2)

        # Intentar clickear el botón de consulta para activar la sesión/token
        print("Intentando activar consulta...")
        try:
            # Intentar por ID, por texto o por clase
            btn = None
            try:
                btn = driver.find_element(By.ID, "enviar-consulta-gnc")
            except:
                try:
                    btn = driver.find_element(By.XPATH, "//button[contains(text(), 'Consultar')]")
                except:
                    btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            
            if btn:
                driver.execute_script("arguments[0].scrollIntoView();", btn)
                time.sleep(1)
                btn.click()
                print("✓ Consulta activada.")
                time.sleep(5)
            else:
                print("! No se encontró botón de consulta.")
        except Exception as e:
            print(f"! Error al intentar activar consulta: {e}")

        # Intentar extraer el token del formulario con reintentos
        token = ""
        for _ in range(10):
            try:
                # Buscar en todo el documento
                token_el = driver.find_element(By.NAME, "token")
                token = token_el.get_attribute("value")
                if token:
                    print(f"✓ Token encontrado: {token[:10]}...")
                    break
            except:
                # Scroll para forzar carga si es necesario
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
        
        if not token:
            # Búsqueda desesperada en el HTML
            import re
            match = re.search(r'name="token"s+value="([^"]+)"', driver.page_source)
            if match:
                token = match.group(1)
                print(f"✓ Token encontrado en HTML: {token[:10]}...")
            else:
                print("✗ No se encontró el campo 'token'. Guardando HTML para debug.")
                with open("debug_no_token.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                driver.save_screenshot("error_no_token.png")

        nombres_cuadros = {"1": "Conversiones", "2": "Desmontajes", "3": "Revisiones", "4": "Modificaciones", "5": "Revisiones Cil.", "6": "Cilindro CRPC"}
        
        # SCRIPT DE JS ACTUALIZADO CON URL RELATIVA PARA EVITAR CORS
        js_download_script = """
        var callback = arguments[arguments.length - 1];
        var token = arguments[1];
        var cuadro = arguments[0];
        
        var params = new URLSearchParams();
        params.append('tipo-consulta-gnc', '5;2');
        params.append('cuadro', cuadro);
        params.append('periodo', '2026');
        params.append('desarrollo', '0');
        params.append('Excel', '1');
        params.append('token', token);
        params.append('action', 'sicgnc_consulta_estadisticas');

        // Usamos URL absoluta con gob.ar para asegurar consistencia
        var exportUrl = 'https://www.enargas.gob.ar/secciones/gas-natural-comprimido/exportar-datos-operativos-gnc-xls-pdf-n.php';

        fetch(exportUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: params.toString()
        })
        .then(response => {
            if (!response.ok) throw new Error('Status: ' + response.status);
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
            print(f"Intentando extraer {nombre_cuadro}...")
            base64_data = driver.execute_async_script(js_download_script, valor_cuadro, token)
            
            if base64_data.startswith("ERROR"):
                print(f"✗ Error en JS: {base64_data}")
                continue

            excel_bytes = base64.b64decode(base64_data)
            
            # DIAGNÓSTICO: Si es pequeño, probablemente sea HTML
            if len(excel_bytes) < 10000:
                print(f"⚠ Contenido sospechoso para {nombre_cuadro} ({len(excel_bytes)} bytes). Guardando log...")
                with open(f"debug_{valor_cuadro}.html", "wb") as f:
                    f.write(excel_bytes)
                driver.save_screenshot(f"error_{valor_cuadro}.png")
                continue

            file_path = os.path.join(download_dir, f"{nombre_cuadro}.xls")
            with open(file_path, "wb") as f:
                f.write(excel_bytes)
            
            archivos_ok += 1
            print(f"✓ {nombre_cuadro} guardado ({len(excel_bytes)} bytes).")
            time.sleep(1)

    except Exception as e:
        print(f"FATAL ERROR: {e}")
        driver.save_screenshot("fatal_error.png")
    finally:
        driver.quit()
        print(f"\nResumen: {archivos_ok}/6 archivos descargados.")
        if archivos_ok < 6:
            print("REVISA LOS ARCHIVOS .png Y .html GENERADOS PARA VER EL BLOQUEO.")
            exit(1)

if __name__ == '__main__':
    main()
