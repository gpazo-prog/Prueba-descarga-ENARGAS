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
        with open("debug_inicial.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)

        print("Listando elementos para diagnóstico...")
        try:
            inputs = driver.find_elements(By.TAG_NAME, "input")
            for i in inputs:
                print(f"  - Input: name='{i.get_attribute('name')}', id='{i.get_attribute('id')}', type='{i.get_attribute('type')}', value='{i.get_attribute('value')[:15]}...'")
            
            buttons = driver.find_elements(By.TAG_NAME, "button")
            for b in buttons:
                print(f"  - Button: text='{b.text}', id='{b.get_attribute('id')}'")
        except:
            pass

        print("Configurando filtros...")
        try:
            select_tipo = wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc")))
            Select(select_tipo).select_by_value("5;2")
            print("✓ Tipo consulta seleccionado.")
            time.sleep(2)
            
            select_periodo = wait.until(EC.presence_of_element_located((By.ID, "periodo")))
            Select(select_periodo).select_by_value("2026")
            print("✓ Periodo seleccionado.")
            time.sleep(2)
        except Exception as e:
            print(f"! Error configurando filtros: {e}")

        # Intentar clickear el botón de consulta para activar la sesión/token
        print("Intentando activar consulta...")
        try:
            # Intentar varios selectores para el botón
            selectors = [
                (By.ID, "enviar-consulta-gnc"),
                (By.XPATH, "//input[@id='enviar-consulta-gnc']"),
                (By.XPATH, "//input[@type='button' and contains(@value, 'Consultar')]"),
                (By.XPATH, "//button[contains(text(), 'Consultar')]"),
                (By.CSS_SELECTOR, "input[type='button']"),
                (By.CSS_SELECTOR, "button[type='submit']")
            ]
            
            btn = None
            for sel_type, sel_val in selectors:
                try:
                    btn = driver.find_element(sel_type, sel_val)
                    if btn:
                        print(f"✓ Botón encontrado con: {sel_val}")
                        break
                except:
                    continue
            
            if btn:
                driver.execute_script("arguments[0].scrollIntoView();", btn)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", btn) # Click vía JS es más fiable
                print("✓ Consulta activada.")
                time.sleep(5)
            else:
                print("! No se encontró botón de consulta por ningún método.")
        except Exception as e:
            print(f"! Error crítico al intentar activar consulta: {e}")

        # Intentar extraer el token del formulario con reintentos
        token = ""
        print("Buscando token...")
        for _ in range(10):
            try:
                # 1. Buscar por nombre
                token_el = driver.find_element(By.NAME, "token")
                token = token_el.get_attribute("value")
                if token:
                    print(f"✓ Token encontrado por nombre: {token[:10]}...")
                    break
            except:
                pass
            
            try:
                # 2. Buscar en cualquier input hidden que tenga un valor largo (típico de tokens)
                hiddens = driver.find_elements(By.XPATH, "//input[@type='hidden']")
                for h in hiddens:
                    val = h.get_attribute("value")
                    if val and len(val) > 20:
                        token = val
                        print(f"✓ Token probable encontrado en hidden: {token[:10]}...")
                        break
                if token: break
            except:
                pass
                
            time.sleep(2)
        
        if not token:
            # Búsqueda desesperada en el HTML
            import re
            match = re.search(r'name="token"\s+value="([^"]+)"', driver.page_source)
            if not match:
                match = re.search(r"""token['"]\s*:\s*['"]([^'"]+)['"]""", driver.page_source)
            
            if match:
                token = match.group(1)
                print(f"✓ Token encontrado en código fuente: {token[:10]}...")
            else:
                print("✗ No se encontró el campo 'token'. Guardando estado final.")
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
