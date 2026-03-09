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
        print(f"Navegando a {url}...")
        driver.get(url)
        time.sleep(5)
        
        # Función de diagnóstico
        def debug_screen(step_name):
            print(f"--- DIAGNÓSTICO: {step_name} ---")
            driver.save_screenshot(f"step_{step_name}.png")

        debug_screen("01_inicio")

        # Manejo de Iframes (el formulario suele estar dentro)
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for frame in iframes:
            src = frame.get_attribute("src") or ""
            if "consulta" in src or "gnc" in src:
                print(f"✓ Cambiando al iframe: {src}")
                driver.switch_to.frame(frame)
                debug_screen("02_dentro_iframe")
                break

        print("Configurando filtros...")
        try:
            # 1. Selección de Tipo
            select_tipo = wait.until(EC.presence_of_element_located((By.ID, "tipo-consulta-gnc")))
            Select(select_tipo).select_by_value("5;2")
            print("✓ Tipo seleccionado (5;2)")
            time.sleep(1)
            
            # 2. Selección de Periodo
            select_periodo = wait.until(EC.presence_of_element_located((By.ID, "periodo")))
            Select(select_periodo).select_by_value("2026")
            print("✓ Periodo seleccionado (2026)")
            time.sleep(1)
            
            # 3. Selección de Cuadro
            try:
                select_cuadro = driver.find_element(By.ID, "cuadro")
                Select(select_cuadro).select_by_value("1")
                print("✓ Cuadro inicial seleccionado (1)")
            except:
                print("! No se encontró selector de cuadro por ID 'cuadro'.")

        except Exception as e:
            print(f"! Error en filtros: {e}")

        # Localizar el botón Ver Excel (ID estático confirmado: btn-ver-xls)
        print("Localizando botón 'Ver Excel'...")
        try:
            btn_excel = wait.until(EC.visibility_of_element_located((By.ID, "btn-ver-xls")))
            print("✓ Botón 'btn-ver-xls' encontrado.")
            
            # Hacemos scroll y un clic suave para asegurar que la sesión/token se refresquen si es necesario
            driver.execute_script("arguments[0].scrollIntoView();", btn_excel)
            time.sleep(1)
            # No hacemos clic real para evitar el diálogo de "Guardar como", 
            # solo nos aseguramos de que el token esté presente en el DOM.
        except Exception as e:
            print(f"! Error localizando botón: {e}")

        # Extraer Token del input hidden confirmado
        token = ""
        try:
            token_el = driver.find_element(By.NAME, "token")
            token = token_el.get_attribute("value")
            if token:
                print(f"✓ Token extraído con éxito: {token[:15]}...")
        except:
            print("! No se encontró input 'token' por nombre. Intentando vía JS...")
            token = driver.execute_script("return document.getElementsByName('token')[0]?.value || '';")

        if not token:
            print("✗ ERROR CRÍTICO: No se pudo obtener el token de sesión.")
            driver.save_screenshot("error_token.png")
            exit(1)

        # Mapeo de cuadros para la descarga masiva
        nombres_cuadros = {
            "1": "Conversiones", 
            "2": "Desmontajes", 
            "3": "Revisiones", 
            "4": "Modificaciones", 
            "5": "Revisiones Cil.", 
            "6": "Cilindro CRPC"
        }
        
        # Script de descarga asíncrona (Fetch Tunneling)
        # Este script emula la función GenerarConsultaEstadisticasGNC_N('Excel')
        js_download_script = """
        var callback = arguments[arguments.length - 1];
        var cuadro = arguments[0];
        var token = arguments[1];
        
        var params = new URLSearchParams();
        params.append('tipo-consulta-gnc', '5;2');
        params.append('cuadro', cuadro);
        params.append('periodo', '2026');
        params.append('desarrollo', '0');
        params.append('Excel', '1');
        params.append('token', token);
        params.append('action', 'sicgnc_consulta_estadisticas');

        fetch('https://www.enargas.gob.ar/secciones/gas-natural-comprimido/exportar-datos-operativos-gnc-xls-pdf-n.php', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: params.toString()
        })
        .then(r => {
            if(!r.ok) throw new Error('Error en red: ' + r.status);
            return r.blob();
        })
        .then(blob => {
            var reader = new FileReader();
            reader.onloadend = () => callback(reader.result.split(',')[1]);
            reader.readAsDataURL(blob);
        })
        .catch(e => callback("ERROR: " + e));
        """

        for val, nom in nombres_cuadros.items():
            print(f"Procesando: {nom} (Cuadro {val})...")
            # Ejecutamos el túnel JS para obtener los bytes del archivo
            b64_data = driver.execute_async_script(js_download_script, val, token)
            
            if b64_data.startswith("ERROR"):
                print(f"  ✗ Error en descarga: {b64_data}")
            else:
                file_bytes = base64.b64decode(b64_data)
                # Si el archivo es muy pequeño, probablemente sea un error del servidor devuelto como HTML
                if len(file_bytes) < 500:
                    print(f"  ⚠ Archivo demasiado pequeño ({len(file_bytes)} bytes). Posible error de sesión.")
                else:
                    filename = f"{nom}.xls"
                    with open(os.path.join(download_dir, filename), "wb") as f:
                        f.write(file_bytes)
                    print(f"  ✓ Descargado con éxito ({len(file_bytes)} bytes)")
                    archivos_ok += 1

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
