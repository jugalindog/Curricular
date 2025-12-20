import time
import os
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

# ================= CONFIGURACIÓN ROBUSTA =================

# 1. Rutas
ARCHIVO_EXCEL = '/home/jugalindog/Pasantia/Curricular/Curricular/Historial_Academica/activos/listado.xlsx'
NOMBRE_COLUMNA = 'Documento'
CARPETA_DESCARGA = "/home/jugalindog/Documents/Historias academicas/activos"

# 2. Credenciales
USUARIO = "jugalindog"
CONTRASENA = "Wz6Np3So8" 

# 3. URL
URL_SISTEMA = "https://hrepsia.unal.edu.co/xmlpserver/Produccion/Componentes/Historia_Academica/RE_EST_HCA.xdo?_xpt=0&_xmode=2&_xt=PORTADA"

# 4. TIEMPOS (Adaptado a tu servidor lento)
# Tiempo máximo que esperaremos por una descarga (10 minutos = 600 segundos)
TIEMPO_MAXIMO_ESPERA_DESCARGA = 600 

# =========================================================

# Verificar carpeta
if not os.path.exists(CARPETA_DESCARGA):
    os.makedirs(CARPETA_DESCARGA)

# Configuración Navegador
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": CARPETA_DESCARGA,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True, # Importante para que no abra visor PDF
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)
# Agregamos argumentos para evitar timeouts de red
chrome_options.add_argument("--network-idle-timeout=600000") 

print("Iniciando navegador...")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
# Timeout para encontrar elementos (botones, cajas de texto)
wait = WebDriverWait(driver, 30, ignored_exceptions=[StaleElementReferenceException])

def contar_archivos(carpeta):
    """Cuenta cuántos archivos hay en la carpeta actualmente"""
    return len(glob.glob(os.path.join(carpeta, "*")))

def esperar_y_obtener_nuevo_archivo(carpeta, num_archivos_inicial, timeout=600):
    """
    Espera inteligentemente hasta que aparezca un archivo nuevo y termine de descargarse.
    Retorna la ruta del nuevo archivo o None si falla.
    """
    start_time = time.time()
    print(f"   -> Esperando generación del reporte (Máx {timeout/60} min)...")
    
    while (time.time() - start_time) < timeout:
        # 1. Verificar si hay archivos temporales (.crdownload)
        temp_files = glob.glob(os.path.join(carpeta, "*.crdownload"))
        if temp_files:
            # Si hay un .crdownload, es que ESTÁ descargando. Esperamos.
            print("   -> Descargando en progreso...", end='\r')
            time.sleep(2)
            continue
            
        # 2. Verificar si el número de archivos aumentó
        archivos_actuales = glob.glob(os.path.join(carpeta, "*.pdf"))
        # Filtramos para asegurarnos de tomar el más reciente
        if archivos_actuales:
            archivo_mas_reciente = max(archivos_actuales, key=os.path.getctime)
            
            # Si el archivo más reciente fue creado DESPUÉS de que empezamos a esperar
            if os.path.getctime(archivo_mas_reciente) > start_time:
                # Verificamos que tenga tamaño (que no esté vacío)
                if os.path.getsize(archivo_mas_reciente) > 0:
                    return archivo_mas_reciente
        
        # Feedback visual para que no pienses que se trabó
        tiempo_transcurrido = int(time.time() - start_time)
        if tiempo_transcurrido % 10 == 0: # Imprime cada 10 seg
            print(f"   -> Esperando servidor... ({tiempo_transcurrido}s transcurridos)", end='\r')
        
        time.sleep(2) # Revisamos la carpeta cada 2 segundos

    return None

def realizar_login():
    print("--- INICIANDO SESIÓN ---")
    driver.get(URL_SISTEMA)
    try:
        wait.until(EC.element_to_be_clickable((By.ID, "id"))).clear()
        driver.find_element(By.ID, "id").send_keys(USUARIO)
        driver.find_element(By.ID, "passwd").send_keys(CONTRASENA)
        driver.find_element(By.XPATH, "//input[@value='Sign In']").click()
        print("   -> Datos enviados. Esperando acceso...")
        time.sleep(5) 
    except Exception as e:
        print(f"ERROR LOGIN: {e}")

def proceso_principal():
    try:
        df = pd.read_excel(ARCHIVO_EXCEL, dtype=str)
        total = len(df)
        print(f"Objetivo: Procesar {total} documentos.")
    except Exception as e:
        print(f"Error Excel: {e}"); return

    realizar_login()

    for index, fila in df.iterrows():
        documento_id = fila[NOMBRE_COLUMNA].strip()
        if not documento_id or str(documento_id).lower() == 'nan': continue
        
        print(f"\n[{index+1}/{total}] Procesando: {documento_id}")

        try:
            driver.get(URL_SISTEMA)

            # --- PASO 1 y 2: NAVEGACIÓN ---
            wait.until(EC.element_to_be_clickable((By.ID, "_paramsnum_documento"))).send_keys(documento_id)
            driver.find_element(By.ID, "reportViewApply").click()
            time.sleep(3) # Espera carga inicial tabla

            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@title='REPORTE']"))).click()
            time.sleep(2)

            # --- PASO 3: SOLICITAR DESCARGA ---
            wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@alt='Actions']"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Export')]"))).click()
            
            # Tomamos el tiempo justo antes de dar clic final
            # Esto nos ayuda a filtrar archivos viejos
            tiempo_inicio_clic = time.time()
            
            # CLIC FINAL
            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'PDF')]"))).click()
            
            # --- PASO 4: ESPERA INTELIGENTE (LA MAGIA) ---
            # Aquí el código se quedará esperando hasta 10 minutos
            archivo_descargado = esperar_y_obtener_nuevo_archivo(
                CARPETA_DESCARGA, 
                0, # No necesitamos conteo previo si usamos timestamp
                timeout=TIEMPO_MAXIMO_ESPERA_DESCARGA
            )
            
            print(" " * 60, end='\r') # Limpiar linea de consola

            if archivo_descargado:
                ruta_final = os.path.join(CARPETA_DESCARGA, f"{documento_id}.pdf")
                
                # Renombrar
                if archivo_descargado != ruta_final:
                    if os.path.exists(ruta_final): os.remove(ruta_final)
                    # Loop pequeño de reintento de renombrado (a veces el antivirus bloquea el archivo 1 seg)
                    renombrado_ok = False
                    for _ in range(5):
                        try:
                            os.rename(archivo_descargado, ruta_final)
                            renombrado_ok = True
                            break
                        except:
                            time.sleep(1)
                    
                    if renombrado_ok:
                        print(f"   -> ✅ DESCARGA COMPLETADA: {documento_id}.pdf")
                    else:
                        print(f"   -> ⚠️ ERROR AL RENOMBRAR. Archivo quedó como: {os.path.basename(archivo_descargado)}")
                else:
                    print(f"   -> ✅ Archivo verificado.")
            else:
                print(f"   -> ❌ ERROR: Tiempo agotado (10 min) o fallo en descarga.")

        except Exception as e:
            print(f"   -> 💥 FALLO CRÍTICO en {documento_id}: {e}")
            continue

    # Auditoría final
    total_archivos = len(glob.glob(os.path.join(CARPETA_DESCARGA, "*.pdf")))
    print(f"\n--- REPORTE FINAL: {total_archivos} de {total} procesados ---")
    driver.quit()

if __name__ == "__main__":
    proceso_principal()