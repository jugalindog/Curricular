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
from selenium.common.exceptions import StaleElementReferenceException

# ================= CONFIGURACIÓN =================

ARCHIVO_EXCEL = '/home/jugalindog/Pasantia/Curricular/Curricular/Historial_Academica/activos/listado (Copy).xlsx'
NOMBRE_COLUMNA = 'Documento'
CARPETA_DESCARGA = "/home/jugalindog/Documents/Historias academicas/activos"

USUARIO = "jugalindog"
CONTRASENA = "Wz6Np3So8" 
URL_SISTEMA = "https://hrepsia.unal.edu.co/xmlpserver/Produccion/Componentes/Historia_Academica/RE_EST_HCA.xdo?_xpt=0&_xmode=2&_xt=PORTADA"
TIEMPO_MAXIMO_ESPERA_DESCARGA = 600 

# =================================================

def normalizar_texto(texto):
    """
    Convierte cualquier entrada (int, float, str) a un string limpio.
    Ejemplo: 12345.0 -> '12345', ' 12345 ' -> '12345'
    """
    texto = str(texto).strip()
    if texto.endswith('.0'):
        texto = texto[:-2]
    return texto

def obtener_archivos_ya_descargados(carpeta):
    """Devuelve un SET con los IDs normalizados de archivos existentes"""
    archivos = glob.glob(os.path.join(carpeta, "*.pdf"))
    ids_encontrados = set()
    for f in archivos:
        # Solo nombre sin ruta ni extension
        nombre_limpio = os.path.splitext(os.path.basename(f))[0]
        ids_encontrados.add(normalizar_texto(nombre_limpio))
    return ids_encontrados

def cerrar_pestanas_extra(driver_instance, main_window_handle):
    """Cierra pestañas adicionales y regresa a la principal"""
    try:
        pestañas = driver_instance.window_handles
        if len(pestañas) > 1:
            for pestana in pestañas:
                if pestana != main_window_handle:
                    driver_instance.switch_to.window(pestana)
                    driver_instance.close()
            driver_instance.switch_to.window(main_window_handle)
    except Exception:
        driver_instance.switch_to.window(main_window_handle)

def esperar_y_obtener_nuevo_archivo(carpeta, timeout=600):
    start_time = time.time()
    print(f"   -> Esperando archivo (Timeout: {timeout}s)...")
    
    while (time.time() - start_time) < timeout:
        # Verificar .crdownload
        if glob.glob(os.path.join(carpeta, "*.crdownload")):
            time.sleep(1); continue
            
        # Buscar el más reciente
        archivos = glob.glob(os.path.join(carpeta, "*.pdf"))
        if archivos:
            mas_reciente = max(archivos, key=os.path.getctime)
            # Debe ser creado después de que empezamos esta función (-2 seg margen)
            if os.path.getctime(mas_reciente) > (start_time - 2):
                if os.path.getsize(mas_reciente) > 0:
                    return mas_reciente
        time.sleep(2)
    return None

def proceso_principal():
    # 1. Configuración Navegador
    if not os.path.exists(CARPETA_DESCARGA): os.makedirs(CARPETA_DESCARGA)
    
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": CARPETA_DESCARGA,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_settings.popups": 0
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--network-idle-timeout=600000")

    # 2. Carga de Excel y Análisis de Carpeta (ANTES DEL NAVEGADOR)
    print("--- 1. ANALIZANDO ARCHIVOS EXISTENTES ---")
    try:
        df = pd.read_excel(ARCHIVO_EXCEL, dtype=str) # Forzamos lectura como texto
    except Exception as e:
        print(f"Error leyendo Excel: {e}"); return

    ya_descargados = obtener_archivos_ya_descargados(CARPETA_DESCARGA)
    print(f"   -> Archivos PDF encontrados en carpeta: {len(ya_descargados)}")
    if len(ya_descargados) > 0:
        ejemplo = list(ya_descargados)[0]
        print(f"   -> Ejemplo de archivo detectado: '{ejemplo}'")
    else:
        print("   -> ⚠️ OJO: No se detectaron archivos previos en la carpeta.")

    # 3. Filtrado del DataFrame
    # Creamos una columna temporal normalizada para comparar
    df['id_normalizado'] = df[NOMBRE_COLUMNA].apply(normalizar_texto)
    
    # Filtramos: Solo los que NO están en ya_descargados
    # El caracter ~ invierte la selección (significa NOT)
    df_pendientes = df[~df['id_normalizado'].isin(ya_descargados)]
    
    total_original = len(df)
    total_pendientes = len(df_pendientes)
    omitidos = total_original - total_pendientes
    
    print(f"\n--- RESUMEN DE TAREA ---")
    print(f"   Total en Excel: {total_original}")
    print(f"   Ya descargados: {omitidos} (Se omitirán)")
    print(f"   A descargar:    {total_pendientes}")
    print("------------------------\n")

    if total_pendientes == 0:
        print("¡Todo está descargado! Finalizando.")
        return

    # 4. Iniciar Navegador y Login
    print("Iniciando navegador...")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 30, ignored_exceptions=[StaleElementReferenceException])
    
    # Login
    driver.get(URL_SISTEMA)
    try:
        wait.until(EC.element_to_be_clickable((By.ID, "id"))).clear()
        driver.find_element(By.ID, "id").send_keys(USUARIO)
        driver.find_element(By.ID, "passwd").send_keys(CONTRASENA)
        driver.find_element(By.XPATH, "//input[@value='Sign In']").click()
        time.sleep(5)
    except Exception as e:
        print(f"Error Login: {e}"); driver.quit(); return

    ventana_principal = driver.current_window_handle

    # 5. Bucle de descarga (Solo iteramos sobre los pendientes)
    contador = 0
    for index, fila in df_pendientes.iterrows():
        contador += 1
        documento_id = fila['id_normalizado'] # Usamos el ID limpio
        
        print(f"[{contador}/{total_pendientes}] Procesando ID: {documento_id}")

        try:
            # Asegurar foco
            if driver.current_window_handle != ventana_principal:
                driver.switch_to.window(ventana_principal)
            
            driver.get(URL_SISTEMA)

            # Navegación
            wait.until(EC.element_to_be_clickable((By.ID, "_paramsnum_documento"))).send_keys(documento_id)
            driver.find_element(By.ID, "reportViewApply").click()
            time.sleep(2) 

            # Verificar si existen resultados antes de intentar descargar
            # (Opcional, pero ayuda si el ID no existe)
            
            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@title='REPORTE']"))).click()
            time.sleep(2)

            # Descargar
            wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@alt='Actions']"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Export')]"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'PDF')]"))).click()
            
            # Esperar archivo
            archivo = esperar_y_obtener_nuevo_archivo(CARPETA_DESCARGA, TIEMPO_MAXIMO_ESPERA_DESCARGA)
            
            if archivo:
                destino = os.path.join(CARPETA_DESCARGA, f"{documento_id}.pdf")
                
                # Intentar borrar si existe un corrupto previo
                if os.path.exists(destino): os.remove(destino)
                
                os.rename(archivo, destino)
                print(f"   -> ✅ OK: {documento_id}.pdf")
                
                # Cerrar pestañas extra inmediatamente
                cerrar_pestanas_extra(driver, ventana_principal)
            else:
                print(f"   -> ❌ Error descarga o timeout.")
                cerrar_pestanas_extra(driver, ventana_principal)

        except Exception as e:
            print(f"   -> 💥 Error en {documento_id}: {e}")
            cerrar_pestanas_extra(driver, ventana_principal)
            continue

    print("\nProceso finalizado.")
    driver.quit()

if __name__ == "__main__":
    proceso_principal()