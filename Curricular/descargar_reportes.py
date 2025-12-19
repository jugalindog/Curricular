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

# ================= CONFIGURACIÓN LISTA =================

# 1. Rutas (Tus rutas correctas)
ARCHIVO_EXCEL = '/home/jugalindog/Pasantia/Curricular/Curricular/Historial_Academica/activos/listado.xlsx'
NOMBRE_COLUMNA = 'Documento'
CARPETA_DESCARGA = "/home/jugalindog/Documents/Historias academicas/activos"

# 2. Credenciales
USUARIO = "jugalindog"
CONTRASENA = "Wz6Np3So8"  # <--- Tu contraseña ya está puesta aquí

# 3. URL
URL_SISTEMA = "https://hrepsia.unal.edu.co/xmlpserver/Produccion/Componentes/Historia_Academica/RE_EST_HCA.xdo?_xpt=0&_xmode=2&_xt=PORTADA"

# =======================================================

# Crear carpeta si no existe
if not os.path.exists(CARPETA_DESCARGA):
    try:
        os.makedirs(CARPETA_DESCARGA)
    except OSError as e:
        print(f"Error creando la carpeta de descarga: {e}")
        exit()

# Configuración del Navegador
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": CARPETA_DESCARGA,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)

print("Iniciando navegador...")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# IMPORTANTE: Agregamos una regla para ignorar errores de elementos viejos (Stale) y reintentar
wait = WebDriverWait(driver, 20, ignored_exceptions=[StaleElementReferenceException])

def obtener_ultimo_archivo_pdf(carpeta):
    """Busca el archivo PDF más reciente en la carpeta"""
    archivos = glob.glob(os.path.join(carpeta, "*.pdf"))
    if not archivos: return None
    return max(archivos, key=os.path.getctime)

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
        print(f"ERROR CRÍTICO EN LOGIN: {e}")
        input("Por favor realiza el login manualmente y presiona ENTER aquí...")

def proceso_principal():
    # Cargar Excel
    try:
        df = pd.read_excel(ARCHIVO_EXCEL, dtype=str)
        print(f"Se procesarán {len(df)} documentos.")
    except Exception as e:
        print(f"Error leyendo {ARCHIVO_EXCEL}: {e}"); return

    # Ejecutar Login
    realizar_login()

    # Bucle por cada documento
    for index, fila in df.iterrows():
        documento_id = fila[NOMBRE_COLUMNA].strip()
        
        if not documento_id or str(documento_id).lower() == 'nan': continue
        
        print(f"\n[{index+1}/{len(df)}] Procesando: {documento_id}")

        try:
            # --- CORRECCIÓN CRÍTICA: RECARGAR LA PÁGINA ---
            # Esto soluciona el error "Stale Element Reference" (fallo uno sí, uno no)
            driver.get(URL_SISTEMA)

            # --- PASO 1: BUSCAR DOCUMENTO ---
            # El 'wait' se encargará de esperar a que la página recargue y el elemento sea nuevo
            campo_busqueda = wait.until(EC.element_to_be_clickable((By.ID, "_paramsnum_documento")))
            campo_busqueda.clear()
            campo_busqueda.send_keys(documento_id)
            
            # Click Apply
            boton_apply = driver.find_element(By.ID, "reportViewApply")
            boton_apply.click()
            
            time.sleep(3) # Tu tiempo de espera personalizado

            # --- PASO 2: PESTAÑA REPORTE ---
            pestana_reporte = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@title='REPORTE']")))
            pestana_reporte.click()
            
            time.sleep(1.5) 

            # --- PASO 3: EXPORTAR ---
            # 1. Icono Engranaje 
            boton_engranaje = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@alt='Actions']")))
            boton_engranaje.click()
            
            # 2. Opción Export
            opcion_export = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Export')]")))
            opcion_export.click()
            
            # 3. Opción PDF
            opcion_pdf = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'PDF')]")))
            opcion_pdf.click()

            print("   -> Descargando...")
            time.sleep(2.5) 

            # --- PASO 4: RENOMBRAR ARCHIVO ---
            archivo_descargado = obtener_ultimo_archivo_pdf(CARPETA_DESCARGA)
            
            if archivo_descargado:
                ruta_final = os.path.join(CARPETA_DESCARGA, f"{documento_id}.pdf")
                
                if archivo_descargado != ruta_final:
                    if os.path.exists(ruta_final):
                        os.remove(ruta_final)
                    os.rename(archivo_descargado, ruta_final)
                    print(f"   -> ÉXITO: Guardado como {documento_id}.pdf")
                else:
                    print("   -> El archivo ya tiene el nombre correcto.")
            else:
                print("   -> ALERTA: No se encontró el archivo descargado.")

            # NOTA: Ya no hace falta volver a PORTADA manualmente, 
            # el 'driver.get' del inicio del bucle lo hace automáticamente.

        except Exception as e:
            print(f"   -> ERROR con {documento_id}: {e}")
            # Si falla, el 'continue' hará que el bucle reinicie y recargue la página
            continue

    print("\n--- PROCESO FINALIZADO ---")
    driver.quit()

if __name__ == "__main__":
    proceso_principal()