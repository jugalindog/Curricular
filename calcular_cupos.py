import pandas as pd
import unicodedata
from prerequisitos_corregido import prerequisitos
from prueba10 import malla_curricular # Importar para acceder a los códigos

def normalize_name(name):
    """Convierte a minúsculas, quita espacios y acentos."""
    if not isinstance(name, str):
        return name
    # Usar NFKD para separar caracteres base de diacríticos
    nfkd_form = unicodedata.normalize('NFKD', name.lower().strip())
    # Filtrar para quedarse solo con los caracteres que no son diacríticos (combining characters)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def verificar_prerequisitos(historial_estudiante, prerequisitos_asignatura, mode, debug=False):
    """
    Verifica si un estudiante cumple con los prerrequisitos de una asignatura.
    Puede comparar por 'codigo' o por 'nombre'.
    """
    if not prerequisitos_asignatura or not prerequisitos_asignatura[0]:
        return True

    if debug:
        print(f"[DEBUG] Modo de comparación: {mode}")
        print(f"[DEBUG] Prerequisitos para la asignatura: {prerequisitos_asignatura}")

    # Obtener las asignaturas aprobadas del historial una sola vez para optimizar
    historial_aprobado = historial_estudiante[historial_estudiante['estado'].str.strip() == 'Aprobada']
    
    # Crear un conjunto (set) de códigos o nombres aprobados para búsquedas rápidas
    if mode == 'codigo':
        aprobadas_set = set(historial_aprobado['codigo_asignatura'])
    else: # mode == 'nombre'
        aprobadas_set = set(historial_aprobado['asignatura_normalizada'])

    for grupo_prerequisitos in prerequisitos_asignatura:
        if debug:
            print(f"\n[DEBUG] Verificando GRUPO (condición Y): {grupo_prerequisitos}")

        primer_prereq_nombre = grupo_prerequisitos[0][0]
        if isinstance(primer_prereq_nombre, tuple):
            primer_prereq_nombre = primer_prereq_nombre[0]
        if primer_prereq_nombre == 'Ninguno':
            if debug:
                print("[DEBUG] Es 'Ninguno', retornando True.")
            return True

        cumple_grupo = True
        for prerequisito_tupla in grupo_prerequisitos:
            nombre_prerequisito, codigo_prerequisito = prerequisito_tupla
            
            aprobada = False
            if mode == 'codigo':
                if debug:
                    print(f"  [DEBUG] Verificando prerequisito por CÓDIGO: {nombre_prerequisito} ({codigo_prerequisito})")
                if codigo_prerequisito is None:
                    if debug:
                        print(f"    [DEBUG] Resultado: FALLA (prerrequisito sin código)")
                    cumple_grupo = False
                    break
                aprobada = str(codigo_prerequisito).strip() in aprobadas_set
            
            else: # mode == 'nombre'
                nombre_normalizado = normalize_name(nombre_prerequisito)
                if debug:
                    print(f"  [DEBUG] Verificando prerequisito por NOMBRE: '{nombre_prerequisito}' (normalizado: '{nombre_normalizado}')")
                aprobada = nombre_normalizado in aprobadas_set

            if debug:
                print(f"    [DEBUG] ¿Aprobado?: {aprobada}")

            if not aprobada:
                cumple_grupo = False
                break
        
        if debug:
            print(f"[DEBUG] Resultado del GRUPO: {'CUMPLE' if cumple_grupo else 'NO CUMPLE'}")

        if cumple_grupo:
            if debug:
                print("[DEBUG] Retornando True (un grupo 'Y' se cumplió)")
            return True

    if debug:
        print("[DEBUG] Retornando False (ningún grupo 'Y' se cumplió)")
    return False

def main():
    """
    Función principal que calcula y muestra el número de estudiantes aptos por asignatura.
    """
    # --- DEBUGGING SETUP ---
    # Para activar el modo debug para un caso específico, descomenta y asigna valores a estas variables.
    doc_debug = None
    asignatura_debug = None
    # doc_debug = "1000000041" # Ejemplo
    # asignatura_debug = "Ciclo  II: Ejecución de un proyecto productiv" # Ejemplo

    # --- SELECCIÓN DE MODO ---
    mode = ''
    while mode not in ['codigo', 'nombre']:
        mode = input("¿Cómo desea comparar las asignaturas? (escriba 'codigo' o 'nombre'): ").lower().strip()
        if mode not in ['codigo', 'nombre']:
            print("Opción no válida. Por favor, intente de nuevo.")
    # -------------------------

    try:
        df_estudiantes = pd.read_excel("Estudiantes_simulados.xlsx")
        # Asegurarse que los códigos de asignatura en el DF son strings y no tienen espacios extra
        df_estudiantes['codigo_asignatura'] = df_estudiantes['codigo_asignatura'].astype(str).str.strip()
        # Normalizar nombres en el dataframe de estudiantes para la comparación por nombre
        if mode == 'nombre':
            df_estudiantes['asignatura_normalizada'] = df_estudiantes['asignatura'].apply(normalize_name)

    except FileNotFoundError:
        print("Error: No se encontró el archivo 'Estudiantes_simulados.xlsx'.")
        return

    estudiantes_unicos = df_estudiantes['documento'].unique()

    print(f"\nCalculando cupos por asignatura usando '{mode}'...")
    print("-" * 45)

    for asignatura, prereqs in prerequisitos.items():
        info_asignatura = malla_curricular.get(asignatura)
        
        # Para la comparación por código, necesitamos el código.
        if mode == 'codigo' and (not info_asignatura or 'codigo' not in info_asignatura or info_asignatura['codigo'] is None):
            print(f"- {asignatura}: [No se puede calcular, sin código en la malla]")
            continue
        
        codigo_asignatura_actual = info_asignatura.get('codigo') if info_asignatura else None
        estudiantes_aptos = 0

        for doc_estudiante in estudiantes_unicos:
            historial_estudiante = df_estudiantes[df_estudiantes['documento'] == doc_estudiante]
            
            is_debug_case = str(doc_estudiante) == doc_debug and asignatura == asignatura_debug

            if is_debug_case:
                print(f"\n--- DEBUGGING ---")
                print(f"Estudiante: {doc_estudiante}, Asignatura: {asignatura}")
                if mode == 'codigo':
                    print(f"Verificando si ya vio la materia (código {codigo_asignatura_actual})...")
                else:
                    print(f"Verificando si ya vio la materia (nombre '{asignatura}')...")

            # Verificar si el estudiante ya vio o aprobó la materia
            ya_vio_la_materia = False
            if mode == 'codigo':
                ya_vio_la_materia = (historial_estudiante['codigo_asignatura'] == str(codigo_asignatura_actual)).any()
            else: # mode == 'nombre'
                nombre_asignatura_normalizado = normalize_name(asignatura)
                ya_vio_la_materia = (historial_estudiante['asignatura_normalizada'] == nombre_asignatura_normalizado).any()

            if is_debug_case:
                print(f"¿Ya vio la materia?: {ya_vio_la_materia}")
                print("Llamando a verificar_prerequisitos...")

            cumple_prereqs = verificar_prerequisitos(historial_estudiante, prereqs, mode=mode, debug=is_debug_case)

            if not ya_vio_la_materia and cumple_prereqs:
                estudiantes_aptos += 1
            
            if is_debug_case:
                print(f"Cumple Prerequisitos: {cumple_prereqs}")
                print(f"Resultado final de la verificación: Apto = {not ya_vio_la_materia and cumple_prereqs}")
                print(f"--- FIN DEBUGGING ---\n")
        
        print(f"- {asignatura}: {estudiantes_aptos} estudiantes aptos")

if __name__ == "__main__":
    main()
