import pandas as pd
import unicodedata
import re
from prerequisitos_corregido import prerequisitos
from prueba10 import malla_curricular # Importar para acceder a los códigos

# --- Rutas ---
RUTA_ESTUDIANTES = "/home/jugalindog/Pasantia/Curricular/Curricular/Estudiantes_simulados.xlsx"
#RUTA_ESTUDIANTES = r"C:\Users\jp2g\Documents\PASANTIA\Curricular\Curricular\Estudiantes_simulados.xlsx"
#RUTA_ESTUDIANTES = r"C:\Users\jp2g\Documents\PASANTIA\Curricular\Curricular\Prueba10_con_creditos.xlsx"
def normalize_name(name):
    """Convierte a minúsculas, quita espacios y acentos."""
    if not isinstance(name, str):
        return ""
    # Usar NFKD para separar caracteres base de diacríticos y convertir a minúsculas
    nfkd_form = unicodedata.normalize('NFKD', name.lower().strip())
    # Quitar diacríticos (marcas de combinación)
    s = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    # Reemplazar guiones y otros caracteres de puntuación comunes con un espacio
    s = re.sub(r'[-_()]', ' ', s)
    # Colapsar múltiples espacios en uno solo
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def get_siguiente_semestre(sem_str):
    """Calcula el siguiente periodo académico. Ej: '2022-2S' -> '2023-1S'."""
    if not isinstance(sem_str, str) or '-' not in sem_str:
        return "Proyección Inválida"
    try:
        year_str, period_str = sem_str.split('-')
        year = int(year_str)
        period = int(period_str[0])
        if period == 1:
            return f"{year}-2S"
        else: # period == 2
            return f"{year + 1}-1S"
    except (ValueError, IndexError):
        return "Proyección Inválida"

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
        df_estudiantes = pd.read_excel(RUTA_ESTUDIANTES)
        # --- Data Preparation ---
        # Ensure required columns are of the correct type
        df_estudiantes['codigo_asignatura'] = df_estudiantes['codigo_asignatura'].astype(str).str.strip()
        df_estudiantes['estado'] = df_estudiantes['estado'].astype(str).str.strip()
        # Ensure 'semestre_asignatura' exists for projection
        if 'semestre_asignatura' not in df_estudiantes.columns:
            print("Error: La columna 'semestre_asignatura' es necesaria para la proyección y no se encontró.")
            return
        # Ensure 'semestre_inicio' exists for new student projection
        if 'semestre_inicio' not in df_estudiantes.columns:
            print("Error: La columna 'semestre_inicio' es necesaria para la proyección y no se encontró.")
            return
        # Siempre normalizar nombres para una verificación robusta, independientemente del modo.
        df_estudiantes['asignatura_normalizada'] = df_estudiantes['asignatura'].apply(normalize_name)
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo en la ruta especificada: '{RUTA_ESTUDIANTES}'.")
        return
    except KeyError as e:
        print(f"Error: Falta la columna {e} en el archivo Excel.")
        return

    # --- Initialization ---
    # For the summary of how many students are eligible per course
    resumen_cupos_dict = {asignatura: 0 for asignatura in malla_curricular.keys()}
    # For the detailed projection for each student
    proyecciones_individuales = []
    
    # Create a map from normalized prerequisite names to their rules for robust lookup
    prereq_norm_map = {normalize_name(key): value for key, value in prerequisitos.items()}

    # Get a list of unique students to iterate over
    estudiantes_unicos = df_estudiantes[['documento', 'nombre']].drop_duplicates().to_dict('records')

    print(f"\nCalculando cupos y generando proyección usando '{mode}'...")
    print("-" * 55)

    # --- Main Processing Loop (Student by Student) ---
    for estudiante in estudiantes_unicos:
        doc_estudiante = estudiante['documento']
        nombre_estudiante = estudiante['nombre']
        
        historial_estudiante = df_estudiantes[df_estudiantes['documento'] == doc_estudiante]
        
        # --- Determinar el semestre de inicio para la proyección ---
        # Busca el último semestre que el estudiante cursó activamente.
        ultimo_semestre_cursado = historial_estudiante['semestre_asignatura'].max()
        
        semestre_base_proyeccion = None
        if pd.isna(ultimo_semestre_cursado):
            # Caso 1: Estudiante nuevo sin historial académico.
            # La proyección comienza en su 'semestre_inicio'.
            semestre_inicio_estudiante = historial_estudiante['semestre_inicio'].first()
            if pd.notna(semestre_inicio_estudiante):
                semestre_base_proyeccion = semestre_inicio_estudiante
            else:
                # Si no tiene historial ni semestre de inicio, no se puede proyectar.
                continue
        else:
            # Caso 2: Estudiante con historial existente.
            # La proyección comienza en el semestre SIGUIENTE al último que cursó.
            semestre_base_proyeccion = get_siguiente_semestre(ultimo_semestre_cursado)

        # Find all courses the student is eligible for in the next term
        asignaturas_elegibles = []
        # Use malla_curricular as the source of truth for courses to avoid name inconsistencies
        for asignatura, info_asignatura in malla_curricular.items():
            
            # Find prerequisites for the current course using the normalized map
            # If not found, assume no prerequisites ([[]])
            prereqs = prereq_norm_map.get(normalize_name(asignatura), [[]])

            codigo_asignatura_actual = info_asignatura.get('codigo') if info_asignatura else None

            # 1. Comprobación robusta para ver si el estudiante ya ha cursado la materia.
            #    Esta verificación ahora es independiente del 'modo' seleccionado para los prerrequisitos.
            ya_vio_la_materia = False
            # Primero, intentar la coincidencia por el código oficial de la asignatura (más fiable).
            if codigo_asignatura_actual:
                if (historial_estudiante['codigo_asignatura'] == str(codigo_asignatura_actual)).any():
                    ya_vio_la_materia = True
            
            # Si no se encuentra por código, intentar una coincidencia de respaldo por nombre normalizado.
            # Esto maneja casos donde el código puede estar ausente o incorrecto en el historial del estudiante.
            if not ya_vio_la_materia:
                nombre_malla_normalizado = normalize_name(asignatura)
                historial_nombres_norm = historial_estudiante['asignatura_normalizada']
                
                # Verificación 1: Coincidencia exacta con el nombre normalizado.
                if (historial_nombres_norm == nombre_malla_normalizado).any():
                    ya_vio_la_materia = True
                else:
                    # Verificación 2: Coincidencia flexible (ignorando ' de ').
                    # Esto es para casos como "laboratorio de bioquimica" vs "laboratorio bioquimica".
                    nombre_malla_sin_de = nombre_malla_normalizado.replace(' de ', ' ')
                    # Se aplica la misma transformación a la serie de nombres del historial para comparar.
                    if (historial_nombres_norm.str.replace(' de ', ' ') == nombre_malla_sin_de).any():
                        ya_vio_la_materia = True
            
            # 2. Check if student meets the prerequisites
            is_debug_case = str(doc_estudiante) == doc_debug and asignatura == asignatura_debug
            cumple_prereqs = verificar_prerequisitos(historial_estudiante, prereqs, mode=mode, debug=is_debug_case)

            # 3. If eligible, add to lists for summary and projection
            if not ya_vio_la_materia and cumple_prereqs:
                resumen_cupos_dict[asignatura] += 1
                semestre_malla = info_asignatura.get('semestre') if info_asignatura and info_asignatura.get('semestre') is not None else 99
                asignaturas_elegibles.append({'asignatura': asignatura, 'semestre_malla': semestre_malla})

        # --- Project courses for the current student ---
        if asignaturas_elegibles:
            # Sort eligible courses by their ideal semester in the curriculum
            asignaturas_elegibles.sort(key=lambda x: x['semestre_malla'])
            
            semestre_proyeccion_actual = semestre_base_proyeccion
            semestre_malla_actual = None

            for materia_elegible in asignaturas_elegibles:
                # If we are moving to a new semester level in the malla, we advance the projected semester
                if semestre_malla_actual is not None and materia_elegible['semestre_malla'] > semestre_malla_actual:
                    semestre_proyeccion_actual = get_siguiente_semestre(semestre_proyeccion_actual)
                
                proyecciones_individuales.append({
                    'documento': doc_estudiante,
                    'nombre': nombre_estudiante,
                    'asignatura': materia_elegible['asignatura'],
                    'semestre_proyectado': semestre_proyeccion_actual
                })
                
                semestre_malla_actual = materia_elegible['semestre_malla']

    # --- Finalization and Export ---
    # Convert results to DataFrames
    df_resumen_cupos = pd.DataFrame(list(resumen_cupos_dict.items()), columns=['Asignatura', 'Estudiantes Aptos'])
    df_proyeccion = pd.DataFrame(proyecciones_individuales)

    # Sort the summary for better readability
    df_resumen_cupos = df_resumen_cupos.sort_values(by='Estudiantes Aptos', ascending=False).reset_index(drop=True)

    # Export to a single Excel file with two sheets
    output_filename = "Calculo_y_Proyeccion_Cupos.xlsx"
    try:
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            df_resumen_cupos.to_excel(writer, sheet_name='Resumen_Cupos', index=False)
            if not df_proyeccion.empty:
                df_proyeccion.to_excel(writer, sheet_name='Proyeccion_Individual', index=False)
        print(f"\n✅ Resultados exportados exitosamente a '{output_filename}'")
    except Exception as e:
        print(f"\n⚠️ Error al guardar el archivo Excel: {e}")

    # Print summary to console
    print("\n--- Resumen de Cupos ---")
    print(df_resumen_cupos.to_string())

    if not df_proyeccion.empty:
        print("\n--- Vista Previa de Proyección Individual ---")
        print(df_proyeccion.head())
        
if __name__ == "__main__":
    main()
