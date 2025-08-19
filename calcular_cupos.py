import pandas as pd
from prerequisitos_corregido import prerequisitos
from prueba10 import malla_curricular # Importar para acceder a los códigos

def verificar_prerequisitos(historial_estudiante, prerequisitos_asignatura, debug=False):
    """
    Verifica si un estudiante cumple con los prerrequisitos de una asignatura usando CÓDIGOS.
    """
    if not prerequisitos_asignatura or not prerequisitos_asignatura[0]:
        return True

    if debug:
        print(f"[DEBUG] Prerequisitos para la asignatura: {prerequisitos_asignatura}")

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
            if debug:
                print(f"  [DEBUG] Verificando prerequisito: {nombre_prerequisito} ({codigo_prerequisito})")

            if codigo_prerequisito is None:
                if debug:
                    print(f"    [DEBUG] Resultado: FALLA (prerrequisito sin código)")
                cumple_grupo = False
                break

            aprobada = ((historial_estudiante['codigo_asignatura'].str.strip() == str(codigo_prerequisito).strip()) & (historial_estudiante['estado'].str.strip() == 'Aprobada')).any()
            
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
    doc_debug = "1000000041"
    asignatura_debug = "Ciclo  II: Ejecución de un proyecto productiv"
    # -----------------------

    try:
        df_estudiantes = pd.read_excel("Estudiantes_simulados.xlsx")
        # Asegurarse que los códigos de asignatura en el DF son strings para comparación
        df_estudiantes['codigo_asignatura'] = df_estudiantes['codigo_asignatura'].astype(str)

    except FileNotFoundError:
        print("Error: No se encontró el archivo 'Estudiantes_simulados.xlsx'.")
        return

    estudiantes_unicos = df_estudiantes['documento'].unique()

    print("Calculando cupos por asignatura usando códigos...")
    print("-" * 45)

    for asignatura, prereqs in prerequisitos.items():
        info_asignatura = malla_curricular.get(asignatura)
        if not info_asignatura or 'codigo' not in info_asignatura or info_asignatura['codigo'] is None:
            print(f"- {asignatura}: [No se puede calcular, sin código en la malla]")
            continue
        
        codigo_asignatura_actual = info_asignatura['codigo']
        estudiantes_aptos = 0

        for doc_estudiante in estudiantes_unicos:
            historial_estudiante = df_estudiantes[df_estudiantes['documento'] == doc_estudiante]
            
            is_debug_case = str(doc_estudiante) == doc_debug and asignatura == asignatura_debug

            if is_debug_case:
                print(f"\n--- DEBUGGING ---")
                print(f"Estudiante: {doc_estudiante}, Asignatura: {asignatura}")
                print(f"Verificando si ya vio la materia (código {codigo_asignatura_actual})...")

            ya_vio_la_materia = (historial_estudiante['codigo_asignatura'] == str(codigo_asignatura_actual)).any()

            if is_debug_case:
                print(f"¿Ya vio la materia?: {ya_vio_la_materia}")
                print("Llamando a verificar_prerequisitos...")

            cumple_prereqs = verificar_prerequisitos(historial_estudiante, prereqs, debug=is_debug_case)

            if not ya_vio_la_materia and cumple_prereqs:
                estudiantes_aptos += 1
            
            if is_debug_case:
                print(f"Resultado final de la verificación: Apto = {cumple_prereqs}")
                print(f"--- FIN DEBUGGING ---\n")
        
        print(f"- {asignatura}: {estudiantes_aptos} estudiantes aptos")

if __name__ == "__main__":
    main()