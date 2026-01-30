import pandas as pd
import unicodedata
import re
import os

# --- IMPORTACIONES LOCALES ---
from prerequisitos_corregido import prerequisitos
from malla_FCA import malla_curricular, optativas_produccion

# --- RUTAS ---
#RUTA_ESTUDIANTES = r"C:\Users\JuanPabloGalindoGóme\Documents\Curricular\Historias_academicas.xlsx"
#RUTA_ESTUDIANTES ="/home/jugalindog/Pasantia/Curricular/Curricular/Historias_academicas.xlsx"
RUTA_ESTUDIANTES ="/home/jugalindog/Pasantia/Curricular/Curricular/Historias_academicas3.xlsx"

# --- UNIFICAR MALLA + OPTATIVAS ---
malla_completa = {}

# Malla obligatoria
for k, v in malla_curricular.items():
    v_copy = v.copy()
    v_copy['tipo'] = 'obligatoria'
    malla_completa[k] = v_copy

# Optativas de producción
for k, v in optativas_produccion.items():
    v_copy = v.copy()
    v_copy['tipo'] = 'optativa_produccion'
    malla_completa[k] = v_copy


def normalize_name(name):
    """Convierte a minúsculas, quita espacios y acentos."""
    if not isinstance(name, str): return ""
    nfkd_form = unicodedata.normalize('NFKD', name.lower().strip())
    s = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    s = re.sub(r'[-_()]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s
# --- SETS PARA RECONOCER OPTATIVAS DE PRODUCCIÓN ---


OPT_CODIGOS = set(str(v.get("codigo", "")).strip() for v in optativas_produccion.values())
OPT_NOMBRES = set(normalize_name(k) for k in optativas_produccion.keys())


def get_siguiente_semestre(sem_str):
    """Calcula el siguiente periodo académico."""
    if not isinstance(sem_str, str) or '-' not in sem_str: return "Proyección Inválida"
    try:
        year_str, period_str = sem_str.split('-')
        year, period = int(year_str), int(period_str[0])
        return f"{year}-2S" if period == 1 else f"{year + 1}-1S"
    except (ValueError, IndexError): return "Proyección Inválida"

# --- LÓGICA DE EXENCIONES (PLAN) ---
def obtener_exenciones_por_plan(plan_texto):
    """Devuelve materias a aprobar automáticamente según el plan."""
    if not isinstance(plan_texto, str): return []
    
    CODIGO_MATE_BASICA = '1000001-B'
    CODIGO_LECTO = '1000002-B'
    exenciones = []
    plan_lower = plan_texto.lower()
    
    # Palabras clave detectadas en tus reportes
    if "no deben nivelar" in plan_lower:
        exenciones.extend([CODIGO_MATE_BASICA, CODIGO_LECTO])
    elif "deben nivelar matemáticas" in plan_lower:
        exenciones.append(CODIGO_LECTO) # Ya aprobó lecto
    elif ("deben nivelar lecto" in plan_lower) and "no deben" not in plan_lower:
        exenciones.append(CODIGO_MATE_BASICA) # Ya aprobó mate
        
    return exenciones

# --- VERIFICACIÓN ROBUSTA DE PRERREQUISITOS ---
def verificar_prerequisitos(historial_estudiante, prerequisitos_asignatura, mode, debug=False):
    """
    Verifica prerrequisitos manejando errores de formato, códigos faltantes y lógica AND/OR.
    """
    # 1. Estandarizar entrada (Lista o Tupla con Operador)
    lista_grupos = []
    operador = 'AND'

    # Caso: Tupla ( [[...]], 'OR' )
    if isinstance(prerequisitos_asignatura, tuple) and len(prerequisitos_asignatura) == 2:
        lista_grupos = prerequisitos_asignatura[0]
        operador = prerequisitos_asignatura[1]
    # Caso: Lista antigua [[...]]
    elif isinstance(prerequisitos_asignatura, list):
        lista_grupos = prerequisitos_asignatura
        operador = 'AND'
    
    # Si la lista está vacía o el primer elemento es vacío
    if not lista_grupos or (len(lista_grupos) > 0 and not lista_grupos[0]):
        return True

    # 2. Pre-cargar materias aprobadas para búsqueda rápida
    historial_aprobado = historial_estudiante[historial_estudiante['estado'].str.strip() == 'Aprobada']
    if mode == 'codigo':
        aprobadas_set = set(historial_aprobado['codigo_asignatura'])
    else:
        aprobadas_set = set(historial_aprobado['asignatura_normalizada'])

    resultados_grupos = []

    for grupo in lista_grupos:
        # Manejo de 'Ninguno'
        if isinstance(grupo, (list, tuple)) and len(grupo) > 0:
            primer_elem = grupo[0]
            if isinstance(primer_elem, (list, tuple)) and len(primer_elem) > 0 and primer_elem[0] == 'Ninguno':
                resultados_grupos.append(True)
                continue

        cumple_grupo = True
        
        # Iterar sobre cada materia dentro del grupo (Lógica interna del grupo suele ser AND)
        for prereq in grupo:
            nombre_req = ""
            codigo_req = None
            
            # --- BLINDAJE: Extracción segura de datos ---
            try:
                if isinstance(prereq, (list, tuple)):
                    if len(prereq) >= 2:
                        nombre_req, codigo_req = prereq[0], prereq[1]
                    elif len(prereq) == 1:
                        nombre_req = prereq[0]
                else:
                    # Formato desconocido, saltar
                    continue
            except Exception:
                continue
            
            # Limpieza de tuplas anidadas extrañas (ej: (('Nombre',), 'Cod'))
            if isinstance(nombre_req, tuple): nombre_req = nombre_req[0]

            # Verificación
            aprobada = False
            if mode == 'codigo':
                if codigo_req and str(codigo_req).strip() in aprobadas_set:
                    aprobada = True
            else: # mode == 'nombre'
                if normalize_name(nombre_req) in aprobadas_set:
                    aprobada = True
            
            if not aprobada:
                cumple_grupo = False
                break
        
        resultados_grupos.append(cumple_grupo)

    # 3. Resultado Final
    if operador == 'OR':
        return any(resultados_grupos)
    else: # AND
        return all(resultados_grupos)

def main():
    mode = ''
    while mode not in ['codigo', 'nombre']:
        mode = input("¿Comparar por 'codigo' o 'nombre'? ").lower().strip()

    print(f"\n--- Iniciando Procesamiento ({mode}) ---")

    if not os.path.exists(RUTA_ESTUDIANTES):
        print(f"❌ Error: No se encuentra {RUTA_ESTUDIANTES}")
        return

    try:
        df_estudiantes = pd.read_excel(RUTA_ESTUDIANTES)
        # Limpieza inicial
        df_estudiantes.columns = [c.lower() for c in df_estudiantes.columns] # Todo a minúsculas
        df_estudiantes['codigo_asignatura'] = df_estudiantes['codigo_asignatura'].astype(str).str.strip()
        df_estudiantes['estado'] = df_estudiantes['estado'].astype(str).str.strip()
        df_estudiantes['asignatura_normalizada'] = df_estudiantes['asignatura'].apply(normalize_name)
        
        if 'plan' not in df_estudiantes.columns:
            print("⚠️ Advertencia: Columna 'plan' no encontrada. Se asume vacío.")
            df_estudiantes['plan'] = ''
            
    except Exception as e:
        print(f"❌ Error leyendo Excel: {e}")
        return

    # Solo contamos obligatorias + el bucket genérico de optativas
    resumen_cupos = {asig: 0 for asig, info in malla_completa.items() if info.get("tipo") != "optativa_produccion"}
    resumen_cupos["Optativa de producción"] = 0
    proyecciones = []
    
    # Mapa de prerrequisitos normalizado
    prereq_map = {normalize_name(k): v for k, v in prerequisitos.items()}

    estudiantes_unicos = df_estudiantes[['documento', 'nombre']].drop_duplicates().to_dict('records')

    for est in estudiantes_unicos:
        doc = est['documento']
        nom = est['nombre']

        elegibles_obligatorias = []
        elegibles_optativas = []
        
        # Copia aislada del historial
        historial = df_estudiantes[df_estudiantes['documento'] == doc].copy()

        # 1. INYECTAR EXENCIONES DEL PLAN
        try:
            planes = historial['plan'].dropna()
            plan_txt = str(planes.iloc[0]) if not planes.empty else ''
        except: plan_txt = ''

        for cod_ex in obtener_exenciones_por_plan(plan_txt):
            # Verificar si ya existe aprobada
            ya_esta = historial[
                (historial['codigo_asignatura'] == cod_ex) & 
                (historial['estado'] == 'Aprobada')
            ]
            if ya_esta.empty:
                # Buscar nombre real en la malla
                nom_real = "Nivelación"
                for k, v in malla_curricular.items():
                    if str(v.get('codigo')) == cod_ex:
                        nom_real = k; break
                
                # Crear fila
                fila = {
                    'documento': doc, 'nombre': nom,
                    'codigo_asignatura': cod_ex, 'asignatura': nom_real,
                    'asignatura_normalizada': normalize_name(nom_real),
                    'estado': 'Aprobada', 
                    'semestre_asignatura': 'Eximido', 
                    'semestre_inicio': 'Eximido',
                    'plan': plan_txt
                }
                historial = pd.concat([historial, pd.DataFrame([fila])], ignore_index=True)

        # 2. DEFINIR SEMESTRE BASE PROYECCIÓN
        ult_sem = historial[historial['semestre_asignatura'] != 'Eximido']['semestre_asignatura'].max()
        
        sem_base = None
        if pd.isna(ult_sem):
            # Usar iloc[0] para evitar error .first()
            sem_inicio_series = historial[historial['semestre_inicio'] != 'Eximido']['semestre_inicio']
            if not sem_inicio_series.empty:
                sem_base = sem_inicio_series.iloc[0]
            else:
                continue # No hay datos suficientes
        else:
            sem_base = get_siguiente_semestre(ult_sem)

        # 3. EVALUAR MATERIAS
        elegibles = []
        
        # Formato: 'Código que pide la Malla': ['Código Alternativo 1', 'Código Alternativo 2']
        equivalencias_codigos = {
            '1000012-B': ['1000013-B'], # Bioestadística se considera vista si tiene Probabilidad
        }
        
        alias_materias = {
        'bioestadistica fundamental': ['bioestadistica', 'probabilidad y estadistica'],
        }
        

        for asignatura, info in malla_completa.items(): # <--- BUSCA ESTA LÍNEA
        

            # --- INICIO DEL BLOQUE NUEVO ---
            cod_malla = str(info.get('codigo', ''))
            nombre_malla_norm = normalize_name(asignatura)
            
            ya_vio_la_materia = False
            
            # A. Chequeo por Código (Exacto)
            if cod_malla and (historial['codigo_asignatura'] == cod_malla).any():
                ya_vio_la_materia = True
                
            # B. Chequeo por Nombre (Normalizado)
            elif (historial['asignatura_normalizada'] == nombre_malla_norm).any():
                ya_vio_la_materia = True

            # C. Chequeo por EQUIVALENCIAS (LO QUE NECESITAS)
            elif cod_malla in equivalencias_codigos:
                codigos_alternativos = equivalencias_codigos[cod_malla]
                # Verifica si tiene alguno de los códigos alternativos aprobados
                if historial[historial['codigo_asignatura'].isin(codigos_alternativos) & (historial['estado'] == 'Aprobada')].shape[0] > 0:
                    ya_vio_la_materia = True

            # D. Chequeo por Alias (Opcional)
            elif nombre_malla_norm in alias_materias:
                nombres_alt = alias_materias[nombre_malla_norm]
                if historial[historial['asignatura_normalizada'].isin(nombres_alt)].shape[0] > 0:
                    ya_vio_la_materia = True
            
            # Si ya la vio, pasamos a la siguiente materia sin hacer nada más
            if ya_vio_la_materia: 
                continue

                # --- PASO 3: BLOQUE PARA OPTATIVAS ---
            if info.get('tipo') == 'optativa_produccion':
                codigo_opt = str(info.get('codigo', '')).strip()

                # Si ya cursó esa optativa, no se vuelve a contar
                if codigo_opt and (historial['codigo_asignatura'] == codigo_opt).any():
                    continue    

            # B. Prerrequisitos
            reqs = prereq_map.get(normalize_name(asignatura), [])
            cumple = verificar_prerequisitos(historial, reqs, mode)

            if cumple:
                if info.get("tipo") == "optativa_produccion":
                    elegibles_optativas.append({
                        "asignatura": asignatura,
                        "semestre_malla": info.get("semestre", 99)
                    })
                else:
                    resumen_cupos[asignatura] += 1
                    elegibles_obligatorias.append({
                        "asignatura": asignatura,
                        "semestre_malla": info.get("semestre", 99)
                    })
                            # --- REGLA: cada estudiante debe tener mínimo 2 optativas de producción ---
        hist_aprob = historial[historial["estado"].str.strip() == "Aprobada"]
        
        # Cuántas optativas de producción ya tiene (por código o por nombre)
        opt_aprob_por_codigo = hist_aprob["codigo_asignatura"].astype(str).str.strip().isin(OPT_CODIGOS).sum()
        opt_aprob_por_nombre = hist_aprob["asignatura_normalizada"].isin(OPT_NOMBRES).sum()
        
        # usa el máximo para no quedarte corto si alguna vino sin código o sin nombre limpio
        opt_aprobadas = int(max(opt_aprob_por_codigo, opt_aprob_por_nombre))
        
        # Cupos que necesita para llegar a mínimo 2
        cupos_necesarios = max(0, 2 - opt_aprobadas)
        
        # Si ya cumple, no sumamos nada
        if cupos_necesarios > 0:
            # Si hay al menos una optativa elegible, sumamos cupos (limitado por opciones disponibles)
            cupos_a_sumar = min(cupos_necesarios, len(elegibles_optativas))
            resumen_cupos["Optativa de producción"] += cupos_a_sumar
    # 4. PROYECCIÓN INDIVIDUAL
        if elegibles:
            elegibles.sort(key=lambda x: x['semestre_malla'])
            sem_actual = sem_base
            sem_malla_ant = None
            
            for mat in elegibles:
                if sem_malla_ant is not None and mat['semestre_malla'] > sem_malla_ant:
                    sem_actual = get_siguiente_semestre(sem_actual)
                
                proyecciones.append({
                    'documento': doc, 'nombre': nom,
                    'asignatura': mat['asignatura'],
                    'semestre_proyectado': sem_actual
                })
                sem_malla_ant = mat['semestre_malla']

    # --- EXPORTAR ---
    try:
        df_res = pd.DataFrame(list(resumen_cupos.items()), columns=['Asignatura', 'Estudiantes Aptos'])
        df_res = df_res.sort_values('Estudiantes Aptos', ascending=False)
        df_proy = pd.DataFrame(proyecciones)

        with pd.ExcelWriter("Calculo_y_Proyeccion_Cupos_3.xlsx", engine='openpyxl') as writer:
            df_res.to_excel(writer, sheet_name='Resumen', index=False)
            if not df_proy.empty:
                df_proy.to_excel(writer, sheet_name='Detalle', index=False)
        print("\n✅ Proceso finalizado con éxito.")
    except Exception as e:
        print(f"\n❌ Error guardando Excel: {e}")

if __name__ == "__main__":
    main()