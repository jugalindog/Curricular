import os
import pandas as pd
import unicodedata

# Se importa la malla curricular para saber qué asignaturas debe ver el estudiante
try:
    from prueba10 import malla_curricular
except ImportError:
    print("ADVERTENCIA: No se pudo importar 'malla_curricular' desde 'prueba10.py'. La proyección no funcionará.")
    malla_curricular = {}

# --- Parámetros del plan ---
CREDITOS_TOTALES_180 = 180   # Hacia grado (sin nivelación)

# --- Rutas ---
RUTA_ARCHIVO = r"C:\Users\jp2g\Documents\PASANTIA\Curricular\Curricular\Estudiantes_simulados.xlsx"
RUTA_SALIDA  = r"C:\Users\jp2g\Documents\PASANTIA\Curricular\Curricular\Avance_y_Proyeccion_Estudiantes.xlsx"

def normalize_name(name: str) -> str:
    """Convierte un string a minúsculas, quita espacios extra y acentos."""
    if not isinstance(name, str): return ""
    return "".join(c for c in unicodedata.normalize('NFKD', name.lower().strip()) if not unicodedata.combining(c))

def cargar_datos():
    if not os.path.exists(RUTA_ARCHIVO):
        raise FileNotFoundError(f"No se encontró el archivo: {RUTA_ARCHIVO}")
    xls = pd.ExcelFile(RUTA_ARCHIVO)
    hoja = "Sheet1" if "Sheet1" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(RUTA_ARCHIVO, sheet_name=hoja)
    return df

def preparar_datos(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza columnas y limpia datos básicos del DataFrame."""
    # Normaliza nombres de columnas
    df.columns = [c.strip().lower() for c in df.columns]

    # Revisa columnas requeridas
    req = {"documento", "nombre", "creditos", "estado", "anulada", "semestre_inicio", "asignatura"}
    faltan = req - set(df.columns)
    if faltan:
        raise ValueError(f"Faltan columnas requeridas: {faltan}")

    # Limpieza básica
    df_limpio = df.drop_duplicates().copy()
    df_limpio["estado"]  = df_limpio["estado"].astype(str).str.strip().str.lower()
    df_limpio["anulada"] = df_limpio["anulada"].astype(str).str.strip().str.upper()  # "NO", "SI"
    df_limpio["creditos"] = pd.to_numeric(df_limpio["creditos"], errors="coerce").fillna(0)
    # Asegurarse de que la columna 'asignatura' es de tipo string para evitar errores
    df_limpio["asignatura"] = df_limpio["asignatura"].astype(str)
    return df_limpio

def calcular_avance(df: pd.DataFrame) -> pd.DataFrame:
    # Aprobadas y no anuladas
    aprob = (df["estado"] == "aprobada") & (df["anulada"] == "NO")
    dfa = df.loc[aprob, ["documento", "nombre", "semestre_inicio", "creditos"]].copy()

    # Suma créditos aprobados por estudiante
    cred = (
        dfa.groupby(["documento", "nombre", "semestre_inicio"], as_index=False, dropna=False)["creditos"]
           .sum()
           .rename(columns={"creditos": "creditos_aprobados"})
    )

    # Avance 180
    cred["avance_180_%"] = (cred["creditos_aprobados"] / CREDITOS_TOTALES_180 * 100).clip(upper=100).round(2)

    # También puedes querer un resumen 1 fila por estudiante:
    resumen = (
        cred.groupby(["documento", "nombre"], as_index=False, dropna=False)
            .agg(creditos_aprobados=("creditos_aprobados", "sum"),
                 semestre_inicio=("semestre_inicio", "min"))
    )
    resumen["avance_180_%"] = (resumen["creditos_aprobados"] / CREDITOS_TOTALES_180 * 100).clip(upper=100).round(2)

    return cred, resumen

def crear_proyeccion_por_asignatura(df: pd.DataFrame) -> pd.DataFrame:
    """
    Genera una proyección de las asignaturas faltantes para cada estudiante,
    basándose en la malla curricular.
    """
    if not malla_curricular:
        print("INFO: El diccionario 'malla_curricular' está vacío. No se puede generar la proyección.")
        return pd.DataFrame()

    # 1. Preparar datos de la malla curricular
    malla_lista = []
    for nombre, detalles in malla_curricular.items():
        # No se proyectan las asignaturas de nivelación como "faltantes" del plan de estudios.
        if 'tipo_asignatura' in detalles and 'Nivelación' in detalles['tipo_asignatura']:
            continue
        malla_lista.append({
            'asignatura_malla': nombre,
            'semestre_malla': detalles.get('semestre'),
            'nombre_normalizado_malla': normalize_name(nombre)
        })
    malla_df = pd.DataFrame(malla_lista)
    malla_set = set(malla_df['nombre_normalizado_malla'])

    # 2. Preparar datos de los estudiantes (solo asignaturas aprobadas)
    # La limpieza de datos ya se hizo en la función preparar_datos()
    aprobadas_df = df[(df['estado'] == 'aprobada') & (df['anulada'] == 'NO')].copy()
    aprobadas_df['nombre_normalizado_historial'] = aprobadas_df['asignatura'].apply(normalize_name)

    # 3. Iterar por cada estudiante para encontrar sus materias pendientes
    proyeccion_total = []
    estudiantes = df[['documento', 'nombre']].drop_duplicates().to_dict('records')

    for estudiante in estudiantes:
        doc = estudiante['documento']
        nom = estudiante['nombre']

        # Conjunto de asignaturas que el estudiante ya aprobó
        historial_estudiante = aprobadas_df[aprobadas_df['documento'] == doc]
        aprobadas_set = set(historial_estudiante['nombre_normalizado_historial'])

        # Asignaturas de la malla que NO están en las aprobadas (pendientes)
        pendientes_set = malla_set - aprobadas_set
        
        # Filtrar la malla para obtener los detalles de las asignaturas pendientes
        proyeccion_estudiante = malla_df[malla_df['nombre_normalizado_malla'].isin(pendientes_set)].copy()
        
        if not proyeccion_estudiante.empty:
            proyeccion_estudiante['documento'] = doc
            proyeccion_estudiante['nombre'] = nom
            proyeccion_total.append(proyeccion_estudiante)

    # 4. Consolidar y dar formato final
    if not proyeccion_total:
        print("INFO: No se generaron proyecciones. Puede que todos los estudiantes hayan completado la malla.")
        return pd.DataFrame()

    df_proyeccion_final = pd.concat(proyeccion_total, ignore_index=True)
    
    # Ordenar y seleccionar columnas para el resultado
    df_proyeccion_final = df_proyeccion_final[['documento', 'nombre', 'asignatura_malla', 'semestre_malla']]
    df_proyeccion_final = df_proyeccion_final.sort_values(by=['documento', 'semestre_malla']).reset_index(drop=True)

    return df_proyeccion_final

def main():
    try:
        df_raw = cargar_datos()
        df_limpio = preparar_datos(df_raw)
        _det, resumen = calcular_avance(df_limpio) # El detalle no se usará para exportar

        # Generar la proyección por asignatura
        proyeccion = crear_proyeccion_por_asignatura(df_limpio)

        # Exporta el resumen de avance y la proyección a hojas separadas
        with pd.ExcelWriter(RUTA_SALIDA, engine="openpyxl") as xw:
            resumen.to_excel(xw, sheet_name="Resumen_Avance", index=False)
            if not proyeccion.empty:
                proyeccion.to_excel(xw, sheet_name="Proyeccion_Asignaturas", index=False)

        print(f"✅ Exportado a: {RUTA_SALIDA}")
        
        print("\nPreview RESUMEN DE AVANCE:")
        print(resumen.head())

        if not proyeccion.empty:
            print("\nPreview PROYECCIÓN DE ASIGNATURAS:")
            print(proyeccion.head())
        else:
            print("\nNo se generó proyección de asignaturas.")

    except Exception as e:
        print(f"⚠️ Error: {e}")

if __name__ == "__main__":
    main()
