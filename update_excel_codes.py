import pandas as pd
import unicodedata
from prueba10 import malla_curricular

def normalize_name(name):
    """Convierte a minúsculas, quita espacios y acentos."""
    if not isinstance(name, str):
        return name
    nfkd_form = unicodedata.normalize('NFKD', name.lower().strip())
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

def add_missing_codes():
    excel_file = "Estudiantes_simulados.xlsx"
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo '{excel_file}'.")
        return

    print("Archivo leído. Buscando y añadiendo códigos faltantes...")

    # Crear un mapa de nombres normalizados a códigos desde la malla curricular
    code_map = {normalize_name(course): details.get('codigo') for course, details in malla_curricular.items()}

    # Diccionario para rastrear los códigos inventados y asignarlos consistentemente
    invented_codes = {}
    invented_counter = 1

    # Asegurarse de que la columna existe
    if 'codigo_asignatura' not in df.columns:
        df['codigo_asignatura'] = None

    # Llenar los códigos faltantes
    for index, row in df.iterrows():
        # Se considera faltante si es nulo o un string vacío
        if pd.isna(row['codigo_asignatura']) or str(row['codigo_asignatura']).strip() == '':
            asignatura_name = row.get('asignatura')
            if isinstance(asignatura_name, str) and asignatura_name.strip() != '':
                normalized = normalize_name(asignatura_name)
                
                # 1. Buscar en la malla curricular
                code = code_map.get(normalized)
                
                if code:
                    df.at[index, 'codigo_asignatura'] = code
                else:
                    # 2. Si no está en la malla, inventar un código
                    if normalized not in invented_codes:
                        invented_codes[normalized] = f"INVENTADO_{invented_counter}"
                        invented_counter += 1
                    
                    df.at[index, 'codigo_asignatura'] = invented_codes[normalized]

    # Guardar el DataFrame actualizado
    try:
        df.to_excel(excel_file, index=False)
        print(f"¡Éxito! El archivo '{excel_file}' ha sido actualizado.")
        if invented_codes:
            print("Se inventaron códigos para las siguientes asignaturas no encontradas en la malla:")
            for name, code in invented_codes.items():
                print(f"- {name}: {code}")

    except Exception as e:
        print(f"Error al guardar el archivo: {e}")

if __name__ == "__main__":
    add_missing_codes()
