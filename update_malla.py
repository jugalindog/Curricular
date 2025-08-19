import pandas as pd
import ast
import pprint

# Read the content of prueba10.py
try:
    with open('prueba10.py', 'r', encoding='utf-8') as f:
        content = f.read()
except FileNotFoundError:
    print("Error: No se encontró el archivo 'prueba10.py'.")
    exit()

# Find the start of the malla_curricular dictionary
start_index = content.find('malla_curricular = {')
if start_index == -1:
    print("Error: No se pudo encontrar el inicio del diccionario 'malla_curricular'.")
    exit()

# Find the opening brace
open_brace_index = content.find('{', start_index)
if open_brace_index == -1:
    print("Error: No se pudo encontrar la llave de apertura del diccionario 'malla_curricular'.")
    exit()

# Find the matching closing brace
brace_count = 1
i = open_brace_index + 1
while i < len(content) and brace_count > 0:
    if content[i] == '{':
        brace_count += 1
    elif content[i] == '}':
        brace_count -= 1
    i += 1

if brace_count != 0:
    print("Error: No se pudo encontrar la llave de cierre del diccionario 'malla_curricular'.")
    exit()

end_index = i
malla_str_with_variable = content[start_index:end_index]
malla_str = content[open_brace_index:end_index]


# Use ast.literal_eval to convert the string to a dictionary
try:
    malla_curricular = ast.literal_eval(malla_str)
except Exception as e:
    print(f"Error al procesar el diccionario malla_curricular: {e}")
    exit()

# Read the Excel file and create the code map
try:
    df = pd.read_excel("Prueba10_con_creditos.xlsx")
    df_unique = df.drop_duplicates(subset=['asignatura'])
    # Convert course names to string to be safe
    df_unique['asignatura'] = df_unique['asignatura'].astype(str)
    code_map = pd.Series(df_unique.codigo_asignatura.values, index=df_unique.asignatura).to_dict()
except FileNotFoundError:
    print("Error: Archivo 'Prueba10_con_creditos.xlsx' no encontrado.")
    exit()
except Exception as e:
    print(f"Error al procesar el archivo de Excel: {e}")
    exit()

# Update the malla_curricular dictionary
for course, details in malla_curricular.items():
    # A simple normalization for matching course names
    normalized_course = course.strip()
    for map_course, code in code_map.items():
        normalized_map_course = map_course.strip()
        if normalized_course.lower() == normalized_map_course.lower():
            details['codigo'] = str(code) # Ensure code is a string
            break

# Convert the updated dictionary back to a pretty-printed string
updated_malla_str = "malla_curricular = " + pprint.pformat(malla_curricular)

# Replace the old dictionary string with the new one in the original content
updated_content = content.replace(malla_str_with_variable, updated_malla_str)

# Write the updated content back to the file
try:
    with open('prueba10.py', 'w', encoding='utf-8') as f:
        f.write(updated_content)
    print("El archivo 'prueba10.py' ha sido actualizado correctamente.")
except Exception as e:
    print(f"Error al escribir en el archivo 'prueba10.py': {e}")