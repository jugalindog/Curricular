import pprint
import unicodedata
from prueba10 import malla_curricular
from prerequisitos_corregido import prerequisitos

# --- Función de Normalización ---
def normalize_name(name):
    """Convierte a minúsculas, quita espacios y acentos."""
    if not isinstance(name, str):
        return name
    nfkd_form = unicodedata.normalize('NFKD', name.lower().strip())
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

# Create a name-to-code mapping from the malla_curricular using normalized names
code_map = {normalize_name(course): (details.get('codigo'), course) for course, details in malla_curricular.items()}

# Update the prerequisites dictionary
updated_prerequisitos = {}
for course, prereq_groups in prerequisitos.items():
    new_prereq_groups = []
    for group in prereq_groups:
        new_group = []
        for prereq_item in group:
            # Extraer el nombre del prerrequisito, manejando diferentes estructuras
            if isinstance(prereq_item, tuple) and len(prereq_item) == 2:
                prereq_name = prereq_item[0]
            elif isinstance(prereq_item, str):
                prereq_name = prereq_item
            else:
                new_group.append(prereq_item) # Mantener estructura no reconocida
                continue

            # Normalizar y buscar el código
            normalized_prereq_name = normalize_name(prereq_name)
            found = code_map.get(normalized_prereq_name)

            if found:
                found_code, original_malla_name = found
                new_group.append((original_malla_name, found_code))
            else:
                new_group.append((prereq_name, None))
        new_prereq_groups.append(new_group)

    updated_prerequisitos[course] = new_prereq_groups

# Overwrite the prerequisitos_corregido.py file
with open('prerequisitos_corregido.py', 'w', encoding='utf-8') as f:
    f.write("prerequisitos = " + pprint.pformat(updated_prerequisitos))

print("El archivo 'prerequisitos_corregido.py' ha sido actualizado con normalización de acentos.")
