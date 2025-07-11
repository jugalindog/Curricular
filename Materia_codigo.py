import re
import fitz
import pandas as pd
import os

CARPETA_PDFS = "/home/jugalindog/Pasantia/Curricular/Historial_Academica"
datos = []

for archivo in os.listdir(CARPETA_PDFS):
    if not archivo.endswith(".pdf"):
        continue

    ruta_pdf = os.path.join(CARPETA_PDFS, archivo)
    try:
        doc = fitz.open(ruta_pdf)
        texto = "\n".join([page.get_text() for page in doc])
        doc.close()
    except Exception as e:
        print(f"Error con {archivo}: {e}")
        continue

    nombre_match = re.search(r'Nombre:\s*(.+)', texto)
    documento_match = re.search(r'Documento:\s*(\d+)', texto)
    if not nombre_match or not documento_match:
        continue

    nombre = nombre_match.group(1).strip()
    documento = documento_match.group(1).strip()

    bloques = re.split(r'(?:PRIMER|SEGUNDO)\s+PERIODO\s+(\d{4}-[12]S)', texto)

    for i in range(1, len(bloques), 2):
        semestre = bloques[i]
        contenido = bloques[i + 1]
        lineas = contenido.splitlines()

        j = 0
        while j < len(lineas):
            linea = lineas[j].strip()

            # Detectar nombre y código al final de la línea
            match_asig = re.match(r'^(.+?)\s*\((\d{6,7}(?:-B)?)\)$', linea)
            if match_asig:
                nombre_asig = match_asig.group(1).strip()
                codigo = match_asig.group(2).strip()
                tipo_asig = ''
                nota = ''
                estado = 'Reprobada'

                # Leer líneas posteriores
                j += 1
                while j < len(lineas):
                    l = lineas[j].strip()

                    # Si encontramos otra asignatura, retrocedemos
                    if re.match(r'^(.+?)\s*\((\d{6,7}(?:-B)?)\)$', l):
                        j -= 1
                        break

                    # Buscar nota y estado
                    if re.search(r'(Aprobada|Reprobada|SI\*)', l):
                        nota_match = re.search(r'([\d,]+)', l)
                        if nota_match:
                            nota = nota_match.group(1).replace(',', '.')
                            try:
                                nota_val = float(nota)
                            except:
                                nota_val = 0.0
                        estado = 'Aprobada' if 'Aprobada' in l else 'Reprobada'
                    
                    # Buscar tipo de asignatura
                    elif any(t in l for t in ['Obligatoria', 'Optativa', 'Libre Elección', 'Nivelación']):
                        tipo_asig = l

                    j += 1

                datos.append({
                    'nombre': nombre,
                    'documento': documento,
                    'codigo_asignatura': codigo,
                    'asignatura': nombre_asig,
                    'tipo_asignatura': tipo_asig,
                    'nota': float(nota) if nota.replace('.', '', 1).isdigit() else 0.0,
                    'estado': estado,
                    'semestre_inicio': '2018-2S',
                    'semestre_asignatura': semestre
                })
            j += 1

# Guardar resultado
df = pd.DataFrame(datos)
df.to_excel("historia_academica_robusta_con_codigo.xlsx", index=False)
print("✅ Archivo listo: historia_academica_robusta_con_codigo.xlsx")
