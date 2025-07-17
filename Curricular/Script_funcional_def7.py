# --- IMPORTACIÓN DE LIBRERÍAS ---
import re           # Para búsquedas y manipulaciones con expresiones regulares
import fitz         # PyMuPDF: extracción de texto desde archivos PDF
import pandas as pd # Manejo de estructuras tabulares
import os           # Navegación del sistema de archivos

# Diccionario con textos innecesarios comunes en los PDF que deben eliminarse
basura = {
    0: 'Abreviaturas utilizadas: HAB=Habilitación, VAL=Validación por Pérdida, SUF=Validación por Suficiencia, HAP=Horas de Actividad Presencial, HAI=Horas de Actividad',
    1: 'Independiente, THS=Total Horas Semanales, HOM=Homologada o Convalidada.',
    2: 'SI*: Cancelación por decisión de la universidad soportada en acuerdos, resoluciones y actos académicos',
    3: 'Este es un documento de uso interno de la Universidad Nacional de Colombia. No constituye, ni reemplaza el certificado oficial de notas.',
    4: 'Informe generado por el usuario:',
    5: 'Reporte de Historia Académica',
    6: 'Sistema de Información Académica',
    7: 'Dirección Nacional de Información Académica',
    8: 'Registro y Matrícula'
}
# Ruta de la carpeta que contiene los archivos PDF a procesar

#CARPETA_PDFS = "C:\\Users\\jp2g\\Documents\\PASANTIA\\Curricular\\Curricular\\Historial_Academica"
CARPETA_PDFS = "/home/jugalindog/Pasantia/Curricular/Curricular/Historial_Academica"

# Aquí se almacenarán los resultados para construir el DataFrame
datos = [] 

# Iterar sobre todos los archivos de la carpeta
for archivo in os.listdir(CARPETA_PDFS):
    if not archivo.endswith(".pdf"):
        continue

    ruta_pdf = os.path.join(CARPETA_PDFS, archivo)
    try:
        # Extracción de texto completo desde todas las páginas del PDF
        doc = fitz.open(ruta_pdf)
        texto = "\n".join([page.get_text() for page in doc])
        doc.close()

        # Limpieza del texto extraído
        texto = re.sub(
            r"Informe generado por el usuario:\s+\S+\s+el\s+\w+\s+\d{1,2}\s+de\s+\w+\s+de\s+\d{4}\s+\d{2}:\d{2}",
            '', texto)
        texto = re.sub(r'Página\xa0\d+\xa0de\xa0\d+', '', texto)
        texto = re.sub(r'\n?[A-ZÁÉÍÓÚÑ][^\n]+\s+-\s+\d{7,10}', '', texto)
        # Eliminar textos innecesarios definidos en el diccionario 'basura'
        for b in basura.values():
            texto = texto.replace(b, '')

    except Exception as e:
        print(f"Error con {archivo}: {e}")  # Muestra error si falla la lectura del PDF
        continue  # Salta al siguiente archivo

    # Extraer nombre y documento del texto
    # Usamos expresiones regulares para encontrar el nombre y el documento
    nombre_match = re.search(r'Nombre:\s*(.+)', texto)  # Busca el nombre del estudiante
    documento_match = re.search(r'Documento:\s*(\d+)', texto)  # Busca el número de documento
    if not nombre_match or not documento_match:
        continue  # Si no encuentra nombre o documento, salta al siguiente archivo
    
    # Si se encuentran, extraemos y limpiamos los datos
    nombre = nombre_match.group(1).strip()  # Obtiene el nombre limpio
    documento = documento_match.group(1).strip()  # Obtiene el documento limpio

# Dividir el texto en bloques por semestres
bloques = re.split(r'(?:PRIMER|SEGUNDO)\s+PERIODO\s+(\d{4}-[12]S)', texto)  # Separa el texto por semestres

for i in range(1, len(bloques), 2):  # Itera sobre los bloques de semestre
    semestre = bloques[i]  # El semestre actual
    contenido = bloques[i + 1]  # Contenido del semestre
    lineas = [l.strip() for l in contenido.splitlines() if l.strip()]  # Filtra líneas vacías y limpia espacios

    # Procesar cada línea del bloque
    lineas_unidas = []  # Lista para almacenar líneas procesadas

    j = 0
    while j < len(lineas):  # Itera sobre las líneas del semestre
        actual = lineas[j].strip()  # Limpia la línea actual
        match_codigo = None  # Inicializa variable para coincidencia de código
        codigo = None  # Inicializa variable para código de asignatura

        # Si la línea es solo el código de asignatura entre paréntesis
        if re.fullmatch(r'\((\d{6,7}(?:-B)?)\)', actual):
            codigo = re.findall(r'\((\d{6,7}(?:-B)?)\)', actual)[0]  # Extrae el código
            if j > 0:
                nombre_candidato = lineas[j - 1].strip()  # Toma la línea anterior como posible nombre
                encabezado_claves = ['asignatura', 'créditos', 'hap', 'hai', 'ths', 'tipología', 'calificación', 'anulada', 'n. veces']  # Palabras clave de encabezado
                if not any(p in nombre_candidato.lower() for p in encabezado_claves):  # Si no es encabezado
                    match_codigo = re.match(r'(.+)', nombre_candidato)  # Coincidencia para nombre
                    actual = f"{nombre_candidato} ({codigo})"  # Une nombre y código
                    j += 1  # Avanza el índice

        # Si la línea tiene nombre y código juntos
        elif re.search(r'(.+)\s\((\d{6,7}(?:-B)?)\)$', actual):
            match_codigo = re.search(r'(.+)\s\((\d{6,7}(?:-B)?)\)$', actual)  # Coincidencia para nombre y código
            codigo = match_codigo.group(2)  # Extrae el código

        if match_codigo:  # Si se encontró nombre y código
            texto_previo = match_codigo.group(1).lower()  # Obtiene el nombre en minúsculas
            encabezado_claves = ['asignatura', 'créditos', 'hap', 'hai', 'ths', 'tipología', 'calificación', 'anulada', 'n. veces']  # Palabras clave de encabezado

            if any(p in texto_previo for p in encabezado_claves):  # Si es encabezado
                lineas_unidas.append(actual)  # Agrega la línea tal cual
            else:
                nombre_final = match_codigo.group(1).strip()  # Obtiene el nombre limpio
                nombre_partes = [nombre_final]  # Inicializa lista de partes del nombre
                k = j - 1
                while k >= 0:
                    anterior = lineas[k].strip().lower()  # Toma la línea anterior en minúsculas
                    if re.fullmatch(r'\d+', anterior):  # Si es solo un número, termina
                        break
                    if any(p in anterior for p in encabezado_claves):  # Si es encabezado, termina
                        break
                    nombre_partes.insert(0, lineas[k].strip())  # Agrega la parte al inicio
                    k -= 1
                nombre_completo = " ".join(nombre_partes) + f" ({codigo})"  # Une todas las partes y el código
                lineas_unidas = lineas_unidas[:k + 1]  # Elimina las partes ya usadas
                lineas_unidas.append(nombre_completo)  # Agrega el nombre completo unido
        else:
            lineas_unidas.append(actual)  # Si no hay código, agrega la línea tal cual
        j += 1  # Avanza el índice

    j = 0
    while j < len(lineas_unidas):  # Itera sobre las líneas unidas
        linea = lineas_unidas[j]  # Toma la línea actual
        match_asig = re.search(r'(.+?)\s*\((\d{6,7}(?:-B)?)\)', linea)  # Busca nombre y código de asignatura
        if match_asig:
            nombre_asig = match_asig.group(1).strip()  # Extrae el nombre de la asignatura
            codigo = match_asig.group(2).strip()  # Extrae el código de la asignatura
            tipo_asig = ''  # Inicializa tipo de asignatura
            nota = ''  # Inicializa nota
            estado = 'Reprobada'  # Inicializa estado
            anulada = 'NO'  # Inicializa anulada

            # Detecta el tipo de asignatura en el nombre
            for tipop in ['Obligatoria (C)', 'Fund. Obligatoria', 'Fund. Optativa', 'Disciplinar', 'Libre Elección (L)', 'Nivelación (E)', 'Optativa (T)']:
                if tipop in nombre_asig:
                    tipo_asig = tipop  # Asigna el tipo
                    nombre_asig = nombre_asig.replace(tipop, '').strip()  # Limpia el nombre
                    break

            detalles = []
            j += 1
            while j < len(lineas_unidas):
                siguiente = lineas_unidas[j].strip()
                if re.search(r'(.+?)\s*\((\d{6,7}(?:-B)?)\)', siguiente):
                    j -= 1
                    break
                detalles.append(siguiente)
                j += 1

            # Procesa los detalles para nota, estado, anulada y tipo
            for detalle in detalles:
                if re.search(r'(Aprobada|Reprobada|SI\*)', detalle):
                    nota_match = re.search(r'([\d,\.]+)', detalle)
                    if nota_match:
                        nota = nota_match.group(1).replace(',', '.')
                    estado = 'Aprobada' if 'Aprobada' in detalle else 'Reprobada'
                if 'Anulada' in detalle or 'SI' in detalle:
                    anulada = 'SI'
                if any(t in detalle for t in ['Obligatoria', 'Optativa', 'Libre Elección', 'Nivelación']):
                    tipo_asig = detalle

            # Agrega los datos de la asignatura al array principal
            datos.append({
                'nombre': nombre,
                'documento': documento,
                'codigo_asignatura': codigo,
                'asignatura': nombre_asig,
                'tipo_asignatura': tipo_asig,
                'nota': float(nota) if nota.replace('.', '', 1).isdigit() else 0.0,
                'estado': estado,
                'anulada': anulada,
                'semestre_inicio': '2018-2S',
                'semestre_asignatura': semestre
            })
        j += 1  # Avanza el índice

# Exportar resultado
df = pd.DataFrame(datos)  # Crea el DataFrame con los datos recolectados
df.to_excel("Script_funcional_def7.xlsx", index=False)  # Exporta a Excel
print("✅ Archivo listo: H_A_R_op7.xlsx")
