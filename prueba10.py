# --- Importación de librerías ---
import re
import fitz  # PyMuPDF
import pandas as pd
import os

# --- Claves para identificar encabezados (mueven arriba) ---
encabezado_claves = ['asignatura', 'créditos', 'hap', 'hai', 'ths', 'tipología', 'calificación', 'anulada', 'n. veces']

# --- Basura a eliminar ---
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

# --- Malla curricular con créditos y tipo ---
malla_curricular = {
   "Introducción a la ingeniería agronómica": {"semestre": 1, "creditos": 2, "tipo_asignatura": "Disciplinar"},
    "Matemáticas Básicas":                     {"semestre": 1, "creditos": 3, "tipo_asignatura": "Nivelación"},
    "Biología de plantas":                     {"semestre": 1, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Lecto-Escritura":                         {"semestre": 1, "creditos": 2, "tipo_asignatura": "Nivelación"},
    "Química básica":                          {"semestre": 1, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Cálculo diferencial":                  {"semestre": 1, "creditos": 4, "tipo_asignatura": "Fund. Obligatoria"},
    "Cálculo Integral":                     {"semestre": 2, "creditos": 4, "tipo_asignatura": "Fund. Obligatoria"},
    "Fundamentos de mecánica":              {"semestre": 2, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Botánica taxonómica":                  {"semestre": 2, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Laboratorio de química básica":        {"semestre": 2, "creditos": 1, "tipo_asignatura": "Fund. Obligatoria"},
    "Ciencia del suelo":                    {"semestre": 3, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Laboratorio bioquímica básica":        {"semestre": 3, "creditos": 1, "tipo_asignatura": "Fund. Obligatoria"},
    "Bioquímica básica":                    {"semestre": 3, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Bioestadística fundamental":           {"semestre": 3, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Geomática básica":                     {"semestre": 3, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Agroclimatología":                     {"semestre": 4, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Edafología":                           {"semestre": 4, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Fundamentos de ecología":              {"semestre": 4, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Microbiología":                        {"semestre": 4, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Biología Celular y Molecular Básica":  {"semestre": 4, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Diseño de experimentos":               {"semestre": 4, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Sociología Rural":                     {"semestre": 5, "creditos": 2, "tipo_asignatura": "Disciplinar"},
    "Riegos y drenajes":                    {"semestre": 5, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Mecanización agrícola":                {"semestre": 5, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Génetica general":                     {"semestre": 5, "creditos": 3, "tipo_asignatura": "Fund. Obligatoria"},
    "Fisiología vegetal básica":            {"semestre": 5, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Economía agraria":                     {"semestre": 6, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Entomología":                          {"semestre": 6, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Fitopatología":                        {"semestre": 6, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Fisiología de la producción vegetal":  {"semestre": 6, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Reproducción y multiplicación":      {"semestre": 6, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Gestión agroempresarial":            {"semestre": 7, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Manejo de la fertilidad del suelo":  {"semestre": 7, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Manejo integrado de plagas":         {"semestre": 7, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Manejo Integrado de Enfermedades":   {"semestre": 7, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Manejo integrado de malezas":        {"semestre": 7, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Ciclo i: formulación y evaluación de proyect": {"semestre": 8, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Fitomejoramiento":                             {"semestre": 8, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Agroecosistemas y Sistemas de Producción":     {"semestre": 8, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Tecnología de la Poscosecha":                  {"semestre": 8, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Ciclo  II: Ejecución de un proyecto productiv":{"semestre": 9, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Práctica Profesional":              {"semestre": 10, "creditos": 6, "tipo_asignatura": "Disciplinar"},
    "Trabajo de Grado":                  {"semestre": 10, "creditos": 6, "tipo_asignatura": "Disciplinar"}
}

optativas_produccion = {
    "Produccion de cultivos de clima calido": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de frutales":            {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de hortalizas":          {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de ornamentales":        {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Cultivos perennes industriales":    {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de papa":                {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
}

CARPETA_PDFS = "/home/jugalindog/Pasantia/Curricular/Curricular/Historial_Academica"
datos = []

# --- Procesamiento de PDFs ---
for archivo in os.listdir(CARPETA_PDFS):
    if not archivo.endswith(".pdf"):
        continue

    ruta_pdf = os.path.join(CARPETA_PDFS, archivo)
    try:
        doc = fitz.open(ruta_pdf)
        texto = "\n".join([page.get_text() for page in doc])
        doc.close()

        for b in basura.values():
            texto = texto.replace(b, '')

        texto = re.sub(r"Informe generado.*\d{2}:\d{2}", '', texto)
        texto = re.sub(r'Página\xa0\d+\xa0de\xa0\d+', '', texto)
        texto = re.sub(r'\n?[A-ZÁÉÍÓÚÑ][^\n]+\s+-\s+\d{7,10}', '', texto)

    except Exception as e:
        print(f"❌ Error con {archivo}: {e}")
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
        lineas = [l.strip() for l in contenido.splitlines() if l.strip()]

        lineas_unidas = []
        j = 0
        while j < len(lineas):
            actual = lineas[j].strip()
            match_codigo = None
            codigo = None

            if re.fullmatch(r'\((\d{6,7}(?:-B)?)\)', actual):
                codigo = re.findall(r'\((\d{6,7}(?:-B)?)\)', actual)[0]
                if j > 0:
                    nombre_candidato = lineas[j - 1].strip()
                    if not any(p in nombre_candidato.lower() for p in encabezado_claves):
                        actual = f"{nombre_candidato} ({codigo})"
                        j += 1
            elif re.search(r'(.+)\s\((\d{6,7}(?:-B)?)\)$', actual):
                match_codigo = re.search(r'(.+)\s\((\d{6,7}(?:-B)?)\)$', actual)
                codigo = match_codigo.group(2)

            if match_codigo:
                nombre_final = match_codigo.group(1).strip()
                nombre_partes = [nombre_final]
                k = j - 1
                while k >= 0:
                    anterior = lineas[k].strip().lower()
                    if re.fullmatch(r'\d+', anterior): break
                    if any(p in anterior for p in encabezado_claves): break
                    nombre_partes.insert(0, lineas[k].strip())
                    k -= 1
                nombre_completo = " ".join(nombre_partes) + f" ({codigo})"
                lineas_unidas = lineas_unidas[:k + 1]
                lineas_unidas.append(nombre_completo)
            else:
                lineas_unidas.append(actual)
            j += 1

        # --- Extracción por asignatura ---
        j = 0
        while j < len(lineas_unidas):
            linea = lineas_unidas[j]
            match_asig = re.search(r'(.+?)\s*\((\d{6,7}(?:-B)?)\)', linea)
            if match_asig:
                nombre_asig = match_asig.group(1).strip()
                codigo = match_asig.group(2).strip()
                nota = ''
                estado = 'Reprobada'
                anulada = 'NO'
                creditos = ''
                tipo_detectado = ''

                detalles = []
                j += 1
                while j < len(lineas_unidas):
                    siguiente = lineas_unidas[j].strip()
                    if re.search(r'(.+?)\s*\((\d{6,7}(?:-B)?)\)', siguiente):
                        j -= 1
                        break
                    detalles.append(siguiente)
                    j += 1

                for detalle in detalles:
                    if re.search(r'(Aprobada|Reprobada|SI\*)', detalle):
                        nota_match = re.search(r'([\d,\.]+)', detalle)
                        if nota_match:
                            nota = nota_match.group(1).replace(',', '.')
                        estado = 'Aprobada' if 'Aprobada' in detalle else 'Reprobada'
                    if 'Anulada' in detalle or 'SI' in detalle:
                        anulada = 'SI'
                    if creditos == '' and detalle.isdigit() and 0 < int(detalle) <= 6:
                        creditos = int(detalle)
                    if creditos == '':
                        match_credito = re.search(r'[Cc]réditos\s*:?[\s\.]*(\d+)', detalle)
                        if match_credito:
                            creditos = int(match_credito.group(1))
                    if any(t in detalle for t in ['Obligatoria', 'Optativa', 'Libre Elección', 'Nivelación']):
                        tipo_detectado = detalle

                info_malla = malla_curricular.get(nombre_asig)
                if info_malla:
                    semestre_malla = info_malla["semestre"]
                    if creditos == '':
                        creditos = info_malla["creditos"]
                    tipo_asig = info_malla["tipo_asignatura"]
                else:
                    semestre_malla = ''
                    tipo_asig = 'Libre Elección (L)'

                info_optativa = optativas_produccion.get(nombre_asig)
                if info_optativa:
                    tipo_asig = info_optativa["tipo_asignatura"]
                    if creditos == '':
                        creditos = info_optativa["creditos"]
                    if semestre_malla == '':
                        semestre_malla = info_optativa["semestre"]

                if creditos == '':
                    print(f"⚠️ Créditos no encontrados para: {nombre_asig} ({codigo})")
                    print("🧾 Detalles:", detalles)

                datos.append({
                    'nombre': nombre,
                    'documento': documento,
                    'codigo_asignatura': codigo,
                    'asignatura': nombre_asig,
                    'creditos': creditos,
                    'tipo_asignatura': tipo_asig,
                    'semestre_malla': semestre_malla,
                    'nota': float(nota) if nota.replace('.', '', 1).isdigit() else 0.0,
                    'estado': estado,
                    'anulada': anulada,
                    'semestre_inicio': '2018-2S',
                    'semestre_asignatura': semestre
                })
            j += 1


# Exportar a Excel
df = pd.DataFrame(datos)
df.to_excel("Prueba10_con_creditos.xlsx", index=False)
print("✅ Archivo generado correctamente.")
