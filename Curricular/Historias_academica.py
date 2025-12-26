# -*- coding: utf-8 -*-
"""
Script para procesar archivos PDF de historiales académicos de estudiantes.

Este script realiza las siguientes tareas:
1.  Lee todos los archivos PDF de una carpeta especificada.
2.  Extrae el texto de cada PDF.
3.  Limpia el texto eliminando encabezados, pies de página y otra información irrelevante.
4.  Extrae el nombre y el documento de identidad del estudiante.
5.  Divide el historial en bloques por cada período académico (semestre).
6.  Reconstruye los nombres de las asignaturas que pueden estar divididos en varias líneas.
7.  Extrae la información detallada de cada asignatura: código, nombre, nota, estado (aprobada/reprobada),
    si fue anulada, y los créditos.
8.  Enriquece los datos de las asignaturas utilizando diccionarios predefinidos (malla curricular,
    optativas, etc.) para obtener el semestre sugerido en la malla y el tipo de asignatura.
9.  Almacena toda la información extraída en una lista de diccionarios.
10. Exporta los datos consolidados a un archivo Excel.
"""

# --- Importación de librerías ---
import re           # Para búsquedas y manipulaciones con expresiones regulares
import fitz         # PyMuPDF: para la extracción de texto desde archivos PDF
import pandas as pd # Para el manejo de estructuras de datos tabulares (DataFrame)
import os           # Para interactuar con el sistema de archivos (navegar carpetas)

# --- CONFIGURACIÓN GLOBAL ---

# Palabras clave para identificar líneas de encabezado en las tablas de asignaturas.
# Ayuda a evitar que estas líneas se confundan con nombres de asignaturas.
encabezado_claves = ['asignatura', 'créditos', 'hap', 'hai', 'ths', 'tipología', 'calificación', 'anulada', 'n. veces']

# Diccionario con textos genéricos e innecesarios que se encuentran comúnmente en los PDF.
# Estos textos se eliminarán durante la fase de limpieza.
basura = {
    0: 'Abreviaturas utilizadas: HAB=Habilitación, VAL=Validación por Pérdida, SUF=Validación por Suficiencia, HAP=Horas de Actividad Presencial, HAI=Horas de Actividad',
    1: 'Independiente, THS=Total Horas Semanales, HOM=Homologada o Convalidada.',
    2: 'SI*: Cancelación por decisión de la universidad soportada en acuerdos, resoluciones y actos académicos',
    3: 'Este es un documento de uso interno de la Universidad Nacional de Colombia. No constituye, ni reemplaza el certificado oficial de notas.',
    4: 'Informe generado por el usuario:',
    5: 'Reporte de Historia Académica',
    6: 'Sistema de Información Académica',
    7: 'Dirección Nacional de Información Académica',
    8: 'Registro y Matrícula',
    9: 'jugalindog el Friday, December 19, 2025',
    10: 'Créditos HAP'

}

# --- DICCIONARIOS DE ASIGNATURAS (MALLA CURRICULAR) ---
# Estos diccionarios actúan como una base de datos para enriquecer la información
# extraída del PDF, como el semestre ideal, los créditos y la tipología de cada asignatura.

# Malla curricular principal: contiene las asignaturas obligatorias y de fundamentación.
malla_curricular = {'Agroclimatología': {'codigo': '2015880',
                      'creditos': 3,
                      'semestre': 4,
                      'tipo_asignatura': 'Disciplinar'},
 'Agroecosistemas y Sistemas de Producción': {'codigo': '2015881',
                                              'creditos': 3,
                                              'semestre': 8,
                                              'tipo_asignatura': 'Disciplinar'},
 'Bioestadística fundamental': {'codigo': '1000012-B',
                                'creditos': 3,
                                'semestre': 3,
                                'tipo_asignatura': 'Fund. Obligatoria'},
 'Biología Celular y Molecular Básica': {'codigo': '2015882',
                                         'creditos': 3,
                                         'semestre': 4,
                                         'tipo_asignatura': 'Fund. Obligatoria'},
 'Biología de plantas': {'codigo': '2015877',
                         'creditos': 3,
                         'semestre': 1,
                         'tipo_asignatura': 'Fund. Obligatoria'},
 'Bioquímica básica': {'codigo': '1000042-B',
                       'creditos': 3,
                       'semestre': 3,
                       'tipo_asignatura': 'Fund. Obligatoria'},
 'Botánica taxonómica': {'codigo': '2015878',
                         'creditos': 3,
                         'semestre': 2,
                         'tipo_asignatura': 'Fund. Obligatoria'},
 'Ciclo  II: Ejecución de un proyecto productiv': {'codigo': '2015884',
                                                   'creditos': 3,
                                                   'semestre': 9,
                                                   'tipo_asignatura': 'Disciplinar'},
 'Ciclo i: formulación y evaluación de proyect': {'codigo': '2015883',
                                                  'creditos': 3,
                                                  'semestre': 8,
                                                  'tipo_asignatura': 'Disciplinar'},
 'Ciencia del suelo': {'codigo': '2015885',
                       'creditos': 3,
                       'semestre': 3,
                       'tipo_asignatura': 'Disciplinar'},
 'Cálculo Integral': {'codigo': '1000005-B',
                      'creditos': 4,
                      'semestre': 2,
                      'tipo_asignatura': 'Fund. Obligatoria'},
 'Cálculo diferencial': {'codigo': '1000004-B',
                         'creditos': 4,
                         'semestre': 1,
                         'tipo_asignatura': 'Fund. Obligatoria'},
 'Diseño de experimentos': {'codigo': '2015887',
                            'creditos': 3,
                            'semestre': 4,
                            'tipo_asignatura': 'Fund. Obligatoria'},
 'Economía agraria': {'codigo': '2015888',
                      'creditos': 3,
                      'semestre': 6,
                      'tipo_asignatura': 'Disciplinar'},
 'Edafología': {'codigo': '2015889',
                'creditos': 3,
                'semestre': 4,
                'tipo_asignatura': 'Disciplinar'},
 'Entomología': {'codigo': '2015890',
                 'creditos': 3,
                 'semestre': 6,
                 'tipo_asignatura': 'Disciplinar'},
 'Fisiología de la producción vegetal': {'codigo': '2015891',
                                         'creditos': 3,
                                         'semestre': 6,
                                         'tipo_asignatura': 'Disciplinar'},
 'Fisiología vegetal básica': {'codigo': '2015892',
                               'creditos': 3,
                               'semestre': 5,
                               'tipo_asignatura': 'Disciplinar'},
 'Fitomejoramiento': {'codigo': '2015893',
                      'creditos': 3,
                      'semestre': 8,
                      'tipo_asignatura': 'Disciplinar'},
 'Fitopatología': {'codigo': '2015894',
                   'creditos': 3,
                   'semestre': 6,
                   'tipo_asignatura': 'Disciplinar'},
 'Fundamentos de ecología': {'codigo': '1000011-B',
                             'creditos': 3,
                             'semestre': 4,
                             'tipo_asignatura': 'Fund. Obligatoria'},
 'Fundamentos de mecánica': {'codigo': '1000019-B',
                             'creditos': 3,
                             'semestre': 2,
                             'tipo_asignatura': 'Fund. Obligatoria'},
 'Geomática básica': {'codigo': '2015896',
                      'creditos': 3,
                      'semestre': 3,
                      'tipo_asignatura': 'Fund. Obligatoria'},
 'Gestión agroempresarial': {'codigo': '2015922',
                             'creditos': 3,
                             'semestre': 7,
                             'tipo_asignatura': 'Disciplinar'},
 'Génetica general': {'codigo': '2015895',
                      'creditos': 3,
                      'semestre': 5,
                      'tipo_asignatura': 'Fund. Obligatoria'},
 'Inglés I- Semestral': {'codigo': '1000044-B',
                         'creditos': 2,
                         'semestre': 1,
                         'tipo_asignatura': 'Nivelación'},
 'Inglés II - Semestral': {'codigo': '1000045-B',
                           'creditos': 2,
                           'semestre': 2,
                           'tipo_asignatura': 'Nivelación'},
 'Inglés III - Semestral': {'codigo': '1000046-B',
                            'creditos': 2,
                            'semestre': 3,
                            'tipo_asignatura': 'Nivelación'},
 'Inglés IV- Semestral': {'codigo': '1000047-B',
                           'creditos': 2,
                           'semestre': 4,
                           'tipo_asignatura': 'Nivelación'},
 'Introducción a la ingeniería agronómica': {'codigo': '2015897',
                                             'creditos': 2,
                                             'semestre': 1,
                                             'tipo_asignatura': 'Disciplinar'},
 'Laboratorio de bioquímica básica': {'codigo': '1000043-B',
                                   'creditos': 2,
                                   'semestre': 3,
                                   'tipo_asignatura': 'Fund. Obligatoria'},
 'Laboratorio de química básica': {'codigo': '2015782',
                                   'creditos': 2,
                                   'semestre': 2,
                                   'tipo_asignatura': 'Fund. Obligatoria'},
 'Lecto-Escritura': {'codigo': '1000002-B',
                     'creditos': 2,
                     'semestre': 1,
                     'tipo_asignatura': 'Nivelación'},
 'Manejo Integrado de Enfermedades': {'codigo': '2015899',
                                      'creditos': 3,
                                      'semestre': 7,
                                      'tipo_asignatura': 'Disciplinar'},
 'Manejo de la fertilidad del suelo': {'codigo': '2015898',
                                       'creditos': 3,
                                       'semestre': 7,
                                       'tipo_asignatura': 'Disciplinar'},
 'Manejo integrado de malezas': {'codigo': '2015900',
                                 'creditos': 3,
                                 'semestre': 7,
                                 'tipo_asignatura': 'Disciplinar'},
 'Manejo integrado de plagas': {'codigo': '2015901',
                                'creditos': 3,
                                'semestre': 7,
                                'tipo_asignatura': 'Disciplinar'},
 'Matemáticas Básicas': {'codigo': '1000001-B',
                         'creditos': 3,
                         'semestre': 1,
                         'tipo_asignatura': 'Nivelación'},
 'Mecanización agrícola': {'codigo': '2015902',
                           'creditos': 3,
                           'semestre': 5,
                           'tipo_asignatura': 'Disciplinar'},
 'Microbiología': {'codigo': '2015903',
                   'creditos': 3,
                   'semestre': 4,
                   'tipo_asignatura': 'Fund. Obligatoria'},
 'Práctica Profesional': {'codigo': '2015934',
                          'creditos': 6,
                          'semestre': 10,
                          'tipo_asignatura': 'Disciplinar'},
 'Química básica': {'codigo': '1000041-B',
                    'creditos': 3,
                    'semestre': 1,
                    'tipo_asignatura': 'Fund. Obligatoria'},
 'Reproducción y multiplicación': {'codigo': '2015907',
                                   'creditos': 3,
                                   'semestre': 6,
                                   'tipo_asignatura': 'Disciplinar'},
 'Riegos y drenajes': {'codigo': '2015908',
                       'creditos': 3,
                       'semestre': 5,
                       'tipo_asignatura': 'Disciplinar'},
 'Sociología Rural': {'codigo': '2015909',
                      'creditos': 2,
                      'semestre': 5,
                      'tipo_asignatura': 'Disciplinar'},
 'Tecnología de la Poscosecha': {'codigo': '2015910',
                                 'creditos': 3,
                                 'semestre': 8,
                                 'tipo_asignatura': 'Disciplinar'},
 'Trabajo de Grado': {'codigo': '2015291',
                      'creditos': 6,
                      'semestre': 10,
                      'tipo_asignatura': 'Disciplinar'}}

# Asignaturas optativas de producción.
optativas_produccion = {
    "Produccion de cultivos de clima calido": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de frutales":            {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de hortalizas":          {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de ornamentales":        {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Cultivos perennes industriales":    {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de papa":                {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
}

# Asignaturas extra o de libre elección comunes.
asignaturas_extra = {
    "Agroecología":                 {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Agrobiodiversidad":            {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Bioprocesos Agroalimentarios": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Computación estadística":      {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Desarrollo Rural":             {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Emprendimiento e innovación en agronegocios": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Evolución y ecología de patógenos de plantas":{"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Fundamentos de Agroindustria": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Genética de Insectos de Interés económico":   {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Gestión ambiental agropecuaria":              {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Investigación de Mercados":    {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Nutrición Mineral de Plantas": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Producción de cannabis medicinal":            {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Suelos vivos":                 {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Sistemas Agroalimentarios Vinculo entre ambiente, sociedad y desarrollo": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"}
}

# Asignaturas de posgrado que pueden ser tomadas por estudiantes de pregrado.
asignaturas_posgrado = {
    "Agroclimatología y cambio climático": {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Agua y nutrición mineral":            {"codigo": "2019978", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Biología de suelos":                  {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Biología molecular":                  {"codigo": "2019986", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Biología y ecología de malezas":      {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Clínica de plantas":                  {"codigo": "2026913", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Decisiones de manejo fitosanitario: aproximación práctica": {"codigo": "2028521", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Degradación química del suelo":       {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fertilizantes y fertilización":       {"codigo": "2019589", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Física de suelos":                    {"codigo": "2020742", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fisiología avanzada en frutales": {"codigo": "2020001", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fisiología de cultivos":          {"codigo": "2028756", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fisiología del desarrollo":       {"codigo": "2020004", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fitopatología avanzada":          {"codigo": "2020007", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Genética avanzada":               {"codigo": "2020009", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Hongos y nemátodos fitopatógenos": {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Métodos multivariados":           {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Pedología":                       {"codigo": "2020745", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Recursos genéticos vegetales":    {"codigo": "2020046", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Taxonomía de insectos":           {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Desarrollo económico del territorio rural": {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Desarrollo rural y territorios":  {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Economía de la empresa agraria y alimentaria": {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Gestión contable financiera":     {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Gestión de agroproyectos":        {"codigo": "2025414", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Mercadeo agroalimentario y territorial": {"codigo": "2026250", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Problemas agrarios colombianos":  {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Sociedad e instituciones rurales": {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Geoestadística":                 {"codigo": "2020012", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Geomática general":              {"codigo": "2020764", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Geoprocesamiento":               {"codigo": None, "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Percepción remota":              {"codigo": "2020039", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Programación sig":               {"codigo": "2027945", "semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
}




#CARPETA_PDFS = "/home/jugalindog/Documents/Historias academicas/activos"
CARPETA_PDFS = "/home/jugalindog/Pasantia/Curricular/Curricular/Historial_Academica/activos"
#CARPETA_PDFS = "C:\\Users\\JuanPabloGalindoGóme\Documents\\Curricular\\Curricular\\Historial_Academica\\activos"
                
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
    Semestre_inicio_match = re.search(r'Periodo admisión:\s*(\d+)', texto)
    plan_match = re.search(r'\(2505\)\s*([^\n]+)', texto)
    

    if not nombre_match or not documento_match:
        continue

    nombre = nombre_match.group(1).strip()
    plan = plan_match.group(1).strip() if plan_match else "Desconocido"
    documento = documento_match.group(1).strip()
    semestre_inicio =Semestre_inicio_match.group(1).strip()

    bloques = re.split(r'(?:PRIMER|SEGUNDO)\s+PERIODO\s+(\d{4}-[12]S)', texto)

    for i in range(1, len(bloques), 2):
        semestre = bloques[i]
        contenido = bloques[i + 1]
        lineas = [l.strip() for l in contenido.splitlines() if l.strip()]

        lineas_unidas = []
        j = 0
        while j < len(lineas):
            actual = lineas[j].strip()
            print(actual)
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
                nombre_asig = re.sub(r'^(Obligatoria|Optativa|Libre Elección|Nivelación)\s*\(.\)\s*', '', nombre_asig, flags=re.IGNORECASE)
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
                
#                print(detalles)
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
                
                info_extra = asignaturas_extra.get(nombre_asig)
                if info_extra:
                    tipo_asig = info_extra["tipo_asignatura"]
                    if creditos == '':
                        creditos = info_extra.get("Creditos", info_extra.get("creditos", 3))
                    if semestre_malla == '':
                         semestre_malla = info_extra.get("semestre", None)
                
                info_posgrado = asignaturas_posgrado.get(nombre_asig)
                if info_posgrado:
                    tipo_asig = info_posgrado["tipo_asignatura"]
                    if creditos == '':
                        creditos = info_posgrado.get("Creditos", info_posgrado.get("creditos", 4))
                    if semestre_malla == '':
                        semestre_malla = info_posgrado.get("semestre", None)


                if creditos == '':
                    print(f"⚠️ Estudiante : {nombre} ")
                    print(f"⚠️ Créditos no encontrados para: {nombre_asig} ({codigo})")
                    print("🧾 Detalles:", detalles)

                datos.append({
                    'nombre': nombre,
                    'documento': documento,
                    'plan': plan,
                    'codigo_asignatura': codigo,
                    'asignatura': nombre_asig,
                    'creditos': creditos,
                    'tipo_asignatura': tipo_asig,
                    'semestre_malla': semestre_malla,
                    'nota': float(nota) if nota.replace('.', '', 1).isdigit() else 0.0,
                    'estado': estado,
                    'anulada': anulada,
                    'semestre_inicio': semestre_inicio,
                    'semestre_asignatura': semestre
                })
            j += 1


# Exportar a Excel
df = pd.DataFrame(datos)
df.to_excel("Historias_academicas.xlsx", index=False)
print("Archivo generado correctamente.")
