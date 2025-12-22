# -*- coding: utf-8 -*-
import re
import fitz  # PyMuPDF
import pandas as pd
import os

# ==============================================================================
# ⚠️ SECCIÓN DE DATOS: PEGA AQUÍ TUS DICCIONARIOS (MALLA, OPTATIVAS, ETC.)
# ==============================================================================

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
                                         'tipo_asignatura': 'Fund. '
                                                            'Obligatoria'},
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


# SI YA TIENES LOS DICCIONARIOS EN ESTE ARCHIVO, NO LOS BORRES.
# SOLO REEMPLAZA EL CÓDIGO QUE SIGUE A PARTIR DE "CONFIGURACIÓN GLOBAL"
# ==============================================================================

# --- CONFIGURACIÓN GLOBAL ---
CARPETA_PDFS = r"C:\Users\JuanPabloGalindoGóme\Documents\Curricular\Curricular\Historial_Academica\activos"  # <--- AJUSTA TU RUTA
ARCHIVO_SALIDA = "Historias_academicas2.xlsx"

def procesar_historias():
    datos = []
    
    # Verificar carpeta
    if not os.path.exists(CARPETA_PDFS):
        print(f"Error: La carpeta {CARPETA_PDFS} no existe.")
        return

    archivos_pdf = [f for f in os.listdir(CARPETA_PDFS) if f.lower().endswith('.pdf')]
    print(f"Encontrados {len(archivos_pdf)} archivos PDF.")

    for archivo in archivos_pdf:
        ruta_pdf = os.path.join(CARPETA_PDFS, archivo)
        print(f"Procesando: {archivo}...")

        try:
            doc = fitz.open(ruta_pdf)
            texto_completo = ""
            for pagina in doc:
                texto_completo += pagina.get_text("text") + "\n"
            doc.close()
        except Exception as e:
            print(f"Error al leer {archivo}: {e}")
            continue

        # --- Limpieza Básica ---
        lineas = texto_completo.split('\n')
        lineas_limpias = [l.strip() for l in lineas if l.strip()]
        texto_unido = "\n".join(lineas_limpias)

        # --- Extracción de Datos del Estudiante ---
        # 1. Nombre
        nombre_match = re.search(r'Nombre:\s*(.+)', texto_unido)
        nombre = nombre_match.group(1).strip() if nombre_match else "Desconocido"

        # 2. Documento
        documento_match = re.search(r'Documento:\s*(\d+)', texto_unido)
        documento = documento_match.group(1).strip() if documento_match else "0"

        # 3. Plan (CORRECCIÓN IMPORTANTE PARA NIVELACIÓN)
        # Busca lo que sigue a (2505)
        plan_match = re.search(r'\(2505\)\s*([^\n]+)', texto_unido)
        if plan_match:
            plan = plan_match.group(1).strip()
        else:
            # Intento secundario por si el formato cambia
            plan_match_simple = re.search(r'Plan:\s*(.+)', texto_unido)
            plan = plan_match_simple.group(1).strip() if plan_match_simple else "Desconocido"

        print(f"   -> Estudiante: {nombre} | Plan: {plan}")

        # --- Procesamiento de Asignaturas ---
        # Unimos líneas para facilitar búsqueda secuencial
        lineas_unidas = [l.strip() for l in lineas if l.strip()]
        
        j = 0
        while j < len(lineas_unidas):
            linea = lineas_unidas[j]
            
            # Regex: Busca "Nombre Materia (CODIGO)"
            # Detecta códigos numéricos largos o con -B (ej: 2015883 o 1000004-B)
            match_asig = re.search(r'(.+?)\s*\((\d{6,7}(?:-B)?)\)', linea)

            if match_asig:
                # 1. Limpieza inicial
                nombre_asig = match_asig.group(1).strip()
                nombre_asig = re.sub(r'^(Obligatoria|Optativa|Libre Elección|Nivelación)\s*\(.\)\s*', '', nombre_asig, flags=re.IGNORECASE)
                codigo = match_asig.group(2).strip()
                # ==============================================================================
                # 🧠 LÓGICA HÍBRIDA: DICCIONARIO + HEURÍSTICA DE TEXTO
                # ==============================================================================
                
                encontrado_en_bd = False
                
                # --- PASO 1: Intentar arreglar usando tus Diccionarios (Lo ideal) ---
                # Revisa todos los diccionarios que tengas disponibles
                listas_asignaturas = []
                if 'malla_curricular' in globals(): listas_asignaturas.append(malla_curricular)
                if 'optativas_produccion' in globals(): listas_asignaturas.append(optativas_produccion)
                if 'asignaturas_extra' in globals(): listas_asignaturas.append(asignaturas_extra)
                
                for diccionario in listas_asignaturas:
                    for nombre_real, info in diccionario.items():
                        if str(info.get('codigo')) == codigo:
                            nombre_asig = nombre_real
                            encontrado_en_bd = True
                            break
                    if encontrado_en_bd: break
                
                # --- PASO 2: Si NO está en diccionarios, intentar unir la siguiente línea ---
                if not encontrado_en_bd:
                    # Verificamos si hay una línea siguiente disponible
                    if j + 1 < len(lineas_unidas):
                        siguiente_linea = lineas_unidas[j + 1].strip()
                        
                        # ANALIZAMOS LA SIGUIENTE LÍNEA:
                        # Si NO tiene formato de código "Nombre (123456)" 
                        # Y NO tiene palabras clave como "Aprobada", "Reprobada" o números de nota
                        es_otra_materia = re.search(r'\((\d{6,7}(?:-B)?)\)', siguiente_linea)
                        es_detalle_nota = re.search(r'(Aprobada|Reprobada|[\d\.]{3,})', siguiente_linea)
                        
                        if not es_otra_materia and not es_detalle_nota:
                            # ¡Es la continuación del nombre!
                            nombre_asig += " " + siguiente_linea
                            # Importante: Avanzamos el índice j para no leer esta línea dos veces
                            j += 1
                # =============================================================

                # Inicializar variables de detalle
                nota = ''
                estado = 'Reprobada'
                anulada = 'NO'
                creditos = ''
                tipo_asig = 'Libre Elección (L)' # Default
                semestre_malla = ''
                semestre_inicio = 'Desconocido' # Puedes mejorar esto extrayendo el encabezado de periodo
                semestre = 'Desconocido' 

                # Buscar semestre (intento simple buscando hacia atrás la fecha tipo 202X-XS)
                # Esto es una mejora opcional, por ahora mantenemos tu lógica de flujo
                
                # Capturar detalles debajo del nombre (nota, creditos, etc)
                detalles = []
                j += 1
                while j < len(lineas_unidas):
                    siguiente = lineas_unidas[j].strip()
                    # Si encontramos OTRA asignatura, paramos
                    if re.search(r'(.+?)\s*\((\d{6,7}(?:-B)?)\)', siguiente):
                        j -= 1
                        break
                    
                    # Detectar Semestre Académico (Encabezado de bloque)
                    # Si encuentras patrones como "2021-1S", guárdalos en una variable externa al while
                    # para asignarlos. Por simplicidad, aquí procesamos detalles de la materia.
                    
                    detalles.append(siguiente)
                    j += 1

                # Analizar detalles
                for detalle in detalles:
                    # Nota y Estado
                    if re.search(r'(Aprobada|Reprobada|SI\*)', detalle):
                        # Extraer nota
                        nota_match = re.search(r'([\d,\.]+)', detalle)
                        if nota_match:
                            nota = nota_match.group(1).replace(',', '.')
                        
                        if 'Aprobada' in detalle: estado = 'Aprobada'
                        elif 'Reprobada' in detalle: estado = 'Reprobada'
                    
                    # Anulada
                    if 'Anulada' in detalle or 'Cancelada' in detalle:
                        anulada = 'SI'

                    # Créditos (si aparecen explícitamente como número solo entre 1 y 6)
                    if creditos == '' and detalle.isdigit() and 0 < int(detalle) <= 20:
                        creditos = int(detalle)
                    
                    # Créditos (con etiqueta)
                    match_credito = re.search(r'[Cc]réditos\s*:?[\s\.]*(\d+)', detalle)
                    if match_credito:
                        creditos = int(match_credito.group(1))

                # Completar datos con Malla Curricular (si no se encontraron en PDF)
                # Usamos el nombre_asig ya corregido
                if 'malla_curricular' in globals():
                    info_malla = malla_curricular.get(nombre_asig)
                    if info_malla:
                        semestre_malla = info_malla.get("semestre", '')
                        tipo_asig = info_malla.get("tipo_asignatura", tipo_asig)
                        if creditos == '': creditos = info_malla.get("creditos", '')

                # Completar con Optativas
                if 'optativas_produccion' in globals() and not semestre_malla:
                    info_opt = optativas_produccion.get(nombre_asig)
                    if info_opt:
                        semestre_malla = info_opt.get("semestre", '')
                        tipo_asig = info_opt.get("tipo_asignatura", tipo_asig)
                        if creditos == '': creditos = info_opt.get("creditos", '')

                # Debug en consola para verificar correcciones
                # print(f"Procesado: {nombre_asig} ({codigo}) - Nota: {nota}")

                datos.append({
                    'nombre': nombre,
                    'documento': documento,
                    'plan': plan,  # <--- CAMPO CLAVE
                    'codigo_asignatura': codigo,
                    'asignatura': nombre_asig, # Nombre corregido
                    'creditos': creditos,
                    'tipo_asignatura': tipo_asig,
                    'semestre_malla': semestre_malla,
                    'nota': float(nota) if str(nota).replace('.', '', 1).isdigit() else 0.0,
                    'estado': estado,
                    'anulada': anulada,
                    'semestre_inicio': semestre_inicio, # Ajustar si tienes lógica de periodos
                    'semestre_asignatura': semestre
                })
            
            j += 1

    # Exportar
    if datos:
        df = pd.DataFrame(datos)
        try:
            df.to_excel(ARCHIVO_SALIDA, index=False)
            print(f"\n✅ Éxito: Archivo guardado en {ARCHIVO_SALIDA}")
        except Exception as e:
            print(f"\n❌ Error al guardar Excel: {e}")
    else:
        print("\n⚠️ No se encontraron datos para exportar.")

if __name__ == "__main__":
    procesar_historias()