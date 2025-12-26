import pdfplumber
import pandas as pd
import re
import os
import glob
from tqdm import tqdm
from colorama import Fore, Style, init

# Inicializar colores
init(autoreset=True)

# ==========================================
# 1. BASES DE DATOS DE ASIGNATURAS
# ==========================================

malla_curricular = {
    'Agroclimatología': {'codigo': '2015880', 'creditos': 3, 'semestre': 4, 'tipo_asignatura': 'Disciplinar'},
    'Agroecosistemas y Sistemas de Producción': {'codigo': '2015881', 'creditos': 3, 'semestre': 8, 'tipo_asignatura': 'Disciplinar'},
    'Bioestadística fundamental': {'codigo': '1000012-B', 'creditos': 3, 'semestre': 3, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Biología Celular y Molecular Básica': {'codigo': '2015882', 'creditos': 3, 'semestre': 4, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Biología de plantas': {'codigo': '2015877', 'creditos': 3, 'semestre': 1, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Bioquímica básica': {'codigo': '1000042-B', 'creditos': 3, 'semestre': 3, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Botánica taxonómica': {'codigo': '2015878', 'creditos': 3, 'semestre': 2, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Ciclo  II: Ejecución de un proyecto productiv': {'codigo': '2015884', 'creditos': 3, 'semestre': 9, 'tipo_asignatura': 'Disciplinar'},
    'Ciclo i: formulación y evaluación de proyect': {'codigo': '2015883', 'creditos': 3, 'semestre': 8, 'tipo_asignatura': 'Disciplinar'},
    'Ciencia del suelo': {'codigo': '2015885', 'creditos': 3, 'semestre': 3, 'tipo_asignatura': 'Disciplinar'},
    'Cálculo Integral': {'codigo': '1000005-B', 'creditos': 4, 'semestre': 2, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Cálculo diferencial': {'codigo': '1000004-B', 'creditos': 4, 'semestre': 1, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Diseño de experimentos': {'codigo': '2015887', 'creditos': 3, 'semestre': 4, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Economía agraria': {'codigo': '2015888', 'creditos': 3, 'semestre': 6, 'tipo_asignatura': 'Disciplinar'},
    'Edafología': {'codigo': '2015889', 'creditos': 3, 'semestre': 4, 'tipo_asignatura': 'Disciplinar'},
    'Entomología': {'codigo': '2015890', 'creditos': 3, 'semestre': 6, 'tipo_asignatura': 'Disciplinar'},
    'Fisiología de la producción vegetal': {'codigo': '2015891', 'creditos': 3, 'semestre': 6, 'tipo_asignatura': 'Disciplinar'},
    'Fisiología vegetal básica': {'codigo': '2015892', 'creditos': 3, 'semestre': 5, 'tipo_asignatura': 'Disciplinar'},
    'Fitomejoramiento': {'codigo': '2015893', 'creditos': 3, 'semestre': 8, 'tipo_asignatura': 'Disciplinar'},
    'Fitopatología': {'codigo': '2015894', 'creditos': 3, 'semestre': 6, 'tipo_asignatura': 'Disciplinar'},
    'Fundamentos de ecología': {'codigo': '1000011-B', 'creditos': 3, 'semestre': 4, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Fundamentos de mecánica': {'codigo': '1000019-B', 'creditos': 3, 'semestre': 2, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Geomática básica': {'codigo': '2015896', 'creditos': 3, 'semestre': 3, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Gestión agroempresarial': {'codigo': '2015922', 'creditos': 3, 'semestre': 7, 'tipo_asignatura': 'Disciplinar'},
    'Génetica general': {'codigo': '2015895', 'creditos': 3, 'semestre': 5, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Inglés I- Semestral': {'codigo': '1000044-B', 'creditos': 2, 'semestre': 1, 'tipo_asignatura': 'Nivelación'},
    'Inglés II - Semestral': {'codigo': '1000045-B', 'creditos': 2, 'semestre': 2, 'tipo_asignatura': 'Nivelación'},
    'Inglés III - Semestral': {'codigo': '1000046-B', 'creditos': 2, 'semestre': 3, 'tipo_asignatura': 'Nivelación'},
    'Inglés IV- Semestral': {'codigo': '1000047-B', 'creditos': 2, 'semestre': 4, 'tipo_asignatura': 'Nivelación'},
    'Introducción a la ingeniería agronómica': {'codigo': '2015897', 'creditos': 2, 'semestre': 1, 'tipo_asignatura': 'Disciplinar'},
    'Laboratorio de bioquímica básica': {'codigo': '1000043-B', 'creditos': 2, 'semestre': 3, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Laboratorio de química básica': {'codigo': '2015782', 'creditos': 2, 'semestre': 2, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Lecto-Escritura': {'codigo': '1000002-B', 'creditos': 2, 'semestre': 1, 'tipo_asignatura': 'Nivelación'},
    'Manejo Integrado de Enfermedades': {'codigo': '2015899', 'creditos': 3, 'semestre': 7, 'tipo_asignatura': 'Disciplinar'},
    'Manejo de la fertilidad del suelo': {'codigo': '2015898', 'creditos': 3, 'semestre': 7, 'tipo_asignatura': 'Disciplinar'},
    'Manejo integrado de malezas': {'codigo': '2015900', 'creditos': 3, 'semestre': 7, 'tipo_asignatura': 'Disciplinar'},
    'Manejo integrado de plagas': {'codigo': '2015901', 'creditos': 3, 'semestre': 7, 'tipo_asignatura': 'Disciplinar'},
    'Matemáticas Básicas': {'codigo': '1000001-B', 'creditos': 3, 'semestre': 1, 'tipo_asignatura': 'Nivelación'},
    'Mecanización agrícola': {'codigo': '2015902', 'creditos': 3, 'semestre': 5, 'tipo_asignatura': 'Disciplinar'},
    'Microbiología': {'codigo': '2015903', 'creditos': 3, 'semestre': 4, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Práctica Profesional': {'codigo': '2015934', 'creditos': 6, 'semestre': 10, 'tipo_asignatura': 'Disciplinar'},
    'Química básica': {'codigo': '1000041-B', 'creditos': 3, 'semestre': 1, 'tipo_asignatura': 'Fund. Obligatoria'},
    'Reproducción y multiplicación': {'codigo': '2015907', 'creditos': 3, 'semestre': 6, 'tipo_asignatura': 'Disciplinar'},
    'Riegos y drenajes': {'codigo': '2015908', 'creditos': 3, 'semestre': 5, 'tipo_asignatura': 'Disciplinar'},
    'Sociología Rural': {'codigo': '2015909', 'creditos': 2, 'semestre': 5, 'tipo_asignatura': 'Disciplinar'},
    'Tecnología de la Poscosecha': {'codigo': '2015910', 'creditos': 3, 'semestre': 8, 'tipo_asignatura': 'Disciplinar'},
    'Trabajo de Grado': {'codigo': '2015291', 'creditos': 6, 'semestre': 10, 'tipo_asignatura': 'Disciplinar'}
}

optativas_produccion = {
    "Produccion de cultivos de clima calido": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de frutales": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de hortalizas": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de ornamentales": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Cultivos perennes industriales": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de papa": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
}

asignaturas_extra = {
    "Agroecología": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Agrobiodiversidad": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Bioprocesos Agroalimentarios": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Computación estadística": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Desarrollo Rural": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Emprendimiento e innovación en agronegocios": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Evolución y ecología de patógenos de plantas": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Fundamentos de Agroindustria": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Genética de Insectos de Interés económico": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Gestión ambiental agropecuaria": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Investigación de Mercados": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Nutrición Mineral de Plantas": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Producción de cannabis medicinal": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Suelos vivos": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"},
    "Sistemas Agroalimentarios Vinculo entre ambiente, sociedad y desarrollo": {"semestre": None, "Creditos": 3, "tipo_asignatura": "Libre Elección"}
}

asignaturas_posgrado = {
    "Agroclimatología y cambio climático": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Agua y nutrición mineral": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Biología de suelos": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Biología molecular": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Biología y ecología de malezas": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Clínica de plantas": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Decisiones de manejo fitosanitario: aproximación práctica": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Degradación química del suelo": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fertilizantes y fertilización": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Física de suelos": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fisiología avanzada en frutales": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fisiología de cultivos": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fisiología del desarrollo": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Fitopatología avanzada": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Genética avanzada": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Hongos y nemátodos fitopatógenos": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Métodos multivariados": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Pedología": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Recursos genéticos vegetales": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Taxonomía de insectos": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Desarrollo económico del territorio rural": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Desarrollo rural y territorios": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Economía de la empresa agraria y alimentaria": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Gestión contable financiera": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Gestión de agroproyectos": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Mercadeo agroalimentario y territorial": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Problemas agrarios colombianos": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Sociedad e instituciones rurales": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Geoestadística": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Geomática general": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Geoprocesamiento": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Percepción remota": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
    "Programación sig": {"semestre": None, "Creditos": 4, "tipo_asignatura": "Posgrado"},
}

# Combinar todas las asignaturas para búsqueda rápida en texto plano
TODAS_LAS_MATERIAS = list(malla_curricular.keys()) + \
                     list(optativas_produccion.keys()) + \
                     list(asignaturas_extra.keys()) + \
                     list(asignaturas_posgrado.keys())

# ==========================================
# 2. FUNCIONES DE LIMPIEZA Y METADATOS
# ==========================================

def clean_value(text):
    """Elimina comillas y comas extra: "Texto" -> Texto"""
    if not text: return ""
    return text.replace('"', '').replace(',', '').strip()

def extract_student_metadata(pdf):
    """
    Extrae metadatos. Usa Regex robustas que funcionan con o sin comillas.
    Ej: "Nombre:","Yeimi" O Nombre: Yeimi
    """
    text_page_1 = pdf.pages[0].extract_text()
    
    metadata = {
        'Nombre': 'Desconocido', 
        'Documento': '---', 
        'Periodo admisión': None, 
        'Periodo de inicio': None, 
        'Plan': None, 
        'Correo electrónico': None, 
        'Acceso': None, 
        'Subacceso': None
    }
    
    patterns = {
        'Nombre': r'Nombre"?:?\s*["\',]?\s*([^"\n\r]+)',
        'Documento': r'Documento"?:?\s*["\',]?\s*(\d+)',
        'Correo electrónico': r'Correo electrónico"?:?\s*["\',]?\s*([\w\.\-]+@[\w\.\-]+)',
        'Plan': r'Plan"?:?\s*["\',]?\s*([^"\n\r]+)',
        'Acceso': r'Acceso"?:?\s*["\',]?\s*([^"\n\r]+)',
        'Subacceso': r'Subacceso"?:?\s*["\',]?\s*([^"\n\r]+)',
        'Periodo admisión': r'Periodo admisión"?:?\s*["\',]?\s*([\d\-]+)',
        'Periodo de inicio': r'Periodo de inicio"?:?\s*["\',]?\s*([\d\-]+)',
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text_page_1)
        if match:
            val = match.group(1).strip()
            val = clean_value(val) # Limpieza extra de comillas
            if key == 'Plan':
                val = val.replace("Ingeniería Agronómica (2505)", "").strip()
            metadata[key] = val
            
    return metadata

def should_stop_processing(text):
    forbidden = ["Historial Bloqueos y Desbloqueos", "Distinciones y estímulos académicos", "Registro de novedades en la historia académica"]
    return any(f in text for f in forbidden)

# ==========================================
# 3. MOTOR DE EXTRACCIÓN (HÍBRIDO)
# ==========================================

def scan_text_lines_fallback(page_text, semester_label):
    """
    PLAN B: Escanea línea por línea buscando nombres de materias conocidos.
    Útil cuando el PDF no tiene tablas detectables o es formato CSV.
    """
    found_records = []
    lines = page_text.split('\n')
    
    for line in lines:
        line_clean = line.replace('"', '').strip() # Quitar comillas CSV
        
        # Buscar coincidencias con la base de datos de materias
        # Ordenamos por longitud inversa para encontrar "Inglés III" antes que "Inglés"
        for materia in sorted(TODAS_LAS_MATERIAS, key=len, reverse=True):
            if materia in line_clean:
                # Extraer números de la misma línea (posibles notas/créditos)
                # Ej: "Matemáticas Básicas","3","4.5"
                numeros = re.findall(r'\b\d+(?:[\.,]\d+)?\b', line_clean.replace(materia, ''))
                
                creditos = ""
                nota = ""
                veces = "1"
                
                # Intentar adivinar cuál es crédito y cuál es nota
                if numeros:
                    for n in numeros:
                        val = float(n.replace(',', '.'))
                        # Crédito suele ser entero pequeño (1-6)
                        if not creditos and val.is_integer() and 1 <= val <= 10:
                            creditos = int(val)
                        # Nota suele ser 0.0 a 5.0
                        elif not nota and 0 <= val <= 5.0:
                            nota = str(val)
                
                anulada = "Si" if "Anulada" in line_clean else "No"
                
                # Tipología basada en texto encontrado
                tipologia = "Desconocida/PDF"
                if "Fund." in line_clean or "Fundamentación" in line_clean: tipologia = "Fundamentación"
                elif "Disc" in line_clean or "Disciplinar" in line_clean: tipologia = "Disciplinar"
                elif "Libre" in line_clean: tipologia = "Libre Elección"
                elif "Nivel" in line_clean: tipologia = "Nivelación"

                found_records.append({
                    'Asignatura': materia,
                    'Codigo Asignatura': '',
                    'Tipología': tipologia,
                    'Créditos': creditos,
                    'Calificación': nota,
                    'Anulada': anulada,
                    'N. Veces': veces,
                    'Semestre Historia': semester_label
                })
                break # Materia encontrada en esta línea, siguiente línea
                
    return found_records

def process_pdf(file_path):
    student_records = []
    try:
        with pdfplumber.open(file_path) as pdf:
            metadata = extract_student_metadata(pdf)
            current_semester_label = "Desconocido"
            stop_scan = False

            for page in pdf.pages:
                if stop_scan: break
                page_text = page.extract_text() or ""
                
                if should_stop_processing(page_text): stop_scan = True
                
                # 1. Detectar Periodos (Semestres)
                semester_regex = r'(?:PRIMER|SEGUNDO|INTER)\s+PERIODO.*?(\d{4}-\d{1,2}S)'
                sem_matches = list(re.finditer(semester_regex, page_text))
                
                # 2. Dividir texto por bloques de semestre
                text_blocks = []
                last_idx = 0
                
                if not sem_matches:
                    text_blocks.append((current_semester_label, page_text))
                else:
                    for match in sem_matches:
                        # Bloque antes del nuevo semestre pertenece al anterior
                        block = page_text[last_idx:match.start()]
                        if block.strip():
                            text_blocks.append((current_semester_label, block))
                        
                        # Actualizar etiqueta actual
                        full_txt = match.group(0)
                        code_match = re.search(r'(\d{4}-\d{1,2}S)', full_txt)
                        current_semester_label = code_match.group(1) if code_match else full_txt
                        last_idx = match.end()
                    
                    # Bloque final
                    text_blocks.append((current_semester_label, page_text[last_idx:]))
                
                # 3. Procesar bloques con Fallback (Scanner)
                # Usamos directamente el Scanner porque demostró ser más seguro para tus PDFs sin bordes
                for sem_label, block_text in text_blocks:
                    if should_stop_processing(block_text): 
                        stop_scan = True
                        break
                        
                    found = scan_text_lines_fallback(block_text, sem_label)
                    for f in found:
                        full_record = metadata.copy()
                        full_record.update(f)
                        student_records.append(full_record)

    except Exception as e:
        return []
        
    return student_records

# ==========================================
# 4. EJECUCIÓN
# ==========================================

def main():
    folder_path = '/home/jugalindog/Pasantia/Curricular/Curricular/Historial_Academica/activos'
    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
    
    print(f"{Fore.CYAN}=== EXTRACTOR TODOTERRENO V3.1 (TEXTO PLANO) ==={Style.RESET_ALL}")
    print(f"Ruta: {folder_path}")
    print(f"Archivos encontrados: {len(pdf_files)}")

    all_data = []
    total_subjects = 0

    with tqdm(total=len(pdf_files), unit="pdf", ncols=100, colour='green') as barra:
        for file_path in pdf_files:
            file_name = os.path.basename(file_path)
            barra.set_description(f"Leyendo: {file_name[:12]}..") 
            
            records = process_pdf(file_path)
            
            if records:
                nom = records[0]['Nombre'][:15]
                doc = records[0]['Documento']
            else:
                nom = "Sin datos"
                doc = "---"

            all_data.extend(records)
            total_subjects += len(records)
            
            barra.set_postfix({"Est": nom, "Doc": doc, "Tot": total_subjects})
            barra.update(1)

    if all_data:
        df = pd.DataFrame(all_data)
        # Ordenar columnas
        cols = ['Nombre', 'Documento', 'Asignatura', 'Créditos', 'Calificación', 'Semestre Historia', 'Tipología', 'Plan']
        cols = [c for c in cols if c in df.columns] + [c for c in df.columns if c not in cols]
        df = df[cols]
        
        out = "Reporte_Todoterreno.xlsx"
        df.to_excel(out, index=False)
        print(f"\n{Fore.CYAN}¡Éxito! Archivo generado: {os.path.abspath(out)}{Style.RESET_ALL}")
    else:
        print(f"\n{Fore.RED}No se extrajeron datos.{Style.RESET_ALL}")

if __name__ == "__main__":
    main()