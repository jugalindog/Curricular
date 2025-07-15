import re
import fitz  # PyMuPDF
import pandas as pd
import os

# --- Textos basura que se deben eliminar del PDF ---
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



CARPETA_PDFS = "/home/jugalindog/Pasantia/Curricular/Curricular/Historial_Academica"
promedios_por_periodo = []
promedios = []
for archivo in os.listdir(CARPETA_PDFS):
    if not archivo.endswith(".pdf"):
        continue

    ruta_pdf = os.path.join(CARPETA_PDFS, archivo)
    try:
        doc = fitz.open(ruta_pdf)
        texto = "\n".join([page.get_text() for page in doc])
        doc.close()

        # Limpiar basura del texto
        for b in basura.values():
            texto = texto.replace(b, '')

        texto = re.sub(r"Informe generado.*\d{2}:\d{2}", '', texto)
        texto = re.sub(r'Página\xa0\d+\xa0de\xa0\d+', '', texto)
        texto = re.sub(r'\n?[A-ZÁÉÍÓÚÑ][^\n]+\s+-\s+\d{7,10}', '', texto)
        texto = re.sub(r'\b\w{3,}\s+el\s+\w+\s+\d{1,2}\s+de\s+\w+\s+de\s+\d{4}\s+\d{2}:\d{2}', '', texto)

    except Exception as e:
        print(f"❌ Error con {archivo}: {e}")
        continue

    # Extraer nombre y documento
    nombre_match = re.search(r'Nombre:\s*(.+)', texto)
    documento_match = re.search(r'Documento:\s*(\d+)', texto)

    if not nombre_match or not documento_match:
        continue
    nombre = nombre_match.group(1).strip()
    documento = documento_match.group(1).strip()

    promedios = [] # Lista para almacenar los promedios extraídos
    # Extraer el bloque entre "Promedios" y "Resumen de créditos"
    bloque_prom = re.search(r'Promedios\s+(.*?)\s+Resumen de créditos', texto, re.DOTALL)
    if bloque_prom:
        texto_prom = bloque_prom.group(1)
        lineas_prom = texto_prom.replace("\n", " ").split()
        
        datos_limpios = []
        for token in lineas_prom:
            token = token.strip()
            if token and token != ',':
                datos_limpios.append(token)

        i = 0
        
        while i < len(datos_limpios) - 4:
            if re.fullmatch(r'\d{4}-[12]S', datos_limpios[i]):
                print("📊 Promedios extraídos:", len(promedios))

                periodo = datos_limpios[i]
                promedio = float(datos_limpios[i+1].replace(",", "."))
                creditos = int(datos_limpios[i+2])
                tipo = datos_limpios[i+4]
                promedios.append({
                    'nombre': nombre,
                    'documento': documento,
                    'periodo': periodo,
                    'promedio': promedio,
                    'creditos': creditos,
                    'tipo': tipo
                })
                i += 5
            else:
                i += 1
                promedios_por_periodo.extend(promedios)

# --- Exportar a Excel ---
df_prom = pd.DataFrame(promedios_por_periodo)
df_prom.to_excel("Promedios_DB.xlsx", index=False)
print("✅ Archivo Promedios_desde_bloques.xlsx generado correctamente.")