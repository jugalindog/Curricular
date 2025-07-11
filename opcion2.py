import re
import pandas as pd
from PyPDF2 import PdfReader

# Ruta al archivo PDF
pdf_path = "/home/jugalindog/Pasantia/Curricular/Historial_Academica/RE_EST_HCA_REPORTE.pdf"

# Cargar el PDF
reader = PdfReader(pdf_path)

# Extraer texto por líneas
raw_lines = []
for page in reader.pages:
    lines = page.extract_text().split("\n")
    raw_lines.extend(lines)

# Datos del estudiante
nombre = "Juan Pablo Galindo Gómez"
documento = "1136887478"

# Variables de control
asignaturas_data = []
periodo_actual = None

# Analizar línea por línea
for line in raw_lines:
    match_periodo = re.match(r"(PRIMER|SEGUNDO)\s+PERIODO\s+\d{4}-\dS", line)
    if match_periodo:
        periodo_actual = match_periodo.group(0)
        continue

    # Intentar extraer una fila de asignatura
    match_asignatura = re.match(
        r"(.+?)\s+(\d{1,2})\s+[0-9\s]{1,4}(.{5,35}?)\s+([\d,]+)\s+(Aprobada|Reprobada|SI\*|SI)?\s+(NO|SI)?\s+(\d)",
        line
    )

    if match_asignatura:
        asignatura = match_asignatura.group(1).strip()
        creditos = int(match_asignatura.group(2))
        tipologia = match_asignatura.group(3).strip()
        nota = float(match_asignatura.group(4).replace(",", "."))
        estado = match_asignatura.group(5) if match_asignatura.group(5) else "Desconocido"
        anulada = match_asignatura.group(6) if match_asignatura.group(6) else "NO"
        veces = int(match_asignatura.group(7))

        asignaturas_data.append({
            "Nombre": nombre,
            "Documento": documento,
            "Periodo": periodo_actual,
            "Asignatura": asignatura,
            "Créditos": creditos,
            "Tipología": tipologia,
            "Nota": nota,
            "Estado": estado,
            "Anulada": anulada,
            "No. Veces": veces
        })

# Crear DataFrame y exportar
df = pd.DataFrame(asignaturas_data)
df.to_excel("historial_asignaturas.xlsx", index=False)

print("Extracción completada. Archivo guardado como 'historial_asignaturas.xlsx'.")
