import re
import fitz        # PyMuPDF
import pandas as pd
import os

# --- Texto fijo que siempre queremos quitar ---------------------------------
BASURA = {
    0: 'Abreviaturas utilizadas: HAB=Habilitación, VAL=Validación por Pérdida, SUF=Validación por Suficiencia, '
       'HAP=Horas de Actividad Presencial, HAI=Horas de Actividad',
    1: 'Independiente, THS=Total Horas Semanales, HOM=Homologada o Convalidada.',
    2: 'SI*: Cancelación por decisión de la universidad soportada en acuerdos, resoluciones y actos académicos',
    3: 'Este es un documento de uso interno de la Universidad Nacional de Colombia. '
       'No constituye, ni reemplaza el certificado oficial de notas.',
    4: 'Informe generado por el usuario:',
    5: 'Reporte de Historia Académica',
    6: 'Sistema de Información Académica',
    7: 'Dirección Nacional de Información Académica',
    8: 'Registro y Matrícula'
}
patron_basura = re.compile("|".join(re.escape(v) for v in BASURA.values()))

# --- Carpeta con PDFs --------------------------------------------------------
CARPETA_PDFS = "/home/jugalindog/Pasantia/Curricular/Historial_Academica"

datos = []   # ← se mantiene durante TODO el recorrido

# ----------------------------------------------------------------------------- 
for archivo in os.listdir(CARPETA_PDFS):
    if not archivo.lower().endswith(".pdf"):
        continue

    ruta_pdf = os.path.join(CARPETA_PDFS, archivo)
    print(f"Procesando: {archivo}")

    try:
        # 1. Leer texto del PDF
        with fitz.open(ruta_pdf) as doc:
            texto = "\n".join(page.get_text() for page in doc)

        # 2. Limpiezas
        texto = re.sub(
            r"Informe generado por el usuario:\s+\S+\s+el\s+\w+\s+\d{1,2}\s+de\s+\w+\s+de\s+\d{4}\s+\d{2}:\d{2}",
            "", texto
        )
        texto = re.sub(r"Página\xa0\d+\xa0de\xa0\d+", "", texto)          # paginación
        texto = patron_basura.sub("", texto)                              # metadatos

        # 3. Nombre y documento
        nm = re.search(r"Nombre:\s*(.+)", texto)
        dm = re.search(r"Documento:\s*(\d+)", texto)
        if not (nm and dm):
            print("   → No encontró nombre/documento, se omite.")
            continue
        nombre, documento = nm.group(1).strip(), dm.group(1).strip()

        # 4. Separar por semestres
        bloques = re.split(r"(?:PRIMER|SEGUNDO)\s+PERIODO\s+(\d{4}-[12]S)", texto)

        for i in range(1, len(bloques), 2):
            semestre = bloques[i]
            lineas   = [l.strip() for l in bloques[i + 1].splitlines() if l.strip()]

            j = 0
            while j < len(lineas):
                linea = lineas[j]

                m = re.search(r"(.+?)\s*\((\d{6,7}(?:-B)?)\)", linea)  # nombre y código
                if m:
                    nombre_asig = m.group(1).strip()
                    codigo      = m.group(2)
                    tipo_asig, nota, estado, anulada = "", "", "Reprobada", "NO"

                    j += 1
                    while j < len(lineas):
                        l = lineas[j]

                        if re.search(r"(.+?)\s*\((\d{6,7}(?:-B)?)\)", l):
                            j -= 1
                            break

                        if re.search(r"(Aprobada|Reprobada|SI\*)", l):
                            nm_nota = re.search(r"([\d,\.]+)", l)
                            if nm_nota:
                                nota = nm_nota.group(1).replace(",", ".")
                            estado = "Aprobada" if "Aprobada" in l else "Reprobada"

                        if "Anulada" in l or "SI" in l:
                            anulada = "SI"

                        elif any(t in l for t in [
                             "Obligatoria", "Optativa", "Libre Elección", "Nivelación"
                        ]):
                            tipo_asig = l
                        j += 1

                    datos.append({
                        "nombre": nombre,
                        "documento": documento,
                        "codigo_asignatura": codigo,
                        "asignatura": nombre_asig,
                        "tipo_asignatura": tipo_asig,
                        "nota": float(nota) if nota.replace(".", "", 1).isdigit() else 0.0,
                        "estado": estado,
                        "anulada": anulada,
                        "semestre_inicio": "2018-2S",
                        "semestre_asignatura": semestre
                    })
                j += 1

    except Exception as e:
        print(f"   ⚠️  Error con {archivo}: {e}")

# 5. Exportar todo lo acumulado
df = pd.DataFrame(datos)
df.to_excel("historia_academica_robusta_final2.xlsx", index=False)
print("✅ Archivo listo con", len(df), "registros: historia_academica_robusta_final.xlsx")
