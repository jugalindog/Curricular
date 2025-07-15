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
    "Fitomejoramiento": {"semestre": 8, "creditos": 3, "tipo_asignatura": "Disciplinar"},
    "Trabajo de Grado": {"semestre": 10, "creditos": 6, "tipo_asignatura": "Disciplinar"},
    "Práctica Profesional": {"semestre": 10, "creditos": 6, "tipo_asignatura": "Disciplinar"},
    "Producción de frutales": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Producción de cultivos de clima calido": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    "Cultivos perennes industriales": {"semestre": 9, "creditos": 3, "tipo_asignatura": "Optativa de Producción"},
    # Agrega más si es necesario...
}

optativas_produccion = [
    "Producción de cultivos de clima calido",
    "Producción de frutales",
    "Producción de hortalizas",
    "Producción de ornamentales",
    "Cultivos perennes industriales",
    "Producción de papa"
]

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

                if nombre_asig in optativas_produccion:
                    tipo_asig = 'Optativa de producción (C)'

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
