# --- Importación de librerías ---
import re
import fitz  # PyMuPDF
import pandas as pd
import os

# Diccionario con texto basura típico de los PDFs que debe ser eliminado
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

# Diccionario de malla curricular para clasificación por semestre y créditos
malla_curricular = {
    "Matemáticas Básicas": {"semestre": 1, "creditos": 3},
    "Biología de plantas": {"semestre": 1, "creditos": 3},
    "Lecto-Escritura":     {"semestre": 1, "creditos": 2},
    "Cálculo diferencial": {"semestre": 1, "creditos": 4},
    "Química básica":      {"semestre": 1, "creditos": 3},
    "Cálculo Integral":    {"semestre": 2, "creditos": 4},
    "Fundamentos de mecánica": {"semestre": 2, "creditos": 3},
    "Botánica taxonómica":     {"semestre": 2, "creditos": 3},
    "Laboratorio de química básica": {"semestre": 2, "creditos": 1},
    "Ciencia del suelo":       {"semestre": 3, "creditos": 3},
    "Laboratorio bioquímica básica": {"semestre": 3, "creditos": 1},
    "Bioquímica básica":       {"semestre": 3, "creditos": 3},
    "Bioestadística fundamental": {"semestre": 3, "creditos": 3},
    "Geomática básica":       {"semestre": 3, "creditos": 2},
    "Agroclimatología":       {"semestre": 4, "creditos": 2},
    "Edafología":             {"semestre": 4, "creditos": 3},
    "Fundamentos de ecología": {"semestre": 4, "creditos": 2},
    "Microbiología":          {"semestre": 4, "creditos": 3},
    "Biología Celular y Molecular Básica": {"semestre": 4, "creditos": 3},
    "Diseño de experimentos": {"semestre": 4, "creditos": 3},
    "Sociología Rural":       {"semestre": 5, "creditos": 2},
    "Riegos y drenajes":      {"semestre": 5, "creditos": 3},
    "Mecanización agrícola":  {"semestre": 5, "creditos": 3},
    "Génetica general":       {"semestre": 5, "creditos": 3},
    "Fisiología vegetal básica": {"semestre": 5, "creditos": 3},
    "Economía agraria":      {"semestre": 6, "creditos": 3},
    "Entomología":           {"semestre": 6, "creditos": 3},
    "Fitopatología":         {"semestre": 6, "creditos": 3},
    "Fisiología de la producción vegetal": {"semestre": 6, "creditos": 3},
    "Reproducción y multiplicación": {"semestre": 6, "creditos": 3},
    "Gestión agroempresarial":       {"semestre": 7, "creditos": 3},
    "Manejo de la fertilidad del suelo": {"semestre": 7, "creditos": 3},
    "Manejo integrado de plagas":    {"semestre": 7, "creditos": 3},
    "Manejo Integrado de Enfermedades": {"semestre": 7, "creditos": 3},
    "Manejo integrado de malezas":   {"semestre": 7, "creditos": 3},
    "Ciclo i: formulación y evaluación de proyectos": {"semestre": 8, "creditos": 3},
    "Fitomejoramiento":              {"semestre": 8, "creditos": 3},
    "Agroecosistemas y Sistemas de Producción": {"semestre": 8, "creditos": 3},
    "Tecnología de la Poscosecha":   {"semestre": 8, "creditos": 3},
    "Ciclo  II: Ejecución de un proyecto productivo": {"semestre": 9, "creditos": 3},
    "Produccion de cultivos de clima calido": {"semestre": 9, "creditos": 3},
    "Producción de frutales":        {"semestre": 9, "creditos": 3},
    "Producción de hortalizas":      {"semestre": 9, "creditos": 3},
    "Producción de ornamentales":    {"semestre": 9, "creditos": 3},
    "Cultivos perennes industriales": {"semestre": 9, "creditos": 3},
    "Producción de papa":             {"semestre": 9, "creditos": 3},
    "Práctica Profesional":           {"semestre": 10, "creditos": 4},
    "Trabajo de Grado":               {"semestre": 10, "creditos": 6}
}

# Lista de asignaturas consideradas como optativas de producción
optativas_produccion = [
    "Produccion de cultivos de clima calido",
    "Producción de frutales",
    "Produccion de hortalizas",
    "Producción de ornamentales",
    "Cultivos perennes industriales",
    "Producción de papa"
]

# dentro del bucle donde se clasifica la asignatura, justo antes de datos.append:
info_malla = malla_curricular.get(nombre_asig)
if info_malla:
    semestre_malla = info_malla["semestre"]
    creditos = info_malla["creditos"]
else:
    semestre_malla = 'Libre Elección'
    creditos = ''
# Y en el bloque datos.append:
    datos.append({
        'nombre': nombre,
        'documento': documento,
        'codigo_asignatura': codigo,
        'asignatura': nombre_asig,
        'tipo_asignatura': tipo_asig,
        'semestre_malla': semestre_malla,
        'creditos': creditos,
        'nota': float(nota) if nota.replace('.', '', 1).isdigit() else 0.0,
        'estado': estado,
        'anulada': anulada,
        'semestre_inicio': '2018-2S',
        'semestre_asignatura': semestre
    })
