import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt, Inches
import io
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador PEP", page_icon="📚", layout="wide")

# Estilo para los textos en gris (Opcional/Obligatorio)
st.markdown("""
    <style> .gray-text { color: #808080; font-size: 0.8em; } </style>
    """, unsafe_allow_html=True)

st.title("📚 Generador PEP - Sección 1: Información del Programa")

# --- LÓGICA DE API KEY ---
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = st.sidebar.text_input("Ingresa tu Google API Key", type="password")

# --- BOTÓN DE EJEMPLO ---
if st.button("✨ Llenar con datos de ejemplo"):
    st.session_state["denominacion"] = "Ingeniería de Software"
    st.session_state["titulo"] = "Ingeniero de Software"
    st.session_state["nivel"] = "Profesional universitario"
    st.session_state["area"] = "Ingeniería, Arquitectura y Urbanismo"
    st.session_state["modalidad"] = "Presencial"
    st.session_state["acuerdo"] = "Acuerdo 012 de 2010"
    st.session_state["instancia"] = "Consejo Superior Universitario"
    st.session_state["registro1"] = "Res. 1234 de 2011"
    st.session_state["snies"] = "102544"
    st.session_state["creditos"] = "160"
    st.session_state["admision"] = "Semestral"
    st.session_state["lugares"] = "Bogotá D.C."
    st.session_state["motivo"] = "El programa fue creado para responder a la creciente demanda de desarrolladores en la región andina y fortalecer la industria 4.0."
    st.session_state["acreditacion1"] = "Res. 5678 de 2020"
    st.session_state["plan1_nom"] = "Plan 2010 v1"
    st.session_state["plan1_fec"] = "2010"

# --- FORMULARIO ---
with st.form("pep_form"):
    col1, col2 = st.columns(2)
    
    with col1:
        denominacion = st.text_input("Denominación del programa", key="denominacion", help="Obligatorio")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)
        
        titulo = st.text_input("Título otorgado", key="titulo")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)
        
        nivel = st.selectbox("Nivel de formación", ["Técnico", "Tecnológico", "Profesional universitario", "Especialización", "Maestría", "Doctorado"], key="nivel")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)
        
        area = st.text_input("Área de formación", key="area")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)
        
        modalidad = st.selectbox("Modalidad de oferta", ["Presencial", "Virtual", "A Distancia", "Dual", "Presencial y Virtual", "Presencial y a Distancia", "Presencial y Dual"], key="modalidad")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)

    with col2:
        acuerdo = st.text_input("Acuerdo de creación (Norma interna)", key="acuerdo")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)
        
        instancia = st.text_input("Instancia interna que aprueba el Programa", key="instancia")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)
        
        registro1 = st.text_input("Resolución Registro calificado 1 (Número y año)", key="registro1")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)
        
        snies = st.text_input("Código SNIES", key="snies")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)
        
        admision = st.selectbox("Periodicidad de admisión", ["Semestral", "Anual"], key="admision")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)

    st.divider()
    st.subheader("Información Adicional y Planes")
    
    col3, col4 = st.columns(2)
    with col3:
        registro2 = st.text_input("Registro calificado 2", key="registro2")
        st.markdown('<p class="gray-text">(Opcional)</p>', unsafe_allow_html=True)
        
        acreditacion1 = st.text_input("Resolución Acreditación en alta calidad 1", key="acreditacion1")
        st.markdown('<p class="gray-text">(Opcional)</p>', unsafe_allow_html=True)
        
        creditos = st.text_input("Créditos académicos", key="creditos")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)

    with col4:
        lugares = st.text_input("Lugares de desarrollo", key="lugares")
        st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)
        
        acreditacion2 = st.text_input("Resolución Acreditación en alta calidad 2", key="acreditacion2")
        st.markdown('<p class="gray-text">(Opcional)</p>', unsafe_allow_html=True)

    motivo = st.text_area("Motivo de creación del Programa", height=150, key="motivo")
    st.markdown('<p class="gray-text">(Obligatorio)</p>', unsafe_allow_html=True)

    st.subheader("Reconocimientos (Opcional)")
    reconocimientos = st.data_editor(
        [{"Año": "", "Nombre": "", "Ganador": "", "Cargo": "Docente"}],
        num_rows="dynamic",
        column_config={
            "Cargo": st.column_config.SelectboxColumn(options=["Docente", "Líder", "Decano", "Estudiante"])
        },
        key="tabla_recon"
    )

    st.subheader("Planes de Estudio")
    p1, p2, p3 = st.columns(3)
    with p1:
        plan1_nom = st.text_input("Nombre Plan v1", key="plan1_nom")
        plan1_fec = st.text_input("Fecha Plan v1", key="plan1_fec")
    with p2:
        plan2_nom = st.text_input("Nombre Plan v2 (Opcional)", key="plan2_nom")
        plan2_fec = st.text_input("Fecha Plan v2 (Opcional)", key="plan2_fec")
    with p3:
        plan3_nom = st.text_input("Nombre Plan v3 (Opcional)", key="plan3_nom")
        plan3_fec = st.text_input("Fecha Plan v3 (Opcional)", key="plan3_fec")

    generar = st.form_submit_button("🚀 Generar Informe Word")

# --- LÓGICA DE GENERACIÓN WORD ---
if generar:
    doc = Document()
    
    # 1.1 Historia del Programa
    doc.add_heading('1.1. Historia del Programa', level=1)
    
    # Párrafo Inicial
    p1 = doc.add_paragraph()
    p1.add_run(f"El Programa de {denominacion} fue creado mediante el {acuerdo} del {instancia} y aprobado mediante la {registro1} del Ministerio de Educación Nacional con Código SNIES {snies}.")

    # Párrafo Acreditación (Condicional)
    if acreditacion1:
        p_acre = doc.add_paragraph()
        p_acre.add_run(f"El Programa desarrolla de manera permanente procesos de autoevaluación y autorregulación, orientados al aseguramiento de la calidad académica. Como resultado de estos procesos, y tras demostrar el cumplimiento integral de los factores, características y lineamientos de alta calidad establecidos por el Consejo Nacional de Acreditación (CNA), el Programa obtuvo la Acreditación en Alta Calidad mediante {acreditacion1}, como reconocimiento a la solidez de sus condiciones académicas, administrativas y de impacto social.")

    # Párrafo Modificaciones Plan Estudios
    planes = [plan1_fec, plan2_fec, plan3_fec]
    resoluciones = [plan1_nom, plan2_nom, plan3_nom]
    planes_llenos = [p for p in planes if p]
    resol_llenas = [r for r in resoluciones if r]
    
    if len(planes_llenos) > 1:
        p_mod = doc.add_paragraph()
        p_mod.add_run(f"El plan de estudios del Programa de {denominacion} ha sido objeto de procesos periódicos de evaluación, con el fin de asegurar su pertinencia académica y su alineación con los avances tecnológicos y las demandas del entorno. Como resultado, se han realizado modificaciones curriculares en los años {', '.join(planes_llenos)}, aprobadas mediante Acuerdo(s) del Consejo Académico Nos. {', '.join(resol_llenas)}, respectivamente.")

    # Párrafo Reconocimientos
    recon_validos = [r for r in reconocimientos if r['Nombre']]
    if recon_validos:
        p_rec = doc.add_paragraph()
        p_rec.add_run(f"El Programa de {denominacion} ha alcanzado importantes logros académicos e institucionales que evidencian su calidad y compromiso con la excelencia. Entre ellos se destacan: ")
        for r in recon_validos:
            p_rec.add_run(f"{r['Nombre']} otorgado a {r['Ganador']} ({r['Cargo']}) en el año {r['Año']}; ").italic = True

    # Línea de Tiempo
    doc.add_heading('Línea de tiempo de los principales hitos del Programa', level=2)
    hitos = [
        f"{plan1_fec}: Creación del Programa",
        f"{registro1.split()[-1] if registro1 else '20XX'}: Obtención del Registro Calificado"
    ]
    if plan2_fec: hitos.append(f"{plan2_fec}: Primera actualización del plan de estudios")
    if registro2: hitos.append(f"{registro2.split()[-1]}: Renovación del Registro Calificado")
    if recon_validos: hitos.append(f"{recon_validos[0]['Año']}: Reconocimientos académicos")
    
    for h in hitos:
        doc.add_paragraph(h, style='List Bullet')

    # Sección 1.2 Generalidades (Directo)
    doc.add_heading('1.2. Generalidades del Programa', level=1)
    tabla = doc.add_table(rows=1, cols=2)
    for k, v in {"Denominación": denominacion, "Título": titulo, "Nivel": nivel, "SNIES": snies, "Modalidad": modalidad}.items():
        row = tabla.add_row().cells
        row[0].text = k
        row[1].text = str(v)

    # Descarga
    target = io.BytesIO()
    doc.save(target)
    st.download_button("📥 Descargar Word", target.getvalue(), f"PEP_{denominacion}.docx")

