import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
import io
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador PEP", page_icon="📚", layout="wide")

# Estilo para etiquetas Opcional/Obligatorio
ST_OPCIONAL = '<span style="color: gray; font-size: 0.8em;">(Opcional)</span>'
ST_OBLIGATORIO = '<span style="color: gray; font-size: 0.8em;">(Obligatorio)</span>'

st.title("📚 Generador PEP - Módulo 1: Información del Programa")

# --- LÓGICA DE API KEY ---
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    with st.sidebar:
        api_key = st.text_input("Google API Key", type="password")

# --- BOTÓN DE DATOS DE EJEMPLO ---
if st.button("🧪 Llenar con datos de ejemplo"):
    st.session_state.ejemplo = {
        "denom": "Ingeniería de Software",
        "titulo": "Ingeniero de Software",
        "nivel": "Profesional universitario",
        "area": "Ingeniería, Arquitectura, Urbanismo y afines",
        "modalidad": "Presencial y Virtual",
        "acuerdo": "Acuerdo 045 de 2010",
        "instancia": "Consejo Superior Universitario",
        "reg1": "Res. 12345 de 2011",
        "reg2": "Res. 67890 de 2018",
        "acred1": "Res. CNA 001 de 2020",
        "creditos": "160",
        "periodo": "Semestral",
        "lugar": "Bogotá D.C. y Medellín",
        "snies": "102938",
        "motivo": "Atender la creciente demanda de transformación digital en el sector productivo nacional.",
        "p1_nom": "Plan Innova v1", "p1_fecha": "2010",
        "p2_nom": "Plan Ajuste v2", "p2_fecha": "2015",
        "p3_nom": "Plan v3", "p3_fecha": "2022"
    }
    st.rerun()

# --- FORMULARIO DE ENTRADA ---
with st.form("pep_form"):
    # Cargar valores si existe ejemplo
    ej = st.session_state.get("ejemplo", {})

    col1, col2 = st.columns(2)
    with col1:
        denom = st.text_input(f"Denominación del programa {ST_OBLIGATORIO}", value=ej.get("denom", ""), help="Nombre oficial", label_visibility="visible")
        titulo = st.text_input(f"Título otorgado {ST_OBLIGATORIO}", value=ej.get("titulo", ""))
        nivel = st.selectbox(f"Nivel de formación {ST_OBLIGATORIO}", ["Técnico", "Tecnológico", "Profesional universitario", "Especialización", "Maestría", "Doctorado"], index=2)
        area = st.text_input(f"Área de formación {ST_OBLIGATORIO}", value=ej.get("area", ""))
    
    with col2:
        modalidad = st.selectbox(f"Modalidad de oferta {ST_OBLIGATORIO}", ["Presencial", "Virtual", "A Distancia", "Dual", "Presencial y Virtual", "Presencial y a Distancia", "Presencial y Dual"])
        acuerdo = st.text_input(f"Acuerdo de creación (Norma interna) {ST_OBLIGATORIO}", value=ej.get("acuerdo", ""))
        instancia = st.text_input(f"Instancia interna que aprueba {ST_OBLIGATORIO}", value=ej.get("instancia", ""))
        snies = st.text_input(f"Código SNIES {ST_OBLIGATORIO}", value=ej.get("snies", ""))

    st.markdown("---")
    col3, col4 = st.columns(2)
    with col3:
        reg1 = st.text_input(f"Resolución Registro calificado 1 {ST_OBLIGATORIO}", value=ej.get("reg1", ""), placeholder="Número y año")
        reg2 = st.text_input(f"Registro calificado 2 {ST_OPCIONAL}", value=ej.get("reg2", ""))
        acred1 = st.text_input(f"Resolución Acreditación 1 {ST_OPCIONAL}", value=ej.get("acred1", ""))
        acred2 = st.text_input(f"Resolución Acreditación 2 {ST_OPCIONAL}", value="")

    with col4:
        creditos = st.text_input(f"Créditos académicos {ST_OBLIGATORIO}", value=ej.get("creditos", ""))
        periodicidad = st.selectbox(f"Periodicidad de admisión {ST_OBLIGATORIO}", ["Semestral", "Anual"])
        lugares = st.text_input(f"Lugares de desarrollo {ST_OBLIGATORIO}", value=ej.get("lugar", ""))

    motivo = st.text_area(f"Motivo de creación del Programa {ST_OBLIGATORIO}", value=ej.get("motivo", ""), height=150)

    st.subheader("Planes de Estudio")
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        p1_nom = st.text_input(f"Nombre Plan v1 {ST_OBLIGATORIO}", value=ej.get("p1_nom", ""))
        p1_fec = st.text_input(f"Fecha Plan v1 {ST_OBLIGATORIO}", value=ej.get("p1_fecha", ""))
    with p_col2:
        p2_nom = st.text_input(f"Nombre Plan v2 {ST_OPCIONAL}", value=ej.get("p2_nom", ""))
        p2_fec = st.text_input(f"Fecha Plan v2 {ST_OPCIONAL}", value=ej.get("p2_fecha", ""))
    with p_col3:
        p3_nom = st.text_input(f"Nombre Plan v3 {ST_OPCIONAL}", value=ej.get("p3_nom", ""))
        p3_fec = st.text_input(f"Fecha Plan v3 {ST_OPCIONAL}", value=ej.get("p3_fecha", ""))

    st.subheader("Reconocimientos (Opcional)")
    recon_data = st.data_editor(
        [{"Año": "", "Nombre": "", "Ganador": "", "Cargo": "Estudiante"}],
        num_rows="dynamic",
        column_config={
            "Cargo": st.column_config.SelectboxColumn(options=["Docente", "Líder", "Decano", "Estudiante"])
        }
    )

    generar = st.form_submit_button("🚀 Generar Módulo 1")

# --- LÓGICA DE GENERACIÓN ---
if generar:
    doc = Document()
    
    # 1.1 Historia del Programa (Lógica de Texto)
    doc.add_heading("1.1. Historia del Programa", level=1)
    
    # Párrafo Base
    p1 = doc.add_paragraph(
        f"El Programa de {denom} fue creado mediante el {acuerdo} del {instancia} "
        f"y aprobada mediante la resolución de Registro Calificado {reg1} del Ministerio de Educación Nacional "
        f"con código SNIES {snies}."
    )

    # Condicional Acreditación
    if acred1:
        p_acred = doc.add_paragraph(
            f"El Programa desarrolla de manera permanente procesos de autoevaluación y autorregulación, "
            f"orientados al aseguramiento de la calidad académica. Como resultado de estos procesos, "
            f"el Programa obtuvo la Acreditación en Alta Calidad mediante {acred1}, como reconocimiento "
            f"a la solidez de sus condiciones académicas y administrativas."
        )

    # Condicional Planes de Estudio
    planes = [f for f in [p1_fec, p2_fec, p3_fec] if f]
    acuerdos_plan = [n for n in [p1_nom, p2_nom, p3_nom] if n]
    if len(planes) > 1:
        p_evol = doc.add_paragraph(
            f"El plan de estudios del Programa de {denom} ha sido objeto de procesos periódicos de evaluación. "
            f"Como resultado, se han realizado modificaciones curriculares en los años {', '.join(planes)}, "
            f"aprobadas mediante {', '.join(acuerdos_plan)}."
        )

    # Reconocimientos
    if any(r["Nombre"] for r in recon_data):
        doc.add_paragraph(f"El Programa de {denom} ha alcanzado importantes logros académicos:")
        for r in recon_data:
            if r["Nombre"]:
                doc.add_paragraph(f"• {r['Año']}: {r['Nombre']} otorgado a {r['Ganador']} ({r['Cargo']}).", style='List Bullet')

    # Línea de tiempo (Hitos)
    doc.add_heading("Línea de tiempo de los principales hitos del Programa", level=2)
    doc.add_paragraph(f"{p1_fec}: Creación del Programa")
    doc.add_paragraph(f"{p1_fec}: Obtención del Registro Calificado")
    if p2_fec: doc.add_paragraph(f"{p2_fec}: Actualización del plan de estudios")
    if acred1: doc.add_paragraph("20XX: Acreditación de Alta Calidad") # Podrías extraer el año de acred1

    # 1.2 Generalidades (Directo)
    doc.add_page_break()
    doc.add_heading("1.2 Generalidades del Programa", level=1)
    generalidades = [
        ("Denominación", denom), ("Título", titulo), ("Nivel", nivel), 
        ("Modalidad", modalidad), ("SNIES", snies), ("Créditos", creditos)
    ]
    for k, v in generalidades:
        p = doc.add_paragraph()
        p.add_run(f"{k}: ").bold = True
        p.add_run(v)

    # Guardar
    bio = io.BytesIO()
    doc.save(bio)
    st.success("¡Documento generado con éxito!")
    st.download_button("📥 Descargar Word", bio.getvalue(), f"PEP_Modulo1_{denom}.docx")
