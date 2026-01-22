import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
import io
import time

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Generador PEP", layout="wide")

# Función para el botón de ejemplo
def cargar_ejemplo():
    st.session_state.denominacion = "Ingeniería de Inteligencia Artificial"
    st.session_state.titulo = "Ingeniero de Inteligencia Artificial"
    st.session_state.nivel = "Profesional universitario"
    st.session_state.area = "Ingeniería, Industria y Construcción"
    st.session_state.modalidad = "Presencial y Virtual"
    st.session_state.acuerdo = "Acuerdo 012 de 2024"
    st.session_state.instancia = "Consejo Superior Universitario"
    st.session_state.registro1 = "Resolución 12345 de 2024"
    st.session_state.acreditacion1 = "Resolución 9876 de 2025"
    st.session_state.creditos = "160"
    st.session_state.periodicidad = "Semestral"
    st.session_state.lugares = "Bogotá y Medellín"
    st.session_state.snies = "102030"
    st.session_state.motivo = "Responder a la creciente demanda de transformación digital y automatización en el país."
    st.session_state.plan1_nom = "Plan Innova 2024"
    st.session_state.plan1_fec = "2024-01-15"

# --- LÓGICA API KEY ---
api_key = st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else ""

st.title("🎓 Generador PEP: Capítulo 1")
if st.button("📝 Llenar con datos de ejemplo"):
    cargar_ejemplo()

with st.form("pep_form"):
    col1, col2 = st.columns(2)
    
    with col1:
        denominacion = st.text_input("Denominación del programa", key="denominacion", help="Obligatorio (gris)")
        titulo = st.text_input("Título otorgado", key="titulo")
        nivel = st.selectbox("Nivel de formación", ["Técnico", "Tecnológico", "Profesional universitario", "Especialización", "Maestría", "Doctorado"], key="nivel")
        area = st.text_input("Área de formación", key="area")
        modalidad = st.selectbox("Modalidad de oferta", ["Presencial", "Virtual", "A Distancia", "Dual", "Presencial y Virtual", "Presencial y a Distancia", "Presencial y Dual"], key="modalidad")
        acuerdo = st.text_input("Acuerdo de creación (Norma interna)", key="acuerdo")
        instancia = st.text_input("Instancia interna que aprueba el Programa", key="instancia")
        registro1 = st.text_input("Resolución Registro calificado 1 (Número y año)", key="registro1")
        registro2 = st.text_input("Registro calificado 2 (Opcional)", key="registro2")
    
    with col2:
        acred1 = st.text_input("Resolución Acreditación en alta calidad 1 (Opcional)", key="acreditacion1")
        acred2 = st.text_input("Resolución Acreditación en alta calidad 2 (Opcional)", key="acreditacion2")
        creditos = st.text_input("Créditos académicos", key="creditos")
        periodicidad = st.selectbox("Periodicidad de admisión", ["Semestral", "Anual"], key="periodicidad")
        lugares = st.text_input("Lugares de desarrollo", key="lugares")
        snies = st.text_input("Código SNIES", key="snies")
        plan1_nom = st.text_input("Nombre del Plan de estudios versión 1", key="plan1_nom")
        plan1_fec = st.text_input("Fecha del Plan de estudios versión 1 (Año)", key="plan1_fec")

    motivo = st.text_area("Motivo de creación del Programa (Descripción amplia)", key="motivo")
    
    st.subheader("Reconocimientos (Opcional)")
    reconocimientos = []
    for i in range(2): # Ejemplo con 2 filas
        r_cols = st.columns(4)
        r_año = r_cols[0].text_input(f"Año {i+1}", key=f"r_año_{i}")
        r_nom = r_cols[1].text_input(f"Nombre Reconocimiento {i+1}", key=f"r_nom_{i}")
        r_gan = r_cols[2].text_input(f"Ganador {i+1}", key=f"r_gan_{i}")
        r_car = r_cols[3].selectbox(f"Cargo {i+1}", ["Docente", "Líder", "Decano", "Estudiante"], key=f"r_car_{i}")
        if r_nom: reconocimientos.append(f"{r_nom} otorgado a {r_gan} ({r_car}) en {r_año}")

    submit = st.form_submit_button("🚀 Generar Word")

if submit:
    doc = Document()
    
    # 1.1 Historia del Programa (Lógica de Plantilla)
    doc.add_heading('1.1. Historia del Programa', level=1)
    
    # Párrafo Base
    p1 = f"El Programa de {denominacion} fue creado mediante el {acuerdo} de {instancia} y aprobado mediante la {registro1} del Ministerio de Educación Nacional con Código SNIES {snies}."
    doc.add_paragraph(p1)

    # Párrafo Acreditación (Condicional)
    if acred1:
        p_acred = (f"El Programa desarrolla de manera permanente procesos de autoevaluación y autorregulación, "
                   f"orientados al aseguramiento de la calidad académica. Como resultado de estos procesos, "
                   f"el Programa obtuvo la Acreditación en Alta Calidad mediante {acred1}, como reconocimiento a la solidez de sus condiciones.")
        doc.add_paragraph(p_acred)

    # Párrafo Reconocimientos (Condicional)
    if reconocimientos:
        p_rec = f"El Programa de {denominacion} ha alcanzado importantes logros académicos e institucionales. Entre ellos se destacan: " + "; ".join(reconocimientos) + "."
        doc.add_paragraph(p_rec)

    # Línea de Tiempo
    doc.add_heading('Línea de tiempo de los principales hitos del Programa', level=2)
    doc.add_paragraph(f"• {plan1_fec[:4] if plan1_fec else '20XX'}: Creación del Programa y Registro Calificado")
    if acred1: doc.add_paragraph("• 20XX: Obtención de Acreditación de Alta Calidad")
    doc.add_paragraph(f"• {plan1_fec}: Implementación del Plan de estudios {plan1_nom}")

    # 1.2 Generalidades (Tabla de datos)
    doc.add_page_break()
    doc.add_heading('1.2. Generalidades del Programa', level=1)
    datos = {
        "Denominación": denominacion, "Título": titulo, "Nivel": nivel,
        "Área": area, "Modalidad": modalidad, "SNIES": snies, "Créditos": creditos
    }
    for k, v in datos.items():
        p = doc.add_paragraph()
        p.add_run(f"{k}: ").bold = True
        p.add_run(v)

    # Descarga
    output = io.BytesIO()
    doc.save(output)
    st.download_button("📥 Descargar PEP", output.getvalue(), "PEP_Cap1.docx")
