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
        r_año = r_cols[0].text_input(f"Año {i+1


