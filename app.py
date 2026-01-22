import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
import io
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador PEP", page_icon="📚", layout="wide")

st.title("📚 Generador de Proyecto Educativo del Programa (PEP)")
st.markdown("---")

# --- LÓGICA DE API KEY ---
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    with st.sidebar:
        st.header("Configuración")
        api_key = st.text_input("Ingresa tu Google API Key", type="password")

# --- ESTRUCTURA DINÁMICA DEL CAPÍTULO 1 ---
# Definimos los campos y sus tipos (text o select)
campos_info_programa = [
    {"label": "Denominación del programa", "tipo": "text"},
    {"label": "Título otorgado", "tipo": "text"},
    {"label": "Nivel de formación", "tipo": "select", "opciones": [
        "Técnico", "Tecnológico", "Profesional universitario", 
        "Especialización", "Maestría", "Doctorado"
    ]},
    {"label": "Área de formación", "tipo": "text"},
    {"label": "Modalidad de oferta", "tipo": "select", "opciones": [
        "Presencial", "Virtual", "A Distancia", "Dual", 
        "Presencial y Virtual", "Presencial y a Distancia", 
        "Presencial y Dual"
    ]},
    {"label": "Acuerdo de creación (Norma interna)", "tipo": "text"},
    {"label": "Registro calificado (Resolución MEN)", "tipo": "text"},
    {"label": "Créditos académicos", "tipo": "text"},
    {"label": "Periodicidad de admisión", "tipo": "select", "opciones": ["Semestral", "Anual"]},
    {"label": "Lugares de desarrollo", "tipo": "text"},
    {"label": "Código SNIES", "tipo": "text"},
]

# --- INTERFAZ DE USUARIO ---
respuestas_info = {}

with st.form("pep_form"):
    st.header("1. Información del Programa")
    
    # Creamos dos columnas para que el formulario no sea tan largo hacia abajo
    col1, col2 = st.columns(2)
    
    for i, campo in enumerate(campos_info_programa):
        # Alternamos entre columna 1 y columna 2
        target_col = col1 if i % 2 == 0 else col2
        
        with target_col:
            if campo["tipo"] == "text":
                respuestas_info[campo["label"]] = st.text_input(campo["label"])
            elif campo["tipo"] == "select":
                respuestas_info[campo["label"]] = st.selectbox(campo["label"], campo["opciones"])
    
    submit = st.form_submit_button("✨ Generar Documento PEP", type="primary")

# --- GENERACIÓN DEL DOCUMENTO ---
if submit:
    with st.status("📄 Generando documento...", expanded=True) as status:
        doc = Document()
        
        # Estilo de fuente global
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        
        # Título
        prog_nombre = respuestas_info["Denominación del programa"].upper()
        doc.add_heading(f'PROYECTO EDUCATIVO DEL PROGRAMA\n{prog_nombre}', 0)
        
        doc.add_heading('1. INFORMACIÓN DEL PROGRAMA', level=1)
        
        # Insertar los datos en formato de lista técnica
        for label, valor in respuestas_info.items():
            p = doc.add_paragraph()
            p.add_run(f"{label}: ").bold = True
            p.add_run(str(valor))
            
        status.update(label="¡Documento generado!", state="complete")
        
        # Descarga
        output = io.BytesIO()
        doc.save(output)
        st.success("✅ ¡Listo!")
        st.download_button(
            label="📥 Descargar Word",
            data=output.getvalue(),
            file_name=f"PEP_{prog_nombre.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


