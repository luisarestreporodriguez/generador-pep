import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt, Inches
import io
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador PEP - Pascual Bravo", page_icon="📚", layout="wide")

st.title("📚 Generador PEP - I.U. Pascual Bravo")
st.markdown("---")

# --- LÓGICA DE API KEY ---
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    with st.sidebar:
        st.header("Configuración")
        api_key = st.text_input("Ingresa tu Google API Key", type="password")

# --- FUNCIÓN DE REDACCIÓN IA ---
def redactar_seccion_ia(titulo_seccion, datos_seccion):
    if not api_key: return "Error: No hay API Key configurada."
    respuestas_reales = {k: v for k, v in datos_seccion.items() if str(v).strip()}
    contexto = "\n".join([f"- {k}: {v}" for k, v in respuestas_reales.items()])
    
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
        Actúa como un Vicerrector Académico experto en aseguramiento de la calidad.
        Tarea: Redactar la sección "{titulo_seccion}" de un Proyecto Educativo del Programa (PEP).
        DATOS SUMINISTRADOS:
        {contexto}
        INSTRUCCIONES:
        1. Usa un lenguaje académico, técnico y fluido.
        2. NO uses listas. Redacta párrafos cohesivos.
        3. Si la información es breve, elabórala respetando la esencia.
        4. Tono institucional de la I.U. Pascual Bravo.
        """
        response = client.models.generate_content(model="gemini-flash-latest", contents=prompt)
        return response.text
    except Exception as e:
        return f"Error en redacción: {str(e)}"

# --- ESTRUCTURA DE CONTENIDOS ---
estructura_pep = {
    "1. Información del Programa": {
        "1.1. Historia del Programa": {"tipo": "especial_historia"},
        "1.2. Generalidades del Programa": {"tipo": "directo"}
    },
    "2. Referentes Conceptuales": {
        "2.1. Naturaleza del Programa": {
            "tipo": "ia",
            "campos": ["Objeto de conocimiento del Programa"]
        },
        "2.2. Fundamentación epistemológica": {
            "tipo": "ia",
            "campos": ["Naturaleza epistemológica e identidad académica", "Relación con desarrollos científicos/tecnológicos"]
        },
        "2.3. Fundamentación académica": {"tipo": "especial_pascual"}
    }
}

# --- INTERFAZ DE USUARIO ---
if st.button("🧪 Llenar con datos de ejemplo"):
    st.session_state.ejemplo = {
        "denom": "Ingeniería de Software", "titulo": "Ingeniero de Software", "nivel": "Profesional universitario",
        "modalidad": "Presencial", "acuerdo": "Acuerdo 001 de 2022", "instancia": "Consejo Directivo",
        "reg1": "Res. 123 de 2023", "snies": "109283", "motivo": "Demanda de perfiles DevOps en la región.",
        "p1_fec": "2023", "p1_nom": "Plan V1", "acred1": "En proceso", "area": "Ingeniería", "creditos": "160", "lugar": "Medellín"
    }
    st.rerun()

with st.form("pep_form"):
    ej = st.session_state.get("ejemplo", {})
    
    # --- CAPÍTULO 1 ---
    st.header("1. Información del Programa")
    col1, col2 = st.columns(2)
    with col1:
        denom = st.text_input("Denominación del programa (Obligatorio)", value=ej.get("denom", ""))
        titulo = st.text_input("Título otorgado (Obligatorio)", value=ej.get("titulo", ""))
        nivel = st.selectbox("Nivel de formación (Obligatorio)", ["Técnico", "Tecnológico", "Profesional universitario", "Especialización", "Maestría"], index=2)
        area = st.text_input("Área de formación (Obligatorio)", value=ej.get("area", ""))
    with col2:
        modalidad = st.selectbox("Modalidad de oferta (Obligatorio)", ["Presencial", "Virtual", "A Distancia", "Dual"], index=0)
        acuerdo = st.text_input("Acuerdo de creación (Obligatorio)", value=ej.get("acuerdo", ""))
        instancia = st.text_input("Instancia que aprueba (Obligatorio)", value=ej.get("instancia", ""))
        snies = st.text_input("Código SNIES (Obligatorio)", value=ej.get("snies", ""))

    reg1 = st.text_input("Registro calificado 1 (Obligatorio)", value=ej.get("reg1", ""))
    acred1 = st.text_input("Acreditación en alta calidad 1 (Opcional)", value=ej.get("acred1", ""))
    creditos = st.text_input("Créditos (Obligatorio)", value=ej.get("creditos", ""))
    lugares = st.text_input("Lugares de desarrollo (Obligatorio)", value=ej.get("lugar", ""))
    
    st.subheader("Planes de Estudio")
    p1_fec = st.text_input("Año Plan v1 (Obligatorio)", value=ej.get("p1_fec", ""))
    p1_nom = st.text_input("Nombre Plan v1 (Obligatorio)", value=ej.get("p1_nom", ""))

    # --- CAPÍTULO 2 ---
    st.header("2. Referentes Conceptuales")
    objeto_con = st.text_area("Objeto de conocimiento del Programa (Obligatorio)", help="¿Qué conoce, investiga y transforma?")
    fund_epi = st.text_area("Fundamentación epistemológica (Instrucciones 1 y 2)")
    
    st.subheader("Certificaciones Temáticas Tempranas")
    cert_data = st.data_editor(
        [{"Nombre": "", "Curso 1": "", "Créditos 1": 0, "Curso 2": "", "Créditos 2": 0}],
        num_rows="dynamic", key="editor_cert"
    )

    submit = st.form_submit_button("🚀 Generar PEP Completo")

# --- PROCESAMIENTO WORD ---
if submit:
    doc = Document()
    
    # 1.1 Historia
    doc.add_heading("1.1. Historia del Programa", level=2)
    historia_txt = (f"El Programa de {denom} fue creado mediante el {acuerdo} del {instancia} "
                    f"y aprobado mediante el Registro Calificado {reg1} con SNIES {snies}.")
    doc.add_paragraph(historia_txt)
    if acred1:
        doc.add_paragraph(f"El programa obtuvo Acreditación en Alta Calidad mediante {acred1}...")

    # 1.2 Generalidades
    doc.add_heading("1.2. Generalidades del Programa", level=2)
    for k, v in [("Título", titulo), ("Nivel", nivel), ("Modalidad", modalidad), ("Créditos", creditos)]:
        p = doc.add_paragraph()
        p.add_run(f"{k}: ").bold = True
        p.add_run(str(v))

    # 2.1 Naturaleza
    doc.add_heading("2.1. Naturaleza del Programa", level=2)
    doc.add_paragraph(redactar_seccion_ia("Naturaleza del Programa", {"Objeto": objeto_con}))

    # 2.2 Epistemología
    doc.add_heading("2.2. Fundamentación epistemológica", level=2)
    doc.add_paragraph(redactar_seccion_ia("Fundamentación Epistemológica", {"Datos": fund_epi}))

    # 2.3 Fundamentación Académica (TEXTO FIJO PASCUAL BRAVO)
    doc.add_heading("2.3. Fundamentación académica", level=2)
    doc.add_paragraph("La fundamentación académica del Programa responde a los Lineamientos Académicos y Curriculares (LAC) de la I.U. Pascual Bravo...")
    doc.add_paragraph("Dentro de los LAC se establece la política de créditos académicos...")
    
    doc.add_heading("Rutas educativas: Certificaciones Temáticas Tempranas", level=3)
    doc.add_paragraph("Las Certificaciones Temáticas Tempranas son el resultado del agrupamiento de competencias...")
    
    # Tabla de Certificaciones
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = 'Certificación', 'Cursos', 'Créditos Totales'
    
    for c in cert_data:
        if c["Nombre"]:
            row = table.add_row().cells
            row[0].text = c["Nombre"]
            row[1].text = f"{c['Curso 1']}, {c['Curso 2']}"
            row[2].text = str(c["Créditos 1"] + c["Créditos 2"])

    # Descarga
    bio = io.BytesIO()
    doc.save(bio)
    st.success("✅ Documento generado.")
    st.download_button("📥 Descargar PEP", bio.getvalue(), f"PEP_{denom}.docx")

