import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
import io
import time

st.set_page_config(page_title="Generador PEP Completo", page_icon="📚")
st.title("📚 Generador PEP - Versión Completa (12 Capítulos)")

# --- CONFIGURACIÓN ---
with st.sidebar:
    api_key = st.text_input("Ingresa tu Google API Key", type="password")

# --- LÓGICA IA ---
def redactar_capitulo(titulo_capitulo, insumos):
    """
    Recibe un título y una lista de respuestas del usuario.
    Genera el capítulo completo de una sola vez.
    """
    if not api_key: return "Falta API Key"
    
    # Unimos todas las respuestas del usuario en un solo texto
    texto_insumo = "\n".join([f"- {k}: {v}" for k, v in insumos.items()])
    
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
        Rol: Experto Curricular.
        Tarea: Redactar el CAPÍTULO: "{titulo_capitulo}" del PEP.
        
        INSUMOS DEL DIRECTOR:
        {texto_insumo}
        
        INSTRUCCIONES:
        1. Redacta un texto cohesivo, académico y formal.
        2. Integra los insumos en una narrativa fluida (no hagas lista de preguntas y respuestas).
        3. Extensión adecuada para un capítulo.
        """
        
        response = client.models.generate_content(
            model="gemini-flash-latest", 
            contents=prompt
        )
        return response.text
    except Exception as e:
        return f"Error: {str(e)}"

# --- ESTRUCTURA DE DATOS (AQUÍ DEFINES TUS 12 CAPÍTULOS) ---
# Puedes agregar tantos capítulos como quieras aquí abajo
estructura_pep = {
    "Capítulo 1: Identidad": [
        "¿Cuál es la Misión?", 
        "¿Cuál es la Visión?", 
        "¿Cuáles son los valores?"
    ],
    "Capítulo 2: Contexto Social": [
        "¿Cuál es la necesidad social del programa?",
        "¿Cuál es la población objetivo?"
    ],
    "Capítulo 3: Perfiles": [
        "Perfil de Ingreso",
        "Perfil de Egreso",
        "Perfil Ocupacional"
    ],
    # ... Agrega aquí tus otros capítulos ...
}

# --- INTERFAZ DINÁMICA ---
respuestas_usuario = {} # Aquí guardaremos todo

with st.form("form_pep_completo"):
    st.info("Responde por secciones para armar el documento completo.")
    
    # Este bucle crea los 12 capítulos en pantalla automáticamente
    for capitulo, preguntas in estructura_pep.items():
        with st.expander(capitulo, expanded=True):
            respuestas_usuario[capitulo] = {}
            for preg in preguntas:
                # Creamos un input único para cada pregunta
                respuestas_usuario[capitulo][preg] = st.text_area(preg, height=80)
    
    enviado = st.form_submit_button("🚀 Generar PEP Completo", type="primary")

# --- PROCESAMIENTO ---
if enviado and api_key:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    
    doc.add_heading('PROYECTO EDUCATIVO DEL PROGRAMA', 0)
    
    barra_progreso = st.progress(0)
    total_caps = len(estructura_pep)
    
    with st.status("Redactando capítulos...", expanded=True) as status:
        
        for i, (capitulo, datos) in enumerate(respuestas_usuario.items()):
            st.write(f"✍️ Redactando {capitulo}...")
            
            # Llamamos a la IA (1 llamada por capítulo, no por pregunta)
            texto_generado = redactar_capitulo(capitulo, datos)
            
            # Guardamos en el Word
            doc.add_heading(capitulo, level=1)
            doc.add_paragraph(texto_generado)
            doc.add_page_break()
            
            # Actualizamos barra
            barra_progreso.progress((i + 1) / total_caps)
            
            # Pausa inteligente (3 segundos entre capítulos es suficiente)
            time.sleep(3)
            
        status.update(label="¡Documento Completado!", state="complete")
    
    # Descarga
    bio = io.BytesIO()
    doc.save(bio)
    st.success("¡Tu PEP de 12 capítulos está listo!")
    st.download_button("📥 Descargar PEP Completo.docx", bio.getvalue(), "PEP_Completo.docx")