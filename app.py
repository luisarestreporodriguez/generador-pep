import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
import io
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador PEP Institucional", page_icon="📚", layout="wide")

st.title("📚 Generador de Proyecto Educativo del Programa (PEP)")
st.markdown("---")

# --- LÓGICA DE API KEY (Nube + Local) ---
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    with st.sidebar:
        st.header("Configuración")
        api_key = st.text_input("Ingresa tu Google API Key", type="password")

# --- FUNCIÓN DE REDACCIÓN ---
def redactar_seccion_ia(titulo_seccion, datos_seccion):
    if not api_key: return "Error: No hay API Key configurada."
    
    # Filtramos solo las respuestas que el usuario llenó
    respuestas_reales = {k: v for k, v in datos_seccion.items() if v.strip()}
    
    # Convertimos los datos en texto para el prompt
    contexto = "\n".join([f"- {k}: {v}" for k, v in respuestas_reales.items()])
    
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
        Actúa como un Vicerrector Académico experto en aseguramiento de la calidad académica en universidad.
        Tarea: Redactar de forma narrativa y fluida la sección "{titulo_seccion}" del PEP.
        
        DATOS SUMINISTRADOS:
        {contexto}
        
        INSTRUCCIONES DE REDACCIÓN:
        1. NO uses listas ni viñetas. Crea párrafos académicos cohesivos.
        2. Menciona fechas y números de resolución de forma natural dentro del texto.
        3. Si la información es breve, compleméntala con un tono institucional formal.
        4. Si algún dato no fue suministrado, no lo menciones ni inventes información.
        """
        
        response = client.models.generate_content(
            model="gemini-flash-latest", 
            contents=prompt
        )
        return response.text
    except Exception as e:
        return f"Error en redacción: {str(e)}"

# --- ESTRUCTURA DE LOS 12 CAPÍTULOS ---
# Aquí puedes ir agregando los demás capítulos siguiendo el mismo formato
estructura_pep = {
    "1. Referentes Históricos": {
        "1.1. Historia del programa": [
            {"label": "Año de creación del Programa", "req": True},
            {"label": "Motivación para la creación del Programa", "req": True},
            {"label": "Resolución e instancia que aprueba la creación", "req": True},
            {"label": "Resolución de aprobación del Programa MEN", "req": True},
            {"label": "Resolución de modificación del plan de estudios (1)", "req": False},
            {"label": "Resolución de modificación del plan de estudios (2)", "req": False},
            {"label": "Resolución de modificación del plan de estudios (3)", "req": False},
            {"label": "Reconocimientos", "req": False},
            {"label": "Resolución de acreditación del Programa (1)", "req": False},
            {"label": "Resolución de acreditación del Programa (2)", "req": False},
        ]
    },
"1.2. Generalidades del Programa": {
            "tipo": "directo",
            "campos": [
                {"label": "Denominación del programa", "req": True},
                {"label": "Título otorgado", "req": True},
                {"label": "Nivel de formación", "req": True},
                {"label": "Área de formación", "req": True},
                {"label": "Modalidad de oferta", "req": True},
                {"label": "Acuerdo de creación (Norma interna)", "req": True},
                {"label": "Registro calificado (Resolución MEN)", "req": True},
                {"label": "Créditos académicos", "req": True},
                {"label": "Periodicidad de admisión", "req": True},
                {"label": "Lugares de desarrollo", "req": True},
                {"label": "Código SNIES", "req": True},
            ]
        }
    }
}

# --- INTERFAZ DE USUARIO ---
respuestas_finales = {}

with st.form("pep_form"):
    st.subheader("Información General")
    nombre_prog = st.text_input("Nombre completo del Programa Académico")
    
    # Generar inputs dinámicamente según la estructura
    for cap, secciones in estructura_pep.items():
        st.header(cap)
        for seccion, campos in secciones.items():
            with st.expander(f"Completar: {seccion}", expanded=True):
                respuestas_finales[seccion] = {}
                for campo in campos:
                    label = f"{campo['label']} {'*' if campo['req'] else '(Opcional)'}"
                    respuestas_finales[seccion][campo['label']] = st.text_area(label, height=70, key=f"{seccion}_{campo['label']}")
    
    submit = st.form_submit_button("✨ Generar Documento Académico", type="primary")

# --- PROCESAMIENTO Y WORD ---
if submit:
    if not api_key:
        st.error("Por favor, configura la API Key.")
    else:
        with st.status("🤖 La IA está redactando los capítulos...", expanded=True) as status:
            doc = Document()
            doc.add_heading(f'PROYECTO EDUCATIVO DEL PROGRAMA\n{nombre_prog.upper()}', 0)
            
            for cap_nombre, secciones in estructura_pep.items():
                doc.add_heading(cap_nombre, level=1)
                
                for seccion_nombre in secciones.keys():
                    st.write(f"Redactando: {seccion_nombre}...")
                    
                    # Llamada a la IA por cada subsección
                    texto_ia = redactar_seccion_ia(seccion_nombre, respuestas_finales[seccion_nombre])
                    
                    doc.add_heading(seccion_nombre, level=2)
                    doc.add_paragraph(texto_ia)
                    
                    # Pausa para evitar bloqueos de cuota
                    time.sleep(4)
            
            status.update(label="¡Redacción completa!", state="complete")
        
        # Guardar y Descargar
        output = io.BytesIO()
        doc.save(output)
        st.success("✅ El documento ha sido generado exitosamente.")
        st.download_button(
            label="📥 Descargar PEP (.docx)",
            data=output.getvalue(),
            file_name=f"PEP_{nombre_prog.replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
