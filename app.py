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

# --- FUNCIÓN DE REDACCIÓN IA ---
def redactar_seccion_ia(titulo_seccion, datos_seccion):
    if not api_key: return "Error: No hay API Key configurada."
    respuestas_reales = {k: v for k, v in datos_seccion.items() if v.strip()}
    contexto = "\n".join([f"- {k}: {v}" for k, v in respuestas_reales.items()])
    
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
        Actúa como un Vicerrector Académico experto.
        Tarea: Redactar de forma narrativa y académica la sección "{titulo_seccion}" del PEP.
        DATOS: {contexto}
        REGLAS: Sin viñetas, párrafos fluidos, tono institucional.
        """
        response = client.models.generate_content(model="gemini-flash-latest", contents=prompt)
        return response.text
    except Exception as e:
        return f"Error en redacción: {str(e)}"

# --- ESTRUCTURA DEL PEP (12 CAPÍTULOS) ---
# 'tipo': 'ia' -> Pasa por la IA para generar párrafos
# 'tipo': 'directo' -> Se pone tal cual en el Word (Pregunta: Respuesta)
estructura_pep = {
    "1. Referentes Históricos": {
        "1.1. Historia del programa": {
            "tipo": "ia",
            "campos": [
                {"label": "Año de creación del Programa", "req": True},
                {"label": "Motivación para la creación del Programa", "req": True},
                {"label": "Resolución e instancia que aprueba la creación", "req": True},
                {"label": "Resolución de aprobación del Programa MEN", "req": True},
                {"label": "Reconocimientos o modificaciones", "req": False},
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
    nombre_prog = st.text_input("Nombre completo del Programa Académico", placeholder="Ej: Ingeniería de Sistemas")
    
    for cap, secciones in estructura_pep.items():
        st.header(cap)
        for seccion_titulo, config in secciones.items():
            with st.expander(f"Completar: {seccion_titulo}", expanded=True):
                respuestas_finales[seccion_titulo] = {}
                for campo in config["campos"]:
                    label = f"{campo['label']} {'*' if campo['req'] else '(Opcional)'}"
                    respuestas_finales[seccion_titulo][campo['label']] = st.text_input(label, key=f"{seccion_titulo}_{campo['label']}")
    
    submit = st.form_submit_button("✨ Generar Documento PEP", type="primary")

# --- PROCESAMIENTO Y WORD ---
if submit:
    if not api_key:
        st.error("Por favor, configura la API Key.")
    else:
        with st.status("🚀 Procesando documento...", expanded=True) as status:
            doc = Document()
            # Título principal
            titulo = doc.add_heading(f'PROYECTO EDUCATIVO DEL PROGRAMA\n{nombre_prog.upper()}', 0)
            
            for cap_nombre, secciones in estructura_pep.items():
                doc.add_heading(cap_nombre, level=1)
                
                for seccion_nombre, config in secciones.items():
                    doc.add_heading(seccion_nombre, level=2)
                    
                    datos_usuario = respuestas_finales[seccion_nombre]
                    
                    if config["tipo"] == "ia":
                        st.write(f"✍️ Redactando narrativa para: {seccion_nombre}...")
                        texto_ia = redactar_seccion_ia(seccion_nombre, datos_usuario)
                        doc.add_paragraph(texto_ia)
                        time.sleep(3) # Pausa anti-bloqueo
                    
                    else:
                        st.write(f"📋 Tabulando datos para: {seccion_nombre}...")
                        # En el tipo 'directo', iteramos y ponemos Pregunta: Respuesta
                        for campo, valor in datos_usuario.items():
                            p = doc.add_paragraph()
                            p.add_run(f"{campo}: ").bold = True
                            p.add_run(valor if valor else "No especificado")
            
            status.update(label="¡PEP Generado!", state="complete")
        
        # Descarga
        output = io.BytesIO()
        doc.save(output)
        st.success("✅ Documento listo para descargar.")
        st.download_button(
            label="📥 Descargar Word (.docx)",
            data=output.getvalue(),
            file_name=f"PEP_{nombre_prog.replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


