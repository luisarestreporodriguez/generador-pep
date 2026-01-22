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

# --- LÓGICA DE API KEY (Nube + Local) ---
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    with st.sidebar:
        st.header("Configuración")
        api_key = st.text_input("Ingresa tu Google API Key", type="password")

# --- FUNCIÓN DE REDACCIÓN IA ---
def redactar_motivo_ia(nombre_prog, motivo_usuario):
    if not api_key: return "Error: No hay API Key."
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
        Actúa como un experto curricular. Redacta de forma académica y fluida 
        la sección 'Motivación de Creación' para el programa {nombre_prog}.
        Insumo del usuario: {motivo_usuario}
        Instrucción: Texto narrativo, formal, un solo párrafo de máximo 150 palabras.
        """
        response = client.models.generate_content(model="gemini-flash-latest", contents=prompt)
        return response.text
    except Exception as e:
        return f"Error en redacción: {str(e)}"

# --- FORMULARIO DE INFORMACIÓN DEL PROGRAMA ---
with st.form("pep_form"):
    st.header("1. Información del Programa")
    
    col1, col2 = st.columns(2)
    
    with col1:
        denominacion = st.text_input("Denominación del programa :gray[(Obligatorio)]")
        titulo = st.text_input("Título otorgado :gray[(Obligatorio)]")
        nivel = st.selectbox("Nivel de formación :gray[(Obligatorio)]", 
                            ["Técnico", "Tecnológico", "Profesional universitario", "Especialización", "Maestría", "Doctorado"])
        area = st.text_input("Área de formación :gray[(Obligatorio)]")
        modalidad = st.selectbox("Modalidad de oferta :gray[(Obligatorio)]", 
                               ["Presencial", "Virtual", "A Distancia", "Dual", "Presencial y Virtual", "Presencial y a Distancia", "Presencial y Dual"])
        acuerdo = st.text_input("Acuerdo de creación (Norma interna) :gray[(Obligatorio)]")
        acuerdo2 = st.text_input("Instancia que aprueba la creación del Programa (Interno) :gray[(Obligatorio)]")


    with col2:
        reg_1 = st.text_input("Registro calificado 1 (Resolución MEN) :gray[(Obligatorio)]")
        reg_2 = st.text_input("Registro calificado 2 (Resolución MEN) :gray[(Opcional)]")
        acred_1 = st.text_input("Acreditación en alta calidad 1 :gray[(Opcional)]")
        acred_2 = st.text_input("Acreditación en alta calidad 2 :gray[(Opcional)]")
        creditos = st.text_input("Créditos académicos :gray[(Obligatorio)]")
        periodicidad = st.selectbox("Periodicidad de admisión :gray[(Obligatorio)]", ["Semestral", "Anual"])
        lugares = st.text_input("Lugares de desarrollo :gray[(Obligatorio)]")
        snies = st.text_input("Código SNIES :gray[(Obligatorio)]")

    st.markdown("---")
    motivo_creacion = st.text_area("Motivo de creación del Programa :gray[(Obligatorio)]", 
                                   placeholder="Describa aquí las razones, necesidades o contexto que dieron origen al programa...",
                                   height=200)

    submit = st.form_submit_button("✨ Generar Documento PEP", type="primary")

# --- PROCESAMIENTO Y GENERACIÓN DE WORD ---
if submit:
    if not denominacion or not motivo_creacion or not api_key:
        st.error("⚠️ Por favor completa los campos obligatorios y asegúrate de tener la API Key.")
    else:
        with st.status("🚀 Procesando información...", expanded=True) as status:
            
            # 1. IA redacta el motivo
            st.write("✍️ Redactando narrativa del motivo de creación...")
            texto_motivo_ia = redactar_motivo_ia(denominacion, motivo_creacion)
            
            # 2. Crear documento Word
            doc = Document()
            doc.add_heading(f'PROYECTO EDUCATIVO DEL PROGRAMA\n{denominacion.upper()}', 0)
            
            # Sección Generalidades (Lista directa)
            doc.add_heading('1. Información General', level=1)
            datos_directos = [
                ("Denominación", denominacion),
                ("Título otorgado", titulo),
                ("Nivel de formación", nivel),
                ("Área de formación", area),
                ("Modalidad de oferta", modalidad),
                ("Acuerdo de creación", acuerdo),
                ("Registro calificado 1", reg_1),
                ("Registro calificado 2", reg_2),
                ("Acreditación 1", acred_1),
                ("Acreditación 2", acred_2),
                ("Créditos académicos", creditos),
                ("Periodicidad", periodicidad),
                ("Lugares de desarrollo", lugares),
                ("Código SNIES", snies),
            ]
            
            for etiqueta, valor in datos_directos:
                if valor: # Solo agrega si no está vacío
                    p = doc.add_paragraph()
                    p.add_run(f"{etiqueta}: ").bold = True
                    p.add_run(valor)

            # Sección redactada por IA
            doc.add_heading('2. Justificación y Motivos de Creación', level=1)
            doc.add_paragraph(texto_motivo_ia)
            
            status.update(label="¡Documento generado!", state="complete")

        # Descarga
        output = io.BytesIO()
        doc.save(output)
        st.success("✅ ¡Hecho! Descarga tu archivo aquí abajo.")
        st.download_button(
            label="📥 Descargar Word (.docx)",
            data=output.getvalue(),
            file_name=f"PEP_{denominacion.replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )



