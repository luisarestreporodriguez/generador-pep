import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
import io
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador PEP", page_icon="📚", layout="wide")

st.title("📚 Generador de Proyecto Educativo del Programa (PEP)")

# --- LÓGICA DE API KEY ---
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = st.sidebar.text_input("Ingresa tu Google API Key", type="password")

# --- FUNCIÓN DE REDACCIÓN ---
def redactar_historia_ia(datos):
    if not api_key: return "Error: Configura la API Key."
    
    try:
        client = genai.Client(api_key=api_key)
        
        # Construimos un contexto detallado para la IA
        prompt = f"""
        Eres un redactor académico. Tu tarea es unir y pulir la "Historia del Programa" usando estos datos exactos:
        - Programa: {datos['Denominación del programa']}
        - Acuerdo: {datos['Acuerdo de creación (Norma interna)']}
        - Instancia: {datos['Instancia interna que aprueba el Programa']}
        - Registro 1: {datos['Resolución Registro calificado 1']}
        - SNIES: {datos['Código SNIES']}
        - Acreditación 1: {datos.get('Resolución Acreditación en alta calidad 1', '')}
        - Motivo de creación: {datos['Motivo de creación del Programa']}
        
        INSTRUCCIONES:
        1. Usa este inicio: "El Programa de {datos['Denominación del programa']} fue creado mediante el {datos['Acuerdo de creación (Norma interna)']} de la {datos['Instancia interna que aprueba el Programa']} y aprobado mediante la {datos['Resolución Registro calificado 1']} del Ministerio de Educación Nacional con código SNIES {datos['Código SNIES']}."
        2. Si hay datos de Acreditación, incluye un párrafo sobre autoevaluación y alta calidad.
        3. Integra el 'Motivo de creación' de forma narrativa.
        4. Crea al final una "Línea de tiempo" con los hitos mencionados (años de creación, registros, etc).
        """
        
        response = client.models.generate_content(model="gemini-flash-latest", contents=prompt)
        return response.text
    except Exception as e:
        return f"Error en IA: {str(e)}"

# --- FORMULARIO ---
with st.form("pep_form"):
    st.header("1. Información del Programa")
    
    col1, col2 = st.columns(2)
    
    with col1:
        denom = st.text_input("Denominación del programa :small_red_triangle:", help="Obligatorio")
        titulo = st.text_input("Título otorgado :small_red_triangle:")
        nivel = st.selectbox("Nivel de formación :small_red_triangle:", ["Técnico", "Tecnológico", "Profesional universitario", "Especialización", "Maestría", "Doctorado"])
        area = st.text_input("Área de formación :small_red_triangle:")
        modalidad = st.selectbox("Modalidad de oferta :small_red_triangle:", ["Presencial", "Virtual", "A Distancia", "Dual", "Presencial y Virtual", "Presencial y a Distancia", "Presencial y Dual"])
        acuerdo = st.text_input("Acuerdo de creación (Norma interna) :small_red_triangle:")
        instancia = st.text_input("Instancia interna que aprueba el Programa :small_red_triangle:")

    with col2:
        reg1 = st.text_input("Resolución Registro calificado 1 (Número y año) :small_red_triangle:")
        reg2 = st.text_input("Registro calificado 2 (Número y año) :gray[(Opcional)]")
        acred1 = st.text_input("Resolución Acreditación en alta calidad 1 :gray[(Opcional)]")
        acred2 = st.text_input("Resolución Acreditación en alta calidad 2 :gray[(Opcional)]")
        creditos = st.text_input("Créditos académicos :small_red_triangle:")
        periodicidad = st.selectbox("Periodicidad de admisión :small_red_triangle:", ["Semestral", "Anual"])
        lugares = st.text_input("Lugares de desarrollo :small_red_triangle:")
        snies = st.text_input("Código SNIES :small_red_triangle:")

    motivo = st.text_area("Motivo de creación del Programa :small_red_triangle:", height=200)

    st.subheader("Reconocimientos :gray[(Opcional)]")
    st.info("Deja en blanco si no aplica.")
    reco_data = st.data_editor(
        [{"Año": "", "Nombre del reconocimiento": "", "Nombre del ganador": "", "Cargo": ""}],
        num_rows="dynamic",
        column_config={
            "Año": st.column_config.TextColumn("Año"),
            "Nombre del reconocimiento": st.column_config.TextColumn("Nombre"),
            "Nombre del ganador": st.column_config.TextColumn("Ganador"),
            "Cargo": st.column_config.TextColumn("Cargo")
        }
    )

    st.subheader("Planes de Estudio")
    p1_nom = st.text_input("Nombre del Plan de estudios versión 1 :small_red_triangle:")
    p1_fec = st.text_input("Fecha del Plan de estudios versión 1 :small_red_triangle:")
    p2_nom = st.text_input("Nombre del Plan de estudios versión 2 :gray[(Opcional)]")
    p2_fec = st.text_input("Fecha del Plan de estudios versión 2 :gray[(Opcional)]")
    
    st.markdown(":small_red_triangle: :gray[(Obligatorio)]")
    
    submit = st.form_submit_button("🚀 Generar PEP")

# --- GENERACIÓN DEL DOCUMENTO ---
if submit:
    datos = {
        "Denominación del programa": denom,
        "Acuerdo de creación (Norma interna)": acuerdo,
        "Instancia interna que aprueba el Programa": instancia,
        "Resolución Registro calificado 1": reg1,
        "Código SNIES": snies,
        "Resolución Acreditación en alta calidad 1": acred1,
        "Motivo de creación del Programa": motivo
    }

    with st.status("Generando documento...") as status:
        doc = Document()
        
        # 1.1 HISTORIA DEL PROGRAMA (Redactado por IA con tus reglas)
        doc.add_heading("1.1. Historia del Programa", level=1)
        historia_texto = redactar_historia_ia(datos)
        doc.add_paragraph(historia_texto)

        # SECCIÓN DE ACREDITACIÓN (Lógica condicional manual)
        if acred1:
            p_acred = doc.add_paragraph()
            p_acred.add_run("\nEl Programa desarrolla de manera permanente procesos de autoevaluación... ")
            p_acred.add_run(f"obtuvo la Acreditación en Alta Calidad mediante {acred1}.").bold = True

        # RECONOCIMIENTOS (Tabla)
        if any(row["Nombre del reconocimiento"] for row in reco_data):
            doc.add_heading("Reconocimientos", level=2)
            table = doc.add_table(rows=1, cols=4)
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Año'
            hdr_cells[1].text = 'Reconocimiento'
            hdr_cells[2].text = 'Ganador'
            hdr_cells[3].text = 'Cargo'
            
            for item in reco_data:
                if item["Nombre del reconocimiento"]:
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(item["Año"])
                    row_cells[1].text = item["Nombre del reconocimiento"]
                    row_cells[2].text = item["Nombre del ganador"]
                    row_cells[3].text = item["Cargo"]

        # 1.2 GENERALIDADES (Tabla de datos directos)
        doc.add_page_break()
        doc.add_heading("1.2. Generalidades del Programa", level=1)
        campos_directos = {
            "Título otorgado": titulo,
            "Nivel de formación": nivel,
            "Área de formación": area,
            "Modalidad de oferta": modalidad,
            "Créditos académicos": creditos,
            "Periodicidad de admisión": periodicidad,
            "Código SNIES": snies
        }
        for k, v in campos_directos.items():
            p = doc.add_paragraph()
            p.add_run(f"{k}: ").bold = True
            p.add_run(str(v))

        # DESCARGA
        output = io.BytesIO()
        doc.save(output)
        st.download_button("📥 Descargar Word", output.getvalue(), "PEP_Reestructurado.docx")
        status.update(label="¡Listo!", state="complete")


