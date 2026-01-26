import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
import io

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador PEP", page_icon="📚", layout="wide")
st.title("Generador PEP - Módulo 1: Información del Programa")

# --- LÓGICA DE API KEY ---
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    with st.sidebar:
        st.header("Configuración Local")
        api_key = st.text_input("Ingresa tu Google API Key", type="password")

# --- FUNCIÓN DE REDACCIÓN IA ---
def redactar_seccion_ia(titulo_seccion, datos_seccion):
    if not api_key: return "Clave de API no configurada. No se pudo generar el texto."
    respuestas_reales = {k: v for k, v in datos_seccion.items() if str(v).strip()}
    contexto = "\n".join([f"- {k}: {v}" for k, v in respuestas_reales.items()])
    
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
        Actúa como un experto en redacción académica.
        Tarea: Redactar un párrafo sobre {titulo_seccion}.
        DATOS: {contexto}
        REGLAS: UN SOLO PÁRRAFO, sin títulos, sin negritas, tono formal. Máximo 150 palabras.
        """
        response = client.models.generate_content(model="gemini-1.5-flash", contents=prompt)
        return response.text.strip()
    except Exception as e:
        return f"Error en redacción: {str(e)}"

# --- BOTÓN DE DATOS DE EJEMPLO ---
if st.button("🧪 Llenar con datos de ejemplo"):
    st.session_state.ejemplo = {
        "denom": "Ingeniería de Sistemas",
        "titulo": "Ingeniero de Sistemas",
        "nivel_idx": 2,
        "area": "Ingeniería, Arquitectura y Urbanismo",
        "modalidad_idx": 4,
        "acuerdo": "Acuerdo 012 de 2015",
        "instancia": "Consejo Académico",
        "reg1": "Res. 4567 de 2016",
        "reg2": "Res. 8901 de 2023",
        "acred1": "Res. 00234 de 2024",
        "creditos": "165",
        "periodo_idx": 0,
        "lugar": "Sede Principal (Medellín)",
        "snies": "54321",
        "motivo": "Necesidad regional de transformación digital.",
        "p1_nom": "Plan 2015", "p1_fec": "2015",
        "p2_nom": "Plan 2020", "p2_fec": "2020",
        "p3_nom": "Plan 2024", "p3_fec": "2024",
        "objeto_con": "Investigación del ciclo de vida del software.",
        "fund_epi": "Racionalismo crítico aplicado.",
        "tabla_cert_ej": [{"Nombre": "Dev Junior", "Curso 1": "Programación", "Créditos 1": 4, "Curso 2": "Web", "Créditos 2": 4}]
    }
    st.rerun()

# --- FORMULARIO ---
with st.form("pep_form"):
    ej = st.session_state.get("ejemplo", {})
    col1, col2 = st.columns(2)
    with col1:
        denom = st.text_input("Denominación (Obligatorio)", value=ej.get("denom", ""))
        titulo = st.text_input("Título (Obligatorio)", value=ej.get("titulo", ""))
        nivel = st.selectbox("Nivel", ["Técnico", "Tecnológico", "Profesional universitario"], index=ej.get("nivel_idx", 2))
        area = st.text_input("Área", value=ej.get("area", ""))
    with col2:
        modalidad = st.selectbox("Modalidad", ["Presencial", "Virtual", "Dual"], index=0)
        acuerdo = st.text_input("Acuerdo de creación", value=ej.get("acuerdo", ""))
        instancia = st.text_input("Instancia", value=ej.get("instancia", ""))
        snies = st.text_input("SNIES", value=ej.get("snies", ""))

    st.markdown("### 🧬 Registros y Planes")
    c3, c4 = st.columns(2)
    with c3:
        reg1 = st.text_input("Registro Calificado 1", value=ej.get("reg1", ""))
        reg2 = st.text_input("Registro Calificado 2", value=ej.get("reg2", ""))
    with c4:
        acred1 = st.text_input("Acreditación 1", value=ej.get("acred1", ""))
        acred2 = st.text_input("Acreditación 2", value="")
    
    creditos = st.text_input("Créditos", value=ej.get("creditos", ""))
    periodicidad = st.selectbox("Periodicidad", ["Semestral", "Anual"])
    lugares = st.text_input("Lugar", value=ej.get("lugar", ""))
    motivo = st.text_area("Motivo de creación", value=ej.get("motivo", ""))

    st.markdown("### 📅 Evolución del Plan")
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        p1_nom = st.text_input("Nombre Plan 1", value=ej.get("p1_nom", ""))
        p1_fec = st.text_input("Año Plan 1", value=ej.get("p1_fec", ""))
    with p_col2:
        p2_nom = st.text_input("Nombre Plan 2", value=ej.get("p2_nom", ""))
        p2_fec = st.text_input("Año Plan 2", value=ej.get("p2_fec", ""))
    with p_col3:
        p3_nom = st.text_input("Nombre Plan 3", value=ej.get("p3_nom", ""))
        p3_fec = st.text_input("Año Plan 3", value=ej.get("p3_fec", ""))

    st.markdown("### 🏆 Reconocimientos")
    recon_data = st.data_editor([{"Año": "", "Nombre del premio": "", "Nombre del Ganador": "", "Cargo": "Docente"}], num_rows="dynamic")

    st.markdown("### 🧠 Capítulo 2")
    objeto_con = st.text_area("Objeto de conocimiento", value=ej.get("objeto_con", ""))
    fund_epi = st.text_area("Fundamentación Epistemológica", value=ej.get("fund_epi", ""))
    cert_data = st.data_editor(ej.get("tabla_cert_ej", [{"Nombre": "", "Curso 1": "", "Créditos 1": 0, "Curso 2": "", "Créditos 2": 0}]), num_rows="dynamic")

    generar = st.form_submit_button("🚀 GENERAR DOCUMENTO PEP", type="primary")

# --- LÓGICA DE GENERACIÓN ---
if generar:
    if not denom or not reg1:
        st.error("⚠️ Falta información obligatoria.")
    else:
        doc = Document()
        # 1.1 Historia
        doc.add_heading("1.1. Historia del Programa", level=1)
        doc.add_paragraph(f"El Programa de {denom} fue creado mediante el {acuerdo} del {instancia} y aprobado mediante la resolución {reg1} (SNIES {snies}).")

        if motivo:
            st.write("✍️ Redactando motivo con IA...")
            doc.add_paragraph(redactar_seccion_ia("Motivo de Creación", {"Motivo": motivo}))

        # Acreditaciones
        if acred1:
            if not acred2:
                doc.add_paragraph(f"El programa obtuvo la Acreditación en alta calidad mediante {acred1}.")
            else:
                doc.add_paragraph(f"El programa obtuvo su primera acreditación mediante {acred1} y fue renovada mediante {acred2}.")

        # Evolución Curricular
        planes_nom = [n for n in [p1_nom, p2_nom, p3_nom] if n]
        planes_fec = [f for f in [p1_fec, p2_fec, p3_fec] if f]
        
        if planes_nom:
            txt_lista = ", ".join(planes_nom[:-1]) + f" y {planes_nom[-1]}" if len(planes_nom) > 1 else planes_nom[0]
            doc.add_paragraph(f"Se han realizado modificaciones curriculares en los años {', '.join(planes_fec)}, aprobadas mediante el {txt_lista}, respectivamente.")

        # Reconocimientos
        recons = [r for r in recon_data if r.get("Nombre del premio", "").strip()]
        if recons:
            doc.add_paragraph("Logros destacados:")
            for r in recons:
                doc.add_paragraph(f"• {r['Nombre del premio']} ({r['Año']}) - {r['Nombre del Ganador']}", style='List Bullet')

        # Generalidades
        doc.add_page_break()
        doc.add_heading("1.2 Generalidades", level=1)
        items = [("Título", titulo), ("SNIES", snies), ("Modalidad", modalidad), ("Créditos", creditos)]
        for k, v in items:
            p = doc.add_paragraph()
            p.add_run(f"{k}: ").bold = True
            p.add_run(str(v))

        # Capítulo 2
        doc.add_heading("2. Referentes Conceptuales", level=1)
        doc.add_paragraph(redactar_seccion_ia("Naturaleza", {"Objeto": objeto_con}))
        doc.add_paragraph(redactar_seccion_ia("Epistemología", {"Datos": fund_epi}))

        # Descarga
        bio = io.BytesIO()
        doc.save(bio)
        st.success("✅ ¡Documento generado!")
        st.download_button("📥 Descargar Word", bio.getvalue(), f"PEP_{denom}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")








































