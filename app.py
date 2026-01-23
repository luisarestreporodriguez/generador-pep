import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
import io
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador PEP", page_icon="📚", layout="wide")

st.title("📚 Generador PEP - Módulo 1: Información del Programa")

# --- LÓGICA DE API KEY (Nube + Local) ---
# Intentamos leer la clave desde los Secrets de Streamlit
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    with st.sidebar:
        st.header("Configuración Local")
        api_key = st.text_input("Ingresa tu Google API Key", type="password")
        if not api_key:
            st.warning("⚠️ Sin API Key la IA no podrá redactar textos largos.")
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
        "campos": [
            {"label": "Objeto de conocimiento del Programa", "req": True, "help": "¿Qué conoce, investiga y transforma este programa?"}
        ]
    },
    "2.2. Fundamentación epistemológica": {
        "tipo": "ia",
        "campos": [
            {"label": "Naturaleza epistemológica e identidad académica", "req": True},
            {"label": "Campo del saber y relación con ciencia/tecnología", "req": True}
        ]
    },
    "2.3. Fundamentación académica": {
        "tipo": "especial_pascual", # Nueva lógica para textos fijos + tabla
        "campos": [] 
    }
}

}


# --- BOTÓN DE DATOS DE EJEMPLO ---
# Usamos session_state para persistir los datos al hacer clic
if st.button("🧪 Llenar con datos de ejemplo"):
    st.session_state.ejemplo = {
        "denom": "Ingeniería de Sistemas",
        "titulo": "Ingeniero de Sistemas",
        "nivel_idx": 2, # Profesional universitario
        "area": "Ingeniería, Arquitectura y Urbanismo",
        "modalidad_idx": 4, # Presencial y Virtual
        "acuerdo": "Acuerdo 012 de 2015",
        "instancia": "Consejo Académico",
        "reg1": "Res. 4567 de 2016",
        "reg2": "Res. 8901 de 2023",
        "acred1": "Res. 00234 de 2024",
        "creditos": "165",
        "periodo_idx": 0, # Semestral
        "lugar": "Sede Principal (Cali)",
        "snies": "54321",
        "motivo": "El programa se fundamenta en la necesidad regional de formar profesionales capaces de liderar la transformación digital y el desarrollo de software de alta complejidad.",
        "p1_nom": "Acuerdo 012-2015", "p1_fec": "2015",
        "p2_nom": "Acuerdo 088-2020", "p2_fec": "2020",
        "p3_nom": "Acuerdo 102-2024", "p3_fec": "2024",
        #DATOS CAPÍTULO 2
        "objeto_con": "El programa investiga el ciclo de vida del software, la arquitectura de sistemas complejos y la integración de IA para transformar procesos industriales.",
        "fund_epi": "El programa se inscribe en el racionalismo crítico y el pragmatismo tecnológico, vinculando la ciencia de la computación con la ingeniería aplicada.",
        # DATOS PARA LAS TABLAS (Se guardan como listas de diccionarios)
        "tabla_recon_ej": [
            {"Año": "2024", "Nombre del premio": "Excelencia Académica", "Nombre del Ganador": "Juan Pérez", "Cargo": "Docente"}
        ],
        "tabla_cert_ej": [
            {"Nombre": "Desarrollador Web Junior", "Curso 1": "Programación I", "Créditos 1": 3, "Curso 2": "Bases de Datos", "Créditos 2": 4},
            {"Nombre": "Analista de Datos", "Curso 1": "Estadística", "Créditos 1": 4, "Curso 2": "Python para Ciencia", "Créditos 2": 4}
        ]
    }
    st.rerun()

# --- FORMULARIO DE ENTRADA ---
with st.form("pep_form"):
    ej = st.session_state.get("ejemplo", {})

    st.markdown("### 📋 1. Identificación General")
    col1, col2 = st.columns(2)
    with col1:
        denom = st.text_input("Denominación del programa (Obligatorio)", value=ej.get("denom", ""))
        titulo = st.text_input("Título otorgado (Obligatorio)", value=ej.get("titulo", ""))
        nivel = st.selectbox("Nivel de formación (Obligatorio)", 
                             ["Técnico", "Tecnológico", "Profesional universitario", "Especialización", "Maestría", "Doctorado"], 
                             index=ej.get("nivel_idx", 2))
        area = st.text_input("Área de formación (Obligatorio)", value=ej.get("area", ""))
    
    with col2:
        modalidad = st.selectbox("Modalidad de oferta (Obligatorio)", 
                                 ["Presencial", "Virtual", "A Distancia", "Dual", "Presencial y Virtual", "Presencial y a Distancia", "Presencial y Dual"],
                                 index=ej.get("modalidad_idx", 0))
        acuerdo = st.text_input("Acuerdo de creación / Norma interna (Obligatorio)", value=ej.get("acuerdo", ""))
        instancia = st.text_input("Instancia interna que aprueba (Obligatorio)", value=ej.get("instancia", ""))
        snies = st.text_input("Código SNIES (Obligatorio)", value=ej.get("snies", ""))

    st.markdown("---")
    st.markdown("### 📄 2. Registros, Acreditaciones y Tiempos")
    col3, col4 = st.columns(2)
    with col3:
        reg1 = st.text_input("Resolución Registro calificado 1 (Obligatorio)", value=ej.get("reg1", ""), placeholder="Número y año")
        reg2 = st.text_input("Registro calificado 2 (Opcional)", value=ej.get("reg2", ""))
        acred1 = st.text_input("Resolución Acreditación en alta calidad 1 (Opcional)", value=ej.get("acred1", ""))
        acred2 = st.text_input("Resolución Acreditación en alta calidad 2 (Opcional)", value="")

    with col4:
        creditos = st.text_input("Créditos académicos (Obligatorio)", value=ej.get("creditos", ""))
        periodicidad = st.selectbox("Periodicidad de admisión (Obligatorio)", ["Semestral", "Anual"], index=ej.get("periodo_idx", 0))
        lugares = st.text_input("Lugares de desarrollo (Obligatorio)", value=ej.get("lugar", ""))

    motivo = st.text_area("Motivo de creación del Programa (Obligatorio)", value=ej.get("motivo", ""), height=100)

    st.markdown("---")
    st.markdown("### 🧬 3. Planes de Estudios")
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        p1_nom = st.text_input("Nombre Plan v1 (Obligatorio)", value=ej.get("p1_nom", ""))
        p1_fec = st.text_input("Fecha/Año Plan v1 (Obligatorio)", value=ej.get("p1_fec", ""))
    with p_col2:
        p2_nom = st.text_input("Nombre Plan v2 (Opcional)", value=ej.get("p2_nom", ""))
        p2_fec = st.text_input("Fecha/Año Plan v2 (Opcional)", value=ej.get("p2_fec", ""))
    with p_col3:
        p3_nom = st.text_input("Nombre Plan v3 (Opcional)", value=ej.get("p3_nom", ""))
        p3_fec = st.text_input("Fecha/Año Plan v3 (Opcional)", value=ej.get("p3_fec", ""))

    st.markdown("---")
    st.markdown("### 🏆 4. Reconocimientos (Opcional)")
    recon_data = st.data_editor(
        [{"Año": "", "Nombre del premio": "", "Nombre del Ganador": "", "Cargo": "Estudiante"}],
        num_rows="dynamic",
        column_config={
            "Cargo": st.column_config.SelectboxColumn(options=["Docente", "Líder", "Decano", "Estudiante"])
        }
        )
  # --- CAPÍTULO 2 ---
  st.markdown("---")
    st.header("2. Referentes Conceptuales")
       # 2.1. Naturaleza del Programa
    objeto_con = st.text_area(
        "Objeto de conocimiento del Programa (Obligatorio)", 
        value=ej.get("objeto_con", ""), 
        help="¿Qué conoce, investiga y transforma?",
        key="input_objeto"
    ) 
   
#2.2. Fundamentación epistemológica
    fund_epi = st.text_area(
        "Fundamentación epistemológica (Instrucciones 1 y 2)",
        value=ej.get("fund_epi", ""), 
        key="input_epi")
    
   #Fundamentación académica 
    st.subheader("Certificaciones Temáticas Tempranas")
 cert_data = st.data_editor(
        ej.get("tabla_cert_ej", [{"Nombre": "", "Curso 1": "", "Créditos 1": 0, "Curso 2": "", "Créditos 2": 0}]),
        num_rows="dynamic",      
        key="editor_cert"
    )
    

    generar = st.form_submit_button("🚀 GENERAR DOCUMENTO PEP", type="primary")

# --- LÓGICA DE GENERACIÓN DEL WORD ---
if generar:
    if not denom or not reg1:
        st.error("⚠️ Falta información obligatoria (Denominación o Registro Calificado).")
    else:
        doc = Document()
        # Estilo base
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)

        # 1.1 Historia del Programa
        doc.add_heading("1.1. Historia del Programa", level=1)
        
        # Párrafo Base
        texto_historia = (
            f"El Programa de {denom} fue creado mediante el {acuerdo} del {instancia} "
            f"y aprobado mediante la resolución de Registro Calificado {reg1} del Ministerio de Educación Nacional "
            f"con código SNIES {snies}."
        )
        doc.add_paragraph(texto_historia)

        # Texto Condicional: Acreditación
        if acred1:
            texto_acred = (
                f"El Programa desarrolla de manera permanente procesos de autoevaluación y autorregulación, "
                f"orientados al aseguramiento de la calidad académica. Como resultado de estos procesos, "
                f"y tras demostrar el cumplimiento integral de los factores, características y lineamientos "
                f"de alta calidad establecidos por el Consejo Nacional de Acreditación (CNA), el Programa "
                f"obtuvo la Acreditación en Alta Calidad mediante {acred1}, como reconocimiento a la solidez "
                f"de sus condiciones académicas, administrativas y de impacto social."
            )
            doc.add_paragraph(texto_acred)

        # Texto Condicional: Evolución Curricular
        planes_fec = [f for f in [p1_fec, p2_fec, p3_fec] if f]
        planes_nom = [n for n in [p1_nom, p2_nom, p3_nom] if n]
        
        if len(planes_fec) > 0:
            texto_planes = (
                f"El plan de estudios del Programa de {denom} ha sido objeto de procesos periódicos de evaluación, "
                f"con el fin de asegurar su pertinencia académica y su alineación con los avances tecnológicos "
                f"y las demandas del entorno. Como resultado, se han realizado modificaciones curriculares "
                f"en los años {', '.join(planes_fec)}, aprobadas mediante Acuerdo(s) Nos. {', '.join(planes_nom)}."
            )
            doc.add_paragraph(texto_planes)

        # Texto Condicional: Reconocimientos
        recons_validos = [r for r in recon_data if r["Nombre del premio"].strip()]
        if recons_validos:
            doc.add_paragraph(
                f"El Programa de {denom} ha alcanzado importantes logros académicos e institucionales "
                f"que evidencian su calidad y compromiso con la excelencia. Entre ellos se destacan:"
            )
            for r in recons_validos:
                doc.add_paragraph(f"• {r['Nombre']} ({r['Año']}): Otorgado a {r['Ganador']}, en su calidad de {r['Cargo']}.", style='List Bullet')

        # Línea de tiempo
        doc.add_heading("Línea de tiempo de los principales hitos del Programa", level=2)
        doc.add_paragraph(f"{p1_fec}: Creación del Programa")
        doc.add_paragraph(f"{p1_fec}: Obtención del Registro Calificado")
        if p2_fec: doc.add_paragraph(f"{p2_fec}: Actualización del plan de estudios")
        if reg2: doc.add_paragraph(f"{reg2.split()[-1] if ' ' in reg2 else '20XX'}: Renovación del Registro Calificado")
        if recons_validos: doc.add_paragraph(f"{recons_validos[0]['Año']}: Reconocimientos académicos")

        # 1.2 Generalidades (Tabla de datos)
        doc.add_page_break()
        doc.add_heading("1.2 Generalidades del Programa", level=1)
        
        items_gen = [
            ("Denominación del programa", denom),
            ("Título otorgado", titulo),
            ("Nivel de formación", nivel),
            ("Área de formación", area),
            ("Modalidad de oferta", modalidad),
            ("Acuerdo de creación", acuerdo),
            ("Registro calificado", reg1),
            ("Créditos académicos", creditos),
            ("Periodicidad de admisión", periodicidad),
            ("Lugares de desarrollo", lugares),
            ("Código SNIES", snies)
        ]
        
        for k, v in items_gen:
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
            


        # Guardar archivo
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("✅ ¡Documento generado!")
        st.download_button(
            label="📥 Descargar Documento Word",
            data=bio.getvalue(),
            file_name=f"PEP_Modulo1_{denom.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

















