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
        "p3_nom": "Acuerdo 102-2024", "p3_fec": "2024"
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
        [{"Año": "", "Nombre": "", "Ganador": "", "Cargo": "Estudiante"}],
        num_rows="dynamic",
        column_config={
            "Cargo": st.column_config.SelectboxColumn(options=["Docente", "Líder", "Decano", "Estudiante"])
        }
    )

    generar = st.form_submit_button("🚀 GENERAR DOCUMENTO WORD", type="primary")

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
        recons_validos = [r for r in recon_data if r["Nombre"].strip()]
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
