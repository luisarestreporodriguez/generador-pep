
import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
import requests
import io
import time
import re 
import os
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH
from collections import defaultdict

# SECCIÓN: HELPERS

def nested_dict():
    """Crea un diccionario infinito para guardar la estructura del Word."""
    return defaultdict(nested_dict)

def is_noise(title):
    """Detecta si un Heading es ruido (tablas, figuras, etc.)."""
    title = title.strip().lower()
    if not title:
        return True
    # Filtramos leyendas comunes que Word a veces confunde con títulos
    ruido = ["tabla", "figura", "imagen", "ilustración", "gráfico", "anexo"]
    return any(title.startswith(r) for r in ruido)

def clean_dict(d):
    """Convierte defaultdict a dict normal y elimina secciones vacías."""
    if not isinstance(d, dict):
        return d
    cleaned = {}
    for k, v in d.items():
        if k == "_content":
            if v.strip():
                cleaned[k] = v.strip()
            continue
        child = clean_dict(v)
        if child:
            cleaned[k] = child
    return cleaned

def docx_to_clean_dict(path):
    """Analiza el Documento Maestro y crea un mapa jerárquico por estilos."""
    doc = Document(path)
    estructura = nested_dict()
    stack = []

    for para in doc.paragraphs:
        text = para.text.strip()
        style = para.style.name

        # Buscamos estilos que empiecen por 'Heading' o 'Título'
        if "Heading" in style or "Título" in style:
            if is_noise(text):
                if stack:
                    current = estructura
                    for item in stack: current = current[item]
                    current["_content"] += text + "\n"
                continue

            try:
                # Intentamos extraer el nivel (ej: 'Heading 2' -> 2)
                level = int(''.join(filter(str.isdigit, style)))
            except:
                level = 1 # Por defecto si no tiene número

            stack = stack[:level-1]
            stack.append(text)

            current = estructura
            for item in stack:
                current = current[item]
            current["_content"] = ""

        else:
            # Es un párrafo normal: se añade al contenido de la sección actual
            if stack and text:
                current = estructura
                for item in stack:
                    current = current[item]
                if "_content" not in current:
                    current["_content"] = ""
                current["_content"] += text + "\n"

    return clean_dict(estructura)

def buscar_contenido_por_titulo(diccionario, titulo_objetivo):
    # 1. Limpiamos el objetivo: solo palabras clave
    palabras_clave = ["conceptualización", "teórica", "epistemológica"]
    
    def extraer_recursivo(nodo):
        texto = nodo.get("_content", "") + "\n"
        for k, v in nodo.items():
            if k != "_content":
                texto += f"\n{k}\n" + extraer_recursivo(v)
        return texto

    for titulo_real, contenido in diccionario.items():
        titulo_min = titulo_real.lower()
        
        # Verificamos si las 3 palabras clave están en el título del Word
        if all(p in titulo_min for p in palabras_clave):
            return extraer_recursivo(contenido)
        
        # Si no, buscamos en los hijos
        if isinstance(contenido, dict):
            res = buscar_contenido_por_titulo(contenido, titulo_objetivo)
            if res: return res
    return ""

    # Bucle principal de búsqueda en el diccionario
    for titulo_real, contenido in diccionario.items():
        titulo_limpio = " ".join(titulo_real.lower().split())
        
        # Si encontramos el título que buscamos (o parte de él)
        if target in titulo_limpio:
            # Llamamos a la función interna para recoger todo lo que hay dentro
            return extraer_todo_el_texto(contenido)
        
        # Si no es el título, pero hay un diccionario dentro, buscamos en los hijos
        if isinstance(contenido, dict):
            res = buscar_contenido_por_titulo(contenido, titulo_objetivo)
            if res: 
                return res
    return ""
    
def obtener_solo_estructura(d):
    """
    Crea una copia del diccionario que contiene solo los títulos, 
    eliminando las claves '_content'.
    """
    if not isinstance(d, dict):
        return d
    # Filtramos para dejar solo las llaves que no son '_content'
    return {k: obtener_solo_estructura(v) for k, v in d.items() if k != "_content"}                



#FUNCIÓN PARA INSERTAR TEXTO DEBAJO DE UN TÍTULO ESPECÍFICO
def insertar_texto_debajo_de_titulo(doc, texto_titulo_buscar, texto_nuevo):
    encontrado = False
    for i, paragraph in enumerate(doc.paragraphs):
        # Busca el título (ignorando mayúsculas/minúsculas)
        if texto_titulo_buscar.lower() in paragraph.text.lower():
            # Si hay un párrafo siguiente, inserta ANTES de él (para quedar debajo del título)
            if i + 1 < len(doc.paragraphs):
                p = doc.paragraphs[i+1].insert_paragraph_before(texto_nuevo)
            else:
                p = doc.add_paragraph(texto_nuevo)
            
            p.alignment = 3  # Justificado
            style = p.style
            style.font.name = 'Arial'
            style.font.size = Pt(11)
            encontrado = True
            break
            
    if not encontrado:
        # Si no encuentra el título, lo avisa y lo pone al final
        st.warning(f" No encontré el título '{texto_titulo_buscar}' en la plantilla. Se agregó al final.")
        doc.add_paragraph(texto_nuevo)

def reemplazar_en_todo_el_doc(doc, diccionario_reemplazos):
    """
    Busca y reemplaza texto en párrafos y tablas.
    """
    # 1. Buscar en párrafos normales
    for paragraph in doc.paragraphs:
        for key, value in diccionario_reemplazos.items():
            if key in paragraph.text:
                # Usamos replace directo sobre el texto del párrafo
                # (Nota: esto borra formatos específicos dentro de la línea, pero es lo más seguro)
                paragraph.text = paragraph.text.replace(key, value)
    
    # 2. Buscar dentro de Tablas (Por si tu portada está maquetada con tablas)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in diccionario_reemplazos.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, value)


# 1. FUNCIONES (El cerebro)
# 1.1 Leer DM
def extraer_secciones_dm(archivo_word, mapa_claves):
    """archivo_word: El archivo subido por st.file_uploader. mapa_claves: Un diccionario que dice {'TITULO EN WORD': 'key_de_streamlit'}"""
    doc = Document(archivo_word)
    resultados = {}

# 1. Extraer todos los párrafos del documento
    todos_los_parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    
    # --- BUSCAR EL PUNTO DE PARTIDA ---
    indice_inicio_real = 0
    punto_partida = "BREVE RESEÑA HISTÓRICA DEL PROGRAMA"
    
    for i, texto in enumerate(todos_los_parrafos):
        if punto_partida in texto.upper():
            indice_inicio_real = i
            break # Encontramos el inicio real, dejamos de buscar
            
    # Creamos una nueva lista que solo contiene lo que hay desde la Reseña en adelante
    parrafos_validos = todos_los_parrafos[indice_inicio_real:]
    
    # --- PROCESO DE EXTRACCIÓN SOBRE LOS PÁRRAFOS VÁLIDOS ---
    for titulo_buscado, key_st in mapa_claves.items():
        contenido_seccion = []
        for i, texto in enumerate(parrafos_validos):
            texto_upper = texto.upper()
            
            # Buscamos el título (asegurándonos de que no sea una línea gigante)
            if titulo_buscado.upper() in texto_upper and len(texto) < 120:
                
                for j in range(i + 1, len(parrafos_validos)):
                    siguiente_p = parrafos_validos[j]
                    sig_upper = siguiente_p.upper()
                    
                   # Parar SOLO si encontramos un título principal (Ej: 3. o 4.)
                    # Bajamos el límite a 60 caracteres para no confundir párrafos con títulos
                    es_nuevo_capitulo = re.match(r'^\d+[\.\s]', siguiente_p.strip())
                    es_otro_titulo_mapa = any(t.upper() == sig_upper for t in mapa_claves.keys())

                    if (es_nuevo_capitulo or es_otro_titulo_mapa) and len(siguiente_p) < 60:
                        break
                        
                    contenido_seccion.append(siguiente_p)
                
                # 1. Guardamos TODO el texto en una variable "secreta" para el Word final
                texto_completo = "\n\n".join(contenido_seccion).strip()
                st.session_state[f"full_{key_st}"] = texto_completo
                
                # 2. Preparamos la VISTA PREVIA para el cuadro de texto
                parrafos_lista = texto_completo.split("\n\n")
                if len(parrafos_lista) > 2:
                    # Mostramos primer párrafo + aviso + último párrafo
                    resumen = f"{parrafos_lista[0]}\n\n[... {len(parrafos_lista)-2} PÁRRAFOS INTERMEDIOS CARGADOS TOTALMENTE ...]\n\n{parrafos_lista[-1]}"
                    resultados[key_st] = resumen
                else:
                    resultados[key_st] = texto_completo
                
                break

    #  PARTE 2: BUSCAR EN TABLAS
    for tabla in doc.tables:
        for fila in tabla.rows:
            # Verificamos que la fila tenga al menos 2 celdas
            if len(fila.cells) >= 2:
                texto_izq = fila.cells[0].text.strip().upper()
                texto_der = fila.cells[1].text.strip()
                
                # Comparamos la celda izquierda con nuestras palabras clave
                for titulo_buscado, key_st in mapa_claves.items():
                    if titulo_buscado.upper() in texto_izq:
                    # SIMPLIFICACIÓN: Guardamos el texto crudo del Word.
                    # La lógica de conversión la haremos en el widget (selectbox)
                        resultados[key_st] = texto_der

    return resultados

#1.2 Cargar BD
@st.cache_data # Esto hace que el Excel se lea una sola vez y no cada que muevas un botón
def cargar_base_datos():
    try:
        # Puedes usar pd.read_csv("programas.csv") si prefieres CSV
        df = pd.read_excel("Programas.xlsx", dtype={'snies_input': str}) 
        # Convertimos el DataFrame en un diccionario donde la llave es el SNIES
        return df.set_index("snies_input").to_dict('index')
    except Exception as e:
        st.warning(f"No se pudo cargar la base de datos de Excel: {e}")
        return {}

#1.3 Carga de datos inicial
BD_PROGRAMAS = cargar_base_datos()

#2. MAPEO Y ESTRUCTURA (DICCIONARIO)
# Mapeo de: "Título exacto en el DM" -> "Key en App Streamlit"
MAPA_EXTRACCION = {
    "OBJETO DE CONOCIMIENTO": "obj_nombre_input",
    "JUSTIFICACIÓN": "justificacion_input",
    "Conceptualización teórica y epistemológica del programa": "input_epi_p1",
    "Mecanismos de evaluación": "input_mec_p1",
    "IDENTIDAD DISCIPLINAR": "input_epi_p2",
    "ITINERARIO FORMATIVO": "input_itinerario",
    "Justificación del Programa": "input_just",
    "JUSTIFICACIÓN DEL PROGRAMA": "input_just"
    

}

#3. DICCIONARIO / ESTRUCTURA
# Agregamos 'key_dm' para que el extractor sepa qué título buscar en el Word
estructura_pep = {
    "1. Información del Programa": {
        "1.1. Historia del Programa": {"tipo": "especial_historia"},
        "1.2. Generalidades del Programa": {"tipo": "directo"}
    },
    "2. Referentes Conceptuales": {
        "2.1. Naturaleza del Programa": {
            "tipo": "directo",
            "key_dm": "OBJETO DE CONOCIMIENTO", # Palabra clave para buscar en el DM
            "campos": [
                {
                    "label": "Objeto de conocimiento del Programa", 
                    "req": True, 
                    "key": "obj_nombre_input",
                    "help": "¿Qué conoce, investiga y transforma este programa?"
                }
            ]
        },
        "2.2. Fundamentación epistemológica": {
            "tipo": "directo",
            "key_dm": "FUNDAMENTACIÓN EPISTEMOLÓGICA",
            "campos": [
                {"label": "Naturaleza epistemológica e identidad académica", "req": True, "key": "input_epi_p1"},
                {"label": "Campo del saber y relación con ciencia/tecnología", "req": True, "key": "input_epi_p2"}
            ]
        },
        "2.3. Fundamentación académica": {
            "tipo": "especial_pascual", 
            "campos": [] 
        }
    }
}


st.markdown("---")

#  CONFIGURACIÓN DE PÁGINA 
st.set_page_config(page_title="Generador Proyecto Educativo", layout="wide")
st.title("Generador PEP - Módulo 1: Información del Programa")
st.markdown("""
Esta herramienta permite generar el PEP de dos formas:
1. **Manual:** Completa los campos en las secciones de abajo.
2. **Automatizada:** Sube el Documento Maestro (DM) y el sistema pre-llenará algunos campos.
""")

   
# SELECTOR DE MODALIDAD
# Usamos un radio button estilizado para elegir el método
metodo_trabajo = st.radio(
    "Selecciona cómo deseas trabajar hoy:",
    ["Manual (Desde cero)", "Automatizado (Cargar Documento Maestro)"],
    horizontal=True,
    help="La opción automatizada intentará pre-llenar los campos usando un archivo Word."
)

# Botón DM
if metodo_trabajo == "Automatizado (Cargar Documento Maestro)":
    st.subheader("2. Carga de Documento Maestro")
    archivo_dm = st.file_uploader("Sube el archivo .docx del Documento Maestro", type=["docx"])
    
    if archivo_dm:
        # --- EL ESCÁNER (Usando tus Helpers para auditar) ---
        dict_maestro = docx_to_clean_dict(archivo_dm)
        with st.expander("🔍 Auditoría de Títulos (Jerarquía Detectada)"):
            if not dict_maestro:
                st.error("No se detectaron estilos de Título en el Word.")
            else:
                estructura_limpia = obtener_solo_estructura(dict_maestro)
                st.write("Jerarquía detectada (usa las flechas para expandir):")
                st.json(estructura_limpia)

        if st.button("Procesar y Pre-llenar desde Word"):
            with st.spinner("Extrayendo fundamentación..."):
            # Generamos el diccionario del maestro
                dict_maestro = docx_to_clean_dict(archivo_dm)
            
            # Título exacto que mencionas
            
            # Extraemos TODO (incluyendo subtítulos)
                contenido_extraido = buscar_contenido_por_titulo(dict_maestro, titulo_dm)
            
            if contenido_extraido:
                # Guardamos en la key que usas para el text_area
                st.session_state["fund_epi_manual"] = contenido_extraido.strip()
                st.success("✅ Fundamentación epistemológica extraída con subtítulos.")
            else:
                st.warning(f"⚠️ No se encontró la sección '{titulo_dm}'")

                # 3. Guardamos el resto de los datos en sus keys originales
                for key, valor in datos_capturados.items():
                    st.session_state[key] = valor             
                
                st.success("✅ Datos extraídos. Revisa el Capítulo 2.")
                st.rerun()




# LÓGICA DE MODALIDAD

with st.expander("Buscador Información general del Programa por SNIES", expanded=True):
    st.subheader("1. Búsqueda del Programa por SNIES")
    
    col_busq, col_btn = st.columns([3, 1])
    
    with col_busq:
        snies_a_buscar = st.text_input("Ingresa el código SNIES:", placeholder="Ej: 54862", key="search_snies_tmp")
        
    with col_btn:
        st.write(" ")
        st.write(" ")
        if st.button("🔍 Consultar Base de Datos"):
            if snies_a_buscar in BD_PROGRAMAS:
                datos_encontrados = BD_PROGRAMAS[snies_a_buscar]

                # 1. Borramos las llaves viejas para que el formulario no se bloquee
                llaves_a_limpiar = ["denom_input", "titulo_input", "snies_input", "acuerdo_input", "instancia_input", "reg1", "cred", "periodo_idx", "estudiantes_input", "acred1", "lugar"
]
                for k in llaves_a_limpiar:
                    if k in st.session_state:
                        del st.session_state[k]
                
                # 2. Inyectamos los nuevos datos del Excel
                for key, valor in datos_encontrados.items():
                    st.session_state[key] = valor
                
                # 3. Guardamos el SNIES que acabamos de buscar
                st.session_state["snies_input"] = snies_a_buscar
                
                st.success(f"✅ Programa encontrado: {datos_encontrados.get('denom_input')}")
                st.rerun()
            else:
                st.error("❌ Código SNIES no registrado en el sistema.")

    st.markdown("---")

# BOTÓN DE DATOS DE EJEMPLO
if st.button("Llenar con datos de ejemplo"):
    for k in ["denom_input", "titulo_input", "snies_input"]:
        if k in st.session_state:
            del st.session_state[k]
    st.session_state.ejemplo = {
        "denom_input": "Ingeniería de Sistemas",
        "titulo_input": "Ingeniero de Sistemas",
        "nivel_idx": 2, # Profesional universitario
        "area_input": "Ingeniería, Arquitectura y Urbanismo",
        "modalidad_input": 4, # Presencial y Virtual
        "acuerdo_input:": "Acuerdo 012 de 2015",
        "instancia_input": "Consejo Académico",
        "reg1": "Res. 4567 de 2016",
        "reg2": "Res. 8901 de 2023",
        "acred1": "Res. 00234 de 2024",
        "cred": "165",
        "estudiantes_input":"40",
        "periodo_idx": 0, # Semestral
        "lugar": "Sede Principal (Cali)",
        "snies": "54321",
        "motivo": "La creación del Programa se fundamenta en la necesidad de formar profesionales capaces de liderar la transformación digital, diseñar y desarrollar soluciones de software de alta complejidad, gestionar sistemas de información y responder de manera innovadora a los retos tecnológicos, organizacionales y sociales del entorno local, nacional e internacional.",
        "p1_nom": "EO1", "p1_fec": "Acuerdo 012-2015",
        "p2_nom": "EO2", "p2_fec": "Acuerdo 088-2020",
        "p3_nom": "EO3", "p3_fec": "Acuerdo 102-2024",
        #DATOS CAPÍTULO 2
        "objeto_nombre": "Sistemas de información",
        "objeto_concep": "Los sistemas de información son conjuntos organizados de personas, datos, procesos, tecnologías y recursos que interactúan de manera integrada para capturar, almacenar, procesar, analizar y distribuir información, con el fin de apoyar la toma de decisiones, la gestión operativa, el control organizacional y la generación de conocimiento. Estos sistemas permiten transformar los datos en información útil y oportuna, facilitando la eficiencia, la innovación y la competitividad en organizaciones de distintos sectores. Su diseño y gestión consideran aspectos técnicos, organizacionales y humanos, garantizando la calidad, seguridad, disponibilidad y uso ético de la información.",        
        "fund_epi": "El programa se inscribe en el racionalismo crítico y el pragmatismo tecnológico, vinculando la ciencia de la computación con la ingeniería aplicada.",
        # DATOS PARA LAS TABLAS (Se guardan como listas de diccionarios)
        "recon_data": [
            {"Año": "2024", "Nombre del premio": "Excelencia Académica", "Nombre del Ganador": "Juan Pérez", "Cargo": "Docente"}
        ],
        "tabla_cert_ej": [
            {"Nombre": "Desarrollador Web Junior", "Curso 1": "Programación I", "Créditos 1": 3, "Curso 2": "Bases de Datos", "Créditos 2": 4},
            {"Nombre": "Analista de Datos", "Curso 1": "Estadística", "Créditos 1": 4, "Curso 2": "Python para Ciencia", "Créditos 2": 4}
        ], #         
        "referencias_data": [
            {
                "Año": "2021", 
                "Autor(es)": "Sommerville, I.", 
                "Revista": "Computer science", 
                "Título del artículo/Libro": "Engineering Software Products"
            },
            {
                "Año": "2023", 
                "Autor(es)": "Pressman, R. & Maxim, B.", 
                "Revista": "Software Engineering Journal", 
                "Título del artículo/Libro": "A Practitioner's Approach"
            }
        ],
    }

    
    st.rerun()

# --- FORMULARIO DE ENTRADA ---
with st.form("pep_form"):
    # 1. Recuperamos datos de ejemplo si existen
    ej = st.session_state.get("ejemplo", {})

    st.markdown("### 1. Identificación General")
    col1, col2 = st.columns(2)
    
    with col1:
        # Denominación del programa
        denom = st.text_input(
            "Denominación del programa :red[•]", 
            value=st.session_state.get("denom_input", ej.get("denom_input", "")),
            key="denom_input"
        )

        # Título otorgado (Ahora bien indentado dentro de col1)
        titulo = st.text_input(
            "Título otorgado :red[•]", 
            value=st.session_state.get("titulo_input", ej.get("titulo_input", "")),
            key="titulo_input"
        )
    
        niveles_opciones = ["Técnico", "Tecnológico", "Profesional universitario", "Especialización", "Maestría", "Doctorado"]
        val_nivel = st.session_state.get("nivel_idx", st.session_state.get("ejemplo", {}).get("nivel_idx", 2))
        try:
            idx_final = int(val_nivel)
        except (ValueError, TypeError):
            idx_final = 2 # Por defecto Profesional
        
        nivel = st.selectbox(
            "Nivel de formación :red[•]", 
            options=niveles_opciones, 
            index=idx_final,
            key="nivel_formacion_widget"
        )
        # Código SNIES
        snies = st.text_input(
            "Código SNIES", 
            value=st.session_state.get("snies_input", ej.get("snies_input", "")),
            key="snies_input"
            )
        # 5. Número de Semestres 
        semestres = st.text_input(
            "Número de semestres (actuales) :red[•]",
            value=st.session_state.get("semestres_input", ej.get("semestres_input", "")),
            placeholder="Ej: 10",
            key="semestres_input"
        )
    
    with col2:
        idx_mod = st.session_state.get("modalidad_idx", 0)
        modalidad = st.selectbox(
            "Modalidad de oferta :red[•]", 
            ["Presencial", "Virtual", "A Distancia", "Dual", "Presencial y Virtual", "Presencial y a Distancia", "Presencial y Dual"],
            index=int(idx_mod) if isinstance(idx_mod, (int, float)) else 0,
            key="modalidad_input"
        )
        
        acuerdo = st.text_input(
            "Acuerdo de creación / Norma interna :red[•]", 
            key="acuerdo_input"
        )

        # Instancia interna
        instancia = st.text_input(
            "Instancia interna que aprueba :red[•]", 
            key="instancia_input"
        )

        # --- Fila 5: Periodicidad y Créditos ---
        col5_1, col5_2 = st.columns(2)
        
        with col5_1:
            periodicidad = st.selectbox(
                "Periodicidad de admisión :red[•]",
                ["Semestral", "Anual", "Trimestral", "Cuatrimestral"],
                index=0,
                key="periodicidad_input"
            )
    
        with col5_2:
            # --- TRUCO DE LIMPIEZA ---
            # Si "cred" ya existe en session_state y no es texto, lo convertimos a la fuerza
            if "cred" in st.session_state and not isinstance(st.session_state["cred"], str):
                st.session_state["cred"] = str(st.session_state["cred"])
            
            # Ahora sí, extraemos el valor inicial con seguridad
            valor_inicial_creditos = str(st.session_state.get("cred", ej.get("cred", "")))
            
            creditos = st.text_input(
                "Créditos académicos (actuales) :red[•]",
                value=valor_inicial_creditos,
                placeholder="Ej: 160",
                key="cred"
            )
    
        # --- Fila 6: Lugar y Estudiantes ---
        col6_1, col6_2 = st.columns(2)
        
        with col6_1:
            lugar = st.text_input(
                "Lugar de desarrollo :red[•]",
                value=st.session_state.get("lugar_input", ej.get("lugar_input", "Medellín - Campus Robledo")),
                key="lugar_input"
            )
    
        with col6_2:
            # --- PROTECCIÓN CONTRA TYPEERROR ---
            # Si el valor en session_state no es string, lo convertimos ahora mismo
            if "estudiantes_input" in st.session_state and not isinstance(st.session_state["estudiantes_input"], str):
                st.session_state["estudiantes_input"] = str(st.session_state["estudiantes_input"])
            
            # Aseguramos que el valor inicial sea string también desde el diccionario 'ej'
            valor_estudiantes = str(st.session_state.get("estudiantes_input", ej.get("estudiantes_input", "")))
            
            estudiantes_primer = st.text_input(
                "Número de estudiantes en primer periodo :red[•]",
                value=valor_estudiantes,
                placeholder="Ej: 40",
                key="estudiantes_input"
            )

    st.markdown("---")
    st.markdown("### 2. Registros y Acreditaciones")
    def forzar_texto(key, fuente):
        # 1. Recuperamos el valor (de la sesión o del ejemplo)
        valor = st.session_state.get(key, fuente.get(key, ""))
        
        # 2. Si es None, lo convertimos a vacío
        if valor is None:
            valor = ""
        
        # 3. Lo convertimos a String (texto) sí o sí, y actualizamos la sesión
        # Esto sobreescribe cualquier "basura" (números o nulos) que haya quedado en memoria
        st.session_state[key] = str(valor)
   
    with st.container(border=True):
        col_reg, col_acred = st.columns(2)

        with col_reg:
            st.markdown("#### **Registros Calificados**")
                              
            forzar_texto("reg1", ej)
            st.text_input(
                "Resolución Registro Calificado 1 :red[•]", 
                placeholder="Ej: Resolución 12345 de 2023",
                key="reg1"
            )
            
            # --- REGISTRO 2 ---
            forzar_texto("reg2", ej)
            st.text_input(
                "Resolución Registro Calificado 2", 
                placeholder="Ej: Resolución 67890 de 2023",
                key="reg2"
            )

            # --- REGISTRO 3 ---
            forzar_texto("reg3", ej)
            st.text_input(
                "Resolución Registro Calificado 3", 
                placeholder="Dejar vacío si no aplica",
                key="reg3"
            )
            
        with col_acred:
            st.markdown("#### **Acreditaciones**")
            
            # --- ACREDITACIÓN 1 ---
            forzar_texto("acred1", ej)
            st.text_input(
                "Resolución Acreditación Alta Calidad 1", 
                placeholder="Ej: Resolución 012345 de 2022",
                key="acred1"
            )
            
            # --- ACREDITACIÓN 2 ---
            forzar_texto("acred2", ej)
            st.text_input(
                "Resolución Acreditación Alta Calidad 2", 
                placeholder="Dejar vacío si no aplica",
                key="acred2"
            )
    
    frase_auto = f"La creación del Programa {denom} se fundamenta en la necesidad de "
    val_motivo = ej.get("motivo", frase_auto)
    motivo = st.text_area("Motivo de creación :red[•]", value=val_motivo, height=150)
      
    st.markdown("---")
    st.markdown("### 3. Modificaciones al Plan de Estudios")
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        p1_nom = st.text_input("Nombre Plan v1:red[•]", value=ej.get("p1_nom", ""), key="p1_nom")
        p1_fec = st.text_input("Acuerdo aprobación Plan v1 :red[•]", value=ej.get("p1_fec", ""), key="p1_fec")
        p1_cred = st.text_input("Número de créditos Plan v1 :red[•]", value=ej.get("p1_cred", ""), key="p1_cred")
        p1_sem = st.text_input("Número de semestres Plan v1:red[•]", value=ej.get("p1_sem", ""), key="p1_sem")
    with p_col2:
        p2_nom = st.text_input("Nombre Plan v2 (Opcional)", value=ej.get("p2_nom", ""), key="p2_nom")
        p2_fec = st.text_input("Acuerdo aprobación Plan v2 (Opcional)", value=ej.get("p2_fec", ""), key="p2_fec")
        p2_cred = st.text_input("Número de créditos Plan v2 (Opcional) :red[•]", value=ej.get("p2_cred", ""),key="p2_cred")
        p2_sem = st.text_input("Número de semestres Plan v2 (Opcional):red[•]",value=ej.get("p2_sem", ""),key="p2_sem")
    with p_col3:
        p3_nom = st.text_input("Nombre Plan v3 (Opcional)", value=ej.get("p3_nom", ""), key="p3_nom")
        p3_fec = st.text_input("Acuerdo aprobación Plan v3 (Opcional)", value=ej.get("p3_fec", ""), key="p3_fec")
        p3_cred = st.text_input("Número de créditos Plan v3 (Opcional)", value=ej.get("p3_cred", ""),key="p3_cred")
        p3_sem = st.text_input("Número de semestresPlan v3(Opcional)", value=ej.get("p3_sem", ""), key="p3_sem")

    st.markdown("---")
    st.markdown("### 🏆 4. Reconocimientos (Opcional)")
    recon_data = st.data_editor(
        ej.get("recon_data", [{"Año": "", "Nombre del premio": "", "Nombre del Ganador": "", "Cargo": "Estudiante"}]),
        num_rows="dynamic",
        key="editor_recon", # Es vital tener una key única
        column_config={
            "Cargo": st.column_config.SelectboxColumn(options=["Docente", "Líder", "Decano", "Estudiante","Docente Investigador", "Investigador"])
        },
        use_container_width=True
    )
    st.session_state["recon_data"] = recon_data
    
#CAPÍTULO 2
    st.markdown("---")
    st.header("2. Referentes Conceptuales")
   # 2. MODO MANUAL Objeto de conocimiento del Programa
    val_obj_nombre = ej.get("objeto_nombre", "")
    
    objeto_nombre = st.text_input(
        "1. ¿Cuál es el Objeto de conocimiento del Programa? :red[•]",
        value=st.session_state.get("obj_nombre_input", val_obj_nombre),
        placeholder="Ejemplo: Sistemas de información",
        key="obj_nombre_input"  #
    )
    # 2. Definición del Objeto (Lo que llenará {{def_oc}})
    st.write("---")
    st.write("**Definición del Objeto de Conocimiento**")
    
    # Selector de método (Asumiendo que tienes una variable global o local 'metodo_trabajo')
    if metodo_trabajo == "Manual":
        # Si es manual, el usuario escribe directamente la definición
        st.text_area(
            "Escriba la definición del Objeto de Conocimiento:",
            value=st.session_state.get("def_oc_manual", ""),
            placeholder="Ingrese el texto aquí...",
            key="def_oc_manual",
            height=200
        )
    else:
        # Si es Automatizado, pedimos los marcadores para buscar en el Word Maestro
        st.info("Configuración de Extracción: Indique dónde inicia y termina la definición en el Documento Maestro.")
        col_ini, col_fin = st.columns(2)
        
        with col_ini:
            st.text_input(
                "Texto de inicio:",
                placeholder="Ej: El objeto de estudio se define...",
                key="inicio_def_oc"
            )
        with col_fin:
            st.text_input(
                "Texto final:",
                placeholder="Ej: ...en el contexto regional.",
                key="fin_def_oc"
            )
    

                    
    # 3. REFERENCIAS (Esto sigue igual para ambos casos)
    st.write(" ")
    st.write("Referencias bibliográficas que sustentan la conceptualización del Objeto de Conocimiento.")
    referencias_previa = ej.get("referencias_data", [
        {"Año": "", "Autor(es) separados por coma": "", "Revista": "", "Título del artículo/Libro": ""}
    ])
    
    referencias_data = st.data_editor(
        referencias_previa,
        num_rows="dynamic",
        key="editor_referencias",
        use_container_width=True,
        column_config={
            "Año": st.column_config.TextColumn("Año", width="small"),
            "Autor(es)": st.column_config.TextColumn("Autor(es)", width="medium"),
            "Revista": st.column_config.TextColumn("Revista", width="medium"),
            "Título del artículo/Libro": st.column_config.TextColumn("Título del artículo/Libro", width="large"),
        }
    )

  # 2.2. Fundamentación epistemológica en Pestañas ---
    st.markdown("---")
    st.subheader("2.2. Fundamentación epistemológica")
    if metodo_trabajo != "Automatizado (Cargar Documento Maestro)":
        
        # ==========================================
        # CASO 1: MODO MANUAL (Aquí SÍ creamos pestañas)
        # ==========================================
        st.info("Utilice las pestañas para completar los tres párrafos de la Fundamentación epistemológica.")
        
        # --- AQUÍ LA CLAVE: Creamos las tabs SOLO si es manual ---
        tab1, tab2, tab3 = st.tabs(["Párrafo 1", "Párrafo 2", "Párrafo 3"])

        # Configuración de columnas para referencias
        config_columnas_ref = {
            "Año": st.column_config.TextColumn("Año", width="small"),
            "Autor(es) separados por coma": st.column_config.TextColumn("Autor(es)", width="medium"),
            "Revista": st.column_config.TextColumn("Revista", width="medium"),
            "Título del artículo/Libro": st.column_config.TextColumn("Título del artículo/Libro", width="large"),
        }

        # Bloque Párrafo 1
        with tab1:
            st.markdown("### Párrafo 1: Marco filósofico")
            st.text_area(
                "¿Cuál es la postura filosófica predominante? :red[•]",
                value=ej.get("fund_epi_p1", ""), 
                height=200,
                key="input_epi_p1",
                placeholder="Ejemplo: El programa se fundamenta en el paradigma de la complejidad..."
            )
            st.write("Referencias bibliográficas (Párrafo 1):")
            st.data_editor(
                ej.get("referencias_epi_p1", [{"Año": "", "Autor(es) separados por coma": "", "Revista": "", "Título del artículo/Libro": ""}]),
                num_rows="dynamic",
                key="editor_refs_p1",
                use_container_width=True,
                column_config=config_columnas_ref
            )

        # Bloque Párrafo 2
        with tab2:
            st.markdown("### Párrafo 2: Identidad disciplinar")
            st.text_area(
                "Origen etimológico y teorías conceptuales :red[•]",
                value=ej.get("fund_epi_p2", ""), 
                height=200,
                key="input_epi_p2",
                placeholder="Ejemplo: La identidad de este programa se define desde..."
            )
            st.write("Referencias bibliográficas (Párrafo 2):")
            st.data_editor(
                ej.get("referencias_epi_p2", [{"Año": "", "Autor(es) separados por coma": "", "Revista": "", "Título del artículo/Libro": ""}]),
                num_rows="dynamic",
                key="editor_refs_p2",
                use_container_width=True,
                column_config=config_columnas_ref
            )

        # Bloque Párrafo 3
        with tab3:
            st.markdown("### Párrafo 3: Intencionalidad social")
            st.text_area(
                "¿Intervención ética y transformadora? :red[•]",
                value=ej.get("fund_epi_p3", ""), 
                height=200,
                key="input_epi_p3",
                placeholder="Ejemplo: Finalmente, la producción de conocimiento..."
            )
            st.write("Referencias bibliográficas (Párrafo 3):")
            st.data_editor(
               ej.get("referencias_epi_p3", [{"Año": "", "Autor(es) separados por coma": "", "Revista": "", "Título del artículo/Libro": ""}]),
                num_rows="dynamic",
                key="editor_refs_p3",
                use_container_width=True,
                column_config=config_columnas_ref
            )

    else:
        
        # CASO 2: MODO AUTOMATIZADO (SIN pestañas)
        st.success("✅ Modo Estructurado: El sistema extraerá automáticamente el contenido de la sección 'Conceptualización teórica y epistemológica del programa' desde el Documento Maestro.")
        # No hay col_ini ni col_fin aquí
    
        #st.info("Configuración de Extracción:  Indique dónde inicia y termina la Conceputalización Teórica y Epistemológica en el Documento Maestro. Fundamentación Epistemológica")
        
        # Aquí NO usamos st.tabs, usamos columnas directas
        #with st.container(border=True):
         #   col_inicio, col_fin = st.columns(2)
            
          #  with col_inicio:
           #     st.text_input(
            #        "Texto de inicio :red[•]", 
             #       placeholder="Ej: 2.2 Fundamentación Epistemológica",
              #      help="Copia y pega las primeras palabras del capítulo en el Word.",
               #     key="txt_inicio_fund_epi"
                #)
            
            #with col_fin:
             #   st.text_input(
              #      "Texto final :red[•]", 
               #     placeholder="Ej: 2.3 Justificación",
                #    help="Copia y pega las primeras palabras del SIGUIENTE capítulo o donde termina este.",
                 #   key="txt_fin_fund_epi"
                #)

  # --- 2.3. Fundamentación Académica ---
    st.markdown("---")
    st.subheader("2.3. Fundamentación Académica")
    
    # ---------------------------------------------------------
    # 2.3.1 MICROCREDENCIALES (Siempre visible)
    # ---------------------------------------------------------
    st.write("***2.3.1. Microcredenciales***")
    st.info("Agregue filas según sea necesario para listar las microcredenciales.")
    
    datos_micro = ej.get("tabla_micro", [
        {"Nombre de la Certificación": "", "Nombre del Curso": "", "Créditos": 0}
    ])
    
    st.data_editor(
        datos_micro,
        num_rows="dynamic", 
        key="editor_microcredenciales",
        use_container_width=True,
        column_config={
            "Nombre de la Certificación": st.column_config.TextColumn("Certificación", width="medium"),
            "Nombre del Curso": st.column_config.TextColumn("Curso Asociado", width="medium"),
            "Créditos": st.column_config.NumberColumn("Créditos", min_value=0, step=1, width="small")
        }
    )

    st.write(" ") 

    # ---------------------------------------------------------
    # 2.3.2 MACROCREDENCIALES (Siempre visible)
    # ---------------------------------------------------------
    st.write("***2.3.2. Macrocredenciales***")
    st.info("Cada fila representa una Certificación (Macrocredencial). Complete los cursos que la componen (máx 3).")

    datos_macro = ej.get("tabla_macro", [
        {
            "Certificación": "", 
            "Curso 1": "", "Créditos 1": 0,
            "Curso 2": "", "Créditos 2": 0,
            "Curso 3": "", "Créditos 3": 0
        }
    ])

    columnas_config = {
        "Certificación": st.column_config.TextColumn(
            "Nombre Macrocredencial", 
            width="medium",
            help="Nombre de la certificación global (ej: Diplomado en Big Data)",
            required=True
        ),
        "Curso 1": st.column_config.TextColumn("Curso 1", width="medium"),
        "Créditos 1": st.column_config.NumberColumn("Créd. 1", width="small", min_value=0, step=1),
        "Curso 2": st.column_config.TextColumn("Curso 2", width="medium"),
        "Créditos 2": st.column_config.NumberColumn("Créd. 2", width="small", min_value=0, step=1),
        "Curso 3": st.column_config.TextColumn("Curso 3", width="medium"),
        "Créditos 3": st.column_config.NumberColumn("Créd. 3", width="small", min_value=0, step=1),
    }

    st.data_editor(
        datos_macro,
        num_rows="dynamic", 
        key="editor_macrocredenciales",
        use_container_width=True,
        column_config=columnas_config
    )
       
    # ---------------------------------------------------------
    # 2.3.3 ÁREAS DE FORMACIÓN (Condicional)
    # ---------------------------------------------------------
    st.write("") 
    st.write("**2.3.3. Áreas de formación**")
    
    # CASO MANUAL
    if metodo_trabajo != "Automatizado (Cargar Documento Maestro)":
        area_especifica = st.text_area(
            "Descripción del Área de Fundamentación Específica :red[•]",
            value=ej.get("fund_especifica_desc", ""),
            height=150,
            placeholder="Desarrolla competencias técnicas y profesionales específicas del programa...",
            key="input_area_especifica"
        )
    # CASO AUTOMATIZADO
    else:
        st.info("Configuración: Defina el párrafo de descripción del Área Específica.")
        with st.container(border=True):
            c1, c2 = st.columns(2)
            c1.text_input("Inicio Descripción Área:", placeholder="Ej: El área específica...", key="ini_area_esp")
            c2.text_input("Fin Descripción Área:", placeholder="Ej: ...ejercicio profesional.", key="fin_area_esp")

    # ---------------------------------------------------------
    # 2.3.4 CURSOS POR ÁREA (Solo configuración Automatizada)
    # ---------------------------------------------------------
    st.write("***2.3.4. Cursos por área de formación***")
    
    # Lista de áreas en el orden solicitado
    areas_formacion = [
        "Formación Humanística",
        "Fundamentación Básica",
        "Formación Básica Profesional",
        "Fundamentación Específica del Programa",
        "Formación Flexible o Complementaria"
    ]

    # CASO AUTOMATIZADO
    if metodo_trabajo == "Automatizado (Cargar Documento Maestro)":
        st.info("Configuración de Extracción: Configure las tablas de cursos para cada área. Deje vacías las que no apliquen.")
            
        # Generamos los bloques de configuración basados en la lista anterior
        for area in areas_formacion:
            # Creamos un ID único reemplazando espacios por guiones bajos
            area_id = area.lower().replace(" ", "_")
            
            with st.expander(f"Tabla: {area}", expanded=False):
                st.markdown(f"**Configuración para {area}**")
                
                # Fila para los marcadores de extracción
                col_tabla_inicio, col_tabla_fin = st.columns(2)
                
                with col_tabla_inicio:
                    st.text_input(
                        f"Texto Inicio Tabla :red[•]", 
                        placeholder=f"Ej: Tabla de cursos {area}",
                        help=f"Copia el título exacto de la tabla de {area} en el Word.",
                        key=f"txt_inicio_{area_id}"
                    )
                    
                with col_tabla_fin:
                    st.text_input(
                        f"Texto Fin Tabla :red[•]", 
                        value="Fuente: Elaboración propia", 
                        help="Texto donde termina la tabla (usualmente la fuente).",
                        key=f"txt_fin_{area_id}"
                    )

    # CASO MANUAL
    else:
        st.info("En el documento final, asegúrese de incluir las tablas de cursos organizadas por:")
        for area in areas_formacion:
            st.write(f"- {area}")

 # Itinerario formativo
    st.write("") 
    st.subheader("3.Itinerario formativo")
    
    area_especifica = st.text_area("Teniendo como fundamento que, en torno a un objeto de conocimiento se pueden estructurar varios programas a diferentes niveles de complejidad, es importante expresar si el programa en la actualidad es único en torno al objeto de conocimiento al que está adscrito o hay otros de mayor o de menor complejidad.:red[•]",
        value=ej.get("fund_especifica_desc", ""),
        height=150,
        placeholder=" Ejemplo si el PEP es de Ingeniería Mecánica, determinar si hay otro programa de menor complejidad como una tecnología Mecánica o uno de mayor complejidad como una especialización o una maestría. Este itinerario debe considerar posibles programas de la misma naturaleza que se puedan desarrollar en el futuro.",
        key="input_itinerario"
    )

     # Justificación del Programa
    st.write("") 
    st.subheader("4.Justificación del Programa")
    
    # CONDICIONAL: Manual vs Automatizado
    if metodo_trabajo != "Automatizado (Cargar Documento Maestro)":
        
        # ==========================================
        # CASO 1: MODO MANUAL
        # ==========================================
        st.write("**Redacción Manual de la Justificación**")
        st.text_area(
            "Demostrar la relevancia del programa en el contexto actual, resaltando su impacto en la solución de problemáticas sociales y productivas. Se debe enfatizar cómo la formación impartida contribuye al desarrollo del entorno local, regional y global, alineándose con las necesidades del sector productivo, las políticas educativas y las tendencias del mercado laboral. :red[•]",
            value=ej.get("justificacion_desc", ""), # Cambiado a una llave más descriptiva
            height=250,
            placeholder="Fundamentar la relevancia del programa con datos actualizados, referencias normativas y estudios sectoriales. Evidenciar su alineación con los Objetivos de Desarrollo Sostenible (ODS), planes de desarrollo nacionales y políticas de educación superior. Incorporar análisis de tendencias internacionales que justifiquen su pertinencia en un contexto globalizado.",
            key="input_just_manual"
        )

    else:
        # ==========================================
        # CASO 2: MODO AUTOMATIZADO
        # ==========================================
        st.info("Configuración de Extracción: Justificación del Programa")
        
        with st.container(border=True):
            col_just_inicio, col_just_fin = st.columns(2)
            
            with col_just_inicio:
                st.text_input(
                    "Texto de inicio :red[•]", 
                    placeholder="Ej: 2.4 Justificación",
                    help="Copia y pega las primeras palabras donde inicia la justificación en el Word.",
                    key="txt_inicio_just"
                )
            
            with col_just_fin:
                st.text_input(
                    "Texto final :red[•]", 
                    placeholder="Ej: 2.5 Objetivos",
                    help="Copia y pega el inicio del siguiente capítulo para marcar el final de la extracción.",
                    key="txt_fin_just"
                )
    # --- SECCIÓN 5: ESTRUCTURA CURRICULAR ---
    st.markdown("---")
    st.header("5. Estructura Curricular")
    
    st.info("Defina el objeto de conocimiento y relacione las perspectivas de intervención con sus respectivas competencias.")

    # 1. Sector social y/o productivo
    with st.container(border=True):
        st.write("***Sector Social y/o Productivo***")
        st.text_area(
            " Sector Social y/o Productivo en el que interviene el Programa:red[•]",
            placeholder="Ejemplo: Sector manufactura...",
            key="sector",
            height=50
        )

    st.write("") # Espacio
    st.write("***Perspectivas de Intervención y Competencias***")
    st.markdown("Complete los cuadros paralelos a continuación:")

    # 2. Generación de los 6 Cuadros Paralelos
    for i in range(1, 7):
        with st.container(border=True):
            st.markdown(f"**Relación de Desempeño #{i}**")
            col_izq, col_der = st.columns(2)
            
            with col_izq:
                st.text_area(
                    f"Objeto de Formación / Perspectiva de intervención {i}",
                    placeholder=f"Defina la perspectiva {i}...",
                    key=f"objeto_formacion_{i}",
                    height=100
                )
                
            with col_der:
                st.text_area(
                    f"Competencia de Desempeño Profesional {i}",
                    placeholder=f"Defina la competencia {i}...",
                    key=f"competencia_desempeno_{i}",
                    height=100
                )

    # Nota al pie para el usuario
    st.caption("Nota: No es obligatorio llenar los 6 campos. El sistema procesará solo aquellos que contengan información.")

    # --- 2.5. Pertinencia Académica ---
    st.markdown("---")
    st.write("***5.2. Pertinencia Académica****")

    if metodo_trabajo == "Automatizado (Cargar Documento Maestro)":
        st.info("Configuración de Extracción: Tabla de Pertinencia Académica")
        
        with st.container(border=True):
            col_pert_inicio, col_pert_fin = st.columns(2)
            
            with col_pert_inicio:
                st.text_input(
                    "Nombre exacto de la Tabla de Pertinencia :red[•]", 
                    placeholder="Ej: Tabla 10. Pertinencia académica del programa",
                    help="Copia y pega el título de la tabla tal como aparece en el Word maestro.",
                    key="txt_inicio_tabla_pertinencia"
                )
            
            with col_pert_fin:
                st.text_input(
                    "Texto final de corte (Fin) :red[•]", 
                    value="Fuente: Elaboración propia", 
                    help="El sistema dejará de copiar cuando encuentre este texto debajo de la tabla.",
                    key="txt_fin_tabla_pertinencia"
                )
    else:
        # Modo Manual
        st.info("En el modo manual, redacte la pertinencia académica directamente en su documento final o cargue la tabla correspondiente.")
        st.text_area(
            "Descripción de la Pertinencia Académica (Opcional)",
            placeholder="Describa cómo el programa se alinea con las tendencias académicas actuales...",
            key="input_pertinencia_manual",
            height=150
        )

    # --- 5.3. Plan de Estudios ---
    st.markdown("---")
    st.write("***5.3. Plan de Estudios***")
    
    st.info("Por favor, cargue la imagen del Plan de Estudios (Malla Curricular) para ser incluida en el documento.")

    # Contenedor de carga de archivo
    with st.container(border=True):
        archivo_imagen = st.file_uploader(
            "Seleccionar imagen del Plan de Estudios :red[•]",
            type=["png", "jpg", "jpeg"],
            help="Soporta formatos PNG, JPG y JPEG. Esta imagen se insertará en la sección 5.3 del Word.",
            key="upload_plan_estudios"
        )

        # Si el usuario sube un archivo, mostrar una vista previa
        if archivo_imagen is not None:
            st.success("✅ Imagen cargada correctamente.")
            
            # Mostramos una vista previa pequeña/mediana
            st.image(
                archivo_imagen, 
                caption="Vista previa del Plan de Estudios cargado", 
                use_container_width=True
            )
            
            # Opción para que el usuario añada un título o fuente a la imagen
            st.text_input(
                "Título/Nota de la imagen (Opcional):",
                value="Gráfico: Plan de Estudios del Programa",
                key="caption_plan_estudios"
            )
        else:
            st.warning("⚠️ No se ha cargado ninguna imagen aún.")

     # --- 5.4 PERFILES ---
    st.markdown("---")
    st.header("5.4. Perfiles")
    
    st.info("Defina los perfiles que caracterizan al programa")

    # Usamos st.container para agrupar visualmente la fila de perfiles
    with st.container(border=True):
        # Creamos tres columnas de igual ancho
        col_prof, col_egr, col_ocup = st.columns(3)
        
        with col_prof:
            st.markdown("### **Perfil Profesional con Experiencia.**")
            #st.caption("(Con Experiencia)")
            st.text_area(
                "Defina el perfil del profesional con experiencia :red[•]",
                placeholder="Describa las capacidades y trayectoria que se esperan del profesional...",
                key="perfil_profesional_exp",
                height=300
            )
            
        with col_egr:
            st.markdown("### **Perfil Profesional del Egresado.**")
            #st.caption("(Al finalizar el programa)")
            st.text_area(
                "Defina el perfil profesional del egresado :red[•]",
                placeholder="Describa las competencias y conocimientos con los que sale el estudiante...",
                key="perfil_profesional_egresado",
                height=300
            )
            
        with col_ocup:
            st.markdown("### **Perfil Ocupacional.**")
            #st.caption("(Campo de acción)")
            st.text_area(
                "Defina el perfil ocupacional :red[•]",
                placeholder="Mencione los cargos, sectores y áreas donde podrá desempeñarse...",
                key="perfil_ocupacional",
                height=300
            )

    # Nota de ayuda para la redacción
    with st.expander("💡 Tips para redactar los perfiles"):
        st.markdown("""
        * **Profesional con experiencia:** SDeclaración que hace el programa académico acerca del resultado esperado de la formación para toda la vida.
        * **Egresado:**  Promesa de valor que la institución hace a los estudiantes y a la sociedad en general.
        * **Ocupacional:** Conjunto de conocimientos, habilidades, destrezas y actitudes que desarrollará el futuro profesional de un programa académico y que le permitirán desempeñarse laboralmente.
        """)
    # --- 7. RECURSOS ACADÉMICOS ---
    st.markdown("---")
    st.header("7. Recursos Académicos")
    
    # 7.1 Entornos académicos
    st.subheader("7.1. Entornos académicos")
    
    st.info("""
        Describa los espacios físicos y virtuales que soportan el programa. 
        Incluya laboratorios, bases de datos, plataformas de aprendizaje (LMS), 
        aulas especializadas y software técnico.
    """)

    with st.container(border=True):
        entornos_desc = st.text_area(
            "Detalle de Entornos Académicos (Físicos y Virtuales) :red[•]",
            value=ej.get("entornos_academicos_desc", ""),
            height=250,
            placeholder="""Ejemplo: El programa cuenta con acceso a laboratorios de última generación equipados con... 
Así mismo, se dispone de la plataforma Canvas para el aprendizaje virtual, acceso a la biblioteca digital con bases de datos como IEEE, Scopus... 
Se hace uso de software especializado como (nombre del software) para las prácticas de...""",
            key="input_entornos_academicos"
        )
        
    # Opcional: Si deseas que puedan listar recursos específicos en una tabla dinámica
    with st.expander("Añadir listado técnico de software o laboratorios (Opcional)"):
        st.write("Si el programa requiere software o equipos específicos, lístelos aquí:")
        
        datos_recursos = ej.get("tabla_recursos_tecnicos", [
            {"Recurso": "", "Tipo": "Software", "Descripción/Uso": ""}
        ])
        
        st.data_editor(
            datos_recursos,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_recursos_tecnicos",
            column_config={
                "Tipo": st.column_config.SelectboxColumn(
                    "Tipo",
                    options=["Software", "Hardware", "Laboratorio", "Base de Datos", "Otro"],
                    required=True
                )
            }
        )


    # --- 7.2. TALENTO HUMANO ---
    st.write("") 
    st.subheader("7.2. Talento Humano")
    
    st.info("""
        Describa el perfil del equipo docente requerido (formación académica, 
        experiencia profesional e investigativa) para garantizar el desarrollo 
        de las funciones de docencia, investigación y extensión del Programa.
    """)

    with st.container(border=True):
        talento_humano_desc = st.text_area(
            "Perfil del equipo docente requerido :red[•]",
            value=ej.get("talento_humano_desc", ""),
            height=250,
            placeholder="""Ejemplo: El programa requiere un equipo docente con formación de posgrado a nivel de Maestría y/o Doctorado en áreas afines a... 
Se valorará la experiencia profesional en el sector de... así como la participación en grupos de investigación categorizados por MinCiencias. 
El equipo debe demostrar competencias pedagógicas para el manejo de entornos virtuales...""",
            key="input_talento_humano"
        )
    
    # Ayuda adicional para el usuario
    with st.expander("💡 ¿Qué debe incluir este perfil?"):
        st.markdown("""
        Al redactar el perfil del talento humano, considere mencionar:
        * **Nivel de formación:** (Especialistas, Magísteres, Doctores).
        * **Experiencia profesional:** Años de trayectoria en el sector productivo.
        * **Capacidades investigativas:** Producción académica o pertenencia a grupos de investigación.
        * **Competencias blandas/pedagógicas:** Capacidad de innovación educativa y uso de TIC.
        """)
    # --- 8. INVESTIGACIÓN, TECNOLOGÍA E INNOVACIÓN ---
    st.markdown("---")
    st.header("8. Investigación, Tecnología e Innovación")
    
    st.info("""
        **Indicaciones:** Describa la organización de la investigación en el programa. 
        Especifique las líneas y grupos de investigación , destacando 
        objetivos y su articulación con el proceso formativo.
    """)

    with st.container(border=True):
        # 1. Descripción General y Grupos
        st.subheader("Estructura de Investigación")
        investigacion_desc = st.text_area(
            "Descripción de Grupos y Líneas de Investigación :red[•]",
            value=ej.get("investigacion_desc", ""),
            height=250,
            placeholder="""Ejemplo: La investigación en el programa se articula a través del Grupo de Investigación (Nombre), categorizado en (A, B, C) por MinCiencias. 
Sus líneas de acción incluyen: 
1. (Línea 1)
2. (Línea 2)
Estas líneas permiten que el estudiante participe activamente en...""",
            key="input_investigacion_general"
        )
    # --- 9. VINCULACIÓN NACIONAL E INTERNACIONAL ---
    st.markdown("---")
    st.header("9. Vinculación Nacional e Internacional")
    
    # 9.1 Estrategias de internacionalización
    st.subheader("9.1. Estrategias de internacionalización")
    
    st.info("""
        **Indicaciones:** Describa las acciones que permiten la visibilidad nacional e internacional del programa. 
        Incluya estrategias como: movilidad académica (estudiantes y docentes), convenios de doble titulación, 
        participación en redes académicas, internacionalización del currículo (COIL, invitados internacionales) 
        y bilingüismo.
    """)

    with st.container(border=True):
        internacionalizacion_desc = st.text_area(
            "Descripción de estrategias de internacionalización :red[•]",
            value=ej.get("internacionalizacion_desc", ""),
            height=300,
            placeholder="""Ejemplo: El programa fomenta la internacionalización a través de convenios marco con universidades de España y México para movilidad estudiantil. 
Se implementa la metodología COIL en las asignaturas de... 
Además, el programa participa activamente en la red (Nombre de la Red) y promueve el bilingüismo mediante el uso de recursos bibliográficos en segunda lengua...""",
            key="input_internacionalizacion"
        )

    # Tabla complementaria opcional para convenios específicos
    with st.expander("📋 Listado de Convenios y Aliados (Opcional)"):
        st.write("Si desea tabular los convenios vigentes, lístelos aquí:")
        datos_convenios = ej.get("tabla_convenios", [
            {"Institución/Aliado": "", "País": "Colombia", "Tipo de Alianza": "Movilidad"}
        ])
        
        st.data_editor(
            datos_convenios,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_convenios",
            column_config={
                "Tipo de Alianza": st.column_config.SelectboxColumn(
                    "Tipo de Alianza",
                    options=["Movilidad Académica", "Doble Titulación", "Investigación Conjunta", "Prácticas Profesionales", "Otro"],
                    required=True
                )
            }
        )

    # --- 10. BIENESTAR UNIVERSITARIO ---
    st.markdown("---")
    st.header("10. Bienestar en el Programa")
    
    st.info("""
        **Indicaciones:** Describa las acciones, programas y servicios de bienestar que 
        impactan directamente a los estudiantes y docentes del programa. 
        Enfoque su respuesta en la **permanencia académica**, el desarrollo humano, 
        la salud, el deporte, la cultura y los apoyos socioeconómicos.
    """)

    with st.container(border=True):
        bienestar_desc = st.text_area(
            "Descripción de estrategias de Bienestar y Permanencia :red[•]",
            value=ej.get("bienestar_desc", ""),
            height=300,
            placeholder="""Ejemplo: El programa se articula con la Política de Bienestar Institucional a través de estrategias de acompañamiento docente (tutorías) para mitigar el riesgo de deserción... 
Se cuenta con programas de apoyo psicosocial, becas socioeconómicas y fomento de la cultura y el deporte. 
Asimismo, se realizan jornadas de integración y seguimiento integral al estudiante desde su ingreso hasta su graduación...""",
            key="input_bienestar"
        )

    # Tabla opcional para programas de apoyo específicos
    with st.expander("📋 Programas Específicos de Apoyo (Opcional)"):
        st.write("Si el programa cuenta con apoyos específicos (ej: tutorías especializadas, bonos, convenios), lístelos aquí:")
        datos_apoyo = [
            {"Programa/Estrategia": "Tutorías Académicas", "Objetivo": "Reducir la pérdida académica"},
            {"Programa/Estrategia": "Acompañamiento Psicológico", "Objetivo": "Salud mental y estabilidad"}
        ]
        
        st.data_editor(
            datos_apoyo,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_apoyos_bienestar"
        )
        
    # --- 11. ESTRUCTURA ADMINISTRATIVA ---
    st.markdown("---")
    st.header("11. Estructura Administrativa")
    
    # 11.1 Imagen de la Estructura
    st.subheader("11.1. Estructura Administrativa del Programa")
    st.info("""
        **Indicaciones:** Cargue la representación gráfica de la estructura organizativa del programa. 
        Recuerde que debe visualizarse la jerarquía desde la **Vicerrectoría de Enseñanza y Aprendizaje** hacia el Programa.
    """)

    with st.container(border=True):
        img_estructura = st.file_uploader(
            "Cargar Organigrama del Programa (PNG, JPG) :red[•]",
            type=["png", "jpg", "jpeg"],
            key="upload_estructura_admin"
        )
        
        if img_estructura:
            st.image(img_estructura, caption="Vista previa: Estructura Administrativa", use_container_width=True)

    st.write("")

    # 11.2 Órganos de decisión (Cuadros Paralelos)
    st.subheader("11.2. Órganos de decisión")
    st.markdown("Describa la conformación y dinámica de los cuerpos colegiados:")

    with st.container(border=True):
        col_comite, col_consejo = st.columns(2)
        
        with col_comite:
            st.markdown("### **Comité Curricular**")
            st.text_area(
                "Descripción del Comité :red[•]",
                placeholder="Conformación (Director, docentes, egresados...), periodicidad de reuniones y funciones principales...",
                key="desc_comite_curricular",
                height=250
            )
            
        with col_consejo:
            st.markdown("### **Consejo de Facultad**")
            st.text_area(
                "Descripción del Consejo :red[•]",
                placeholder="Conformación (Decano, representantes...), periodicidad y rol en la toma de decisiones del programa...",
                key="desc_consejo_facultad",
                height=250
            )

    # Nota de recordatorio institucional
    st.caption("Nota: Estas descripciones deben estar alineadas con el Estatuto General y los reglamentos internos de la I.U. Pascual Bravo.")

    # --- 12. EVALUACIÓN Y MEJORAMIENTO CONTINUO ---
    st.markdown("---")
    st.header("12. Evaluación y Mejoramiento Continuo")
    
    # 12.1 Sistema de Aseguramiento de la Calidad
    st.subheader("12.1. Sistema de Aseguramiento de la Calidad del Programa")
    
    st.info("""
        **Indicaciones:** Describa los procesos específicos del programa para garantizar la calidad académica. 
        Debe evidenciar cómo se evalúa el desempeño, cómo se identifican oportunidades de mejora 
        y la ejecución de planes de acción alineados con la I.U. Pascual Bravo.
    """)

    with st.container(border=True):
        aseguramiento_calidad_desc = st.text_area(
            "Descripción del Sistema de Calidad y Mejora Continua :red[•]",
            value=ej.get("calidad_mejora_desc", ""),
            height=350,
            placeholder="""Ejemplo: El programa implementa el Modelo de Autoevaluación Institucional, realizando jornadas semestrales de revisión de indicadores de... 
Se recolecta información de fuentes primarias (estudiantes, docentes, egresados y empleadores) para alimentar el Plan de Mejoramiento Continuo (PMC). 
Como resultado, se han ejecutado acciones enfocadas en la actualización de contenidos y fortalecimiento de laboratorios...""",
            key="input_aseguramiento_calidad"
        )

    # Bloque de apoyo conceptual
    with st.expander("🔍 Puntos clave para esta sección"):
        st.markdown("""
        Para una redacción robusta, asegúrese de mencionar:
        * **Autoevaluación:** Periodicidad y actores involucrados.
        * **Fuentes de Información:** Encuestas, pruebas Saber Pro, comités.
        * **Planes de Mejoramiento:** Cómo se transforman los hallazgos en acciones concretas.
        * **Impacto:** Resultados obtenidos de ciclos de mejora anteriores.
        """)

    
    generar = st.form_submit_button("GENERAR DOCUMENTO PEP", type="primary")

#  LÓGICA DE GENERACIÓN DEL WORD 
if generar:
    denom = st.session_state.get("denom_input", "")
    titulo = st.session_state.get("titulo_input", "")
    snies = st.session_state.get("snies_input", "")
    semestres = st.session_state.get("semestres_input", "") 
    lugar = st.session_state.get("lugar_input", "")
    creditos_actuales = st.session_state.get("cred", "")
    estudiantes = st.session_state.get("estudiantes_input", "")
    acuerdo = st.session_state.get("acuerdo_input", "")
    instancia = st.session_state.get("instancia_input", "")
    semestres_actuales = st.session_state.get("semestres_input", "") # Nuevo campo
   
    # Registros Calificados y acreditaciones
    reg1 = st.session_state.get("reg1", "")
    reg2 = st.session_state.get("reg2", "")
    reg3 = st.session_state.get("reg3", "")
    acred1 = st.session_state.get("acred1", "")
    acred2 = st.session_state.get("acred2", "")
    
    # Planes de Estudio - Versión 1 (Actual)
    p1_nom = st.session_state.get("p1_nom", "")
    p1_fec = st.session_state.get("p1_fec", "")
    p1_cred = st.session_state.get("p1_cred", "")
    p1_sem = st.session_state.get("p1_sem", "")
    
    # Planes de Estudio - Versión 2 (Anterior)
    p2_nom = st.session_state.get("p2_nom", "")
    p2_fec = st.session_state.get("p2_fec", "")
    p2_cred = st.session_state.get("p2_cred", "")
    p2_sem = st.session_state.get("p2_sem", "")

    # Planes de Estudio - Versión 3 (Antiguo)
    p3_nom = st.session_state.get("p3_nom", "")
    p3_fec = st.session_state.get("p3_fec", "")
    p3_cred = st.session_state.get("p3_cred", "")
    p3_sem = st.session_state.get("p3_sem", "")
   
    if not denom or not reg1:
        st.error("⚠️ Falta información obligatoria (Denominación o Registro Calificado).")
    else:     
        # 1. Cargar la Plantilla
        ruta_plantilla = "PlantillaPEP.docx"  # Asegúrate que el nombre es exacto
        
        if not os.path.exists(ruta_plantilla):
            st.error(f"❌ No encuentro el archivo '{ruta_plantilla}'. Súbelo a la carpeta.")
        else:
            doc = Document(ruta_plantilla)
        datos_portada = {
            "{{DENOMINACION}}": denom.upper(), # Convertimos a MAYÚSCULAS
            "{{SNIES}}": snies,
            # Puedes agregar más aquí si tienes {{TITULO}}, {{LUGAR}}, etc.
        }
        
        reemplazar_en_todo_el_doc(doc, datos_portada)
        
           
            # 1. CREACIÓN
        texto_base = (
                f"El Programa de {denom} fue creado mediante el {acuerdo} del {instancia} "
                f"y aprobado mediante la {reg1} del Ministerio de Educación Nacional "
                f"con código SNIES {snies}"
            )
        if reg3:
            texto_historia = f"{texto_base}, posteriormente recibe la renovación del registro calificado a través de la {reg2} y la {reg3}."
        elif reg2:
            texto_historia = f"{texto_base}, posteriormente recibe la renovación del registro calificado a través de la {reg2}."
        else:
            texto_historia = f"{texto_base}."

        # MOTIVO CREACIÓN
        if motivo and motivo.strip():
            parrafo_motivo = motivo
        else:
            parrafo_motivo ="No se suministró información sobre el motivo de creación."

        # MODIFICACIONES CURRICULARES
        intro_planes = (
            f"El plan de estudios del Programa de {denom} ha sido objeto de procesos periódicos de evaluación, "
            f"con el fin de asegurar su pertinencia académica y su alineación con los avances tecnológicos "
            f"y las demandas del entorno. Como resultado, "
        )

        if p1_nom and p2_nom:
            # CASO 3 PLANES: Menciona P1 (Viejo) -> P2 (Medio) -> P3 (Actual)
            parrafo_planes = (
                f"{intro_planes}se han realizado las modificaciones curriculares al plan {p1_nom} "
                f"aprobado mediante {p1_fec}, con {p1_cred} créditos y {p1_sem} semestres, "
                f"posteriormente se actualiza al plan {p2_nom} mediante {p2_fec}, con {p2_cred} créditos y {p2_sem} semestres "
                f"y por último al plan de estudio vigente {p3_nom} mediante {p3_fec}, con {p3_cred} créditos y {p3_sem} semestres."
            )
            
        elif p2_nom: 
            # CASO 2 PLANES: Asumimos que P2 es el anterior y P3 el actual
            # (P2 -> P3)
            parrafo_planes = (
                f"{intro_planes}se han realizado las modificaciones curriculares al plan {p2_nom} "
                f"aprobado mediante {p2_fec}, con {p2_cred} créditos y {p2_sem} semestres, "
                f"posteriormente se actualiza al plan de estudio vigente {p3_nom} mediante {p3_fec}, "
                f"con {p3_cred} créditos y {p3_sem} semestres."
            )

        elif p1_nom:
            # CASO ALTERNATIVO 2 PLANES: Solo llenaron P1 (Viejo) y P3 (Actual), saltándose el P2
            # (P1 -> P3)
            parrafo_planes = (
                f"{intro_planes}se han realizado las modificaciones curriculares al plan {p1_nom} "
                f"aprobado mediante {p1_fec}, con {p1_cred} créditos y {p1_sem} semestres, "
                f"posteriormente se actualiza al plan de estudio vigente {p3_nom} mediante {p3_fec}, "
                f"con {p3_cred} créditos y {p3_sem} semestres."
            )
            
        else:
            # CASO 1 PLAN (Solo existe el actual P3)
            # Preparamos variables por si faltan datos para que no salga vacío
            nom = p3_nom if p3_nom else "[FALTA NOMBRE PLAN VIGENTE]"
            fec = p3_fec if p3_fec else "[FALTA FECHA]"
            
            parrafo_planes = (
                f"{intro_planes}se estableció el plan de estudios vigente {nom} "
                f"aprobado mediante {fec}, con {p3_cred} créditos y {p3_sem} semestres."
            )
  
        # ACREDITACIÓN
        texto_acred = "" 
        
        acred1 = str(st.session_state.get("acred1", "")).strip()
        acred2 = str(st.session_state.get("acred2", "")).strip()
        
        if acred1 and acred2:
            # Caso: Dos acreditaciones
            texto_acred = (
                f"El programa obtuvo por primera vez la Acreditación en alta calidad otorgada por el "
                f"Consejo Nacional de Acreditación (CNA) a través de la resolución {acred1}, "
                f"esta le fue renovada mediante resolución {acred2}, reafirmando la solidez "
                f"académica, administrativa y de impacto social del Programa."
            )
        elif acred1:
             # Caso: Solo una acreditación 
            texto_acred = (
                f"El programa obtuvo la Acreditación en alta calidad otorgada por el "
                f"Consejo Nacional de Acreditación (CNA) a través de la resolución {acred1}, "
                f"como reconocimiento a su solidez académica, administrativa y de impacto social."
            )
       
       # RECONOCIMIENTOS
        texto_recons = ""
        recon_data = st.session_state.get("recon_data", [])
        
        # Filtramos los vacíos
        recons_validos = [
            r for r in recon_data 
            if isinstance(r, dict) and str(r.get("Nombre del premio", "")).strip()
        ]
        
        if recons_validos:
            # Encabezado del párrafo de reconocimientos
            intro_recon = (
                f"Adicionalmente, el Programa de {denom} ha alcanzado importantes logros académicos e institucionales "
                f"que evidencian su calidad y compromiso con la excelencia. Entre ellos se destacan:"
            )
            lista_items = []        
            for r in recons_validos:
                premio = str(r.get("Nombre del premio", "Premio")).strip()
                anio = str(r.get("Año", "")).strip()
                ganador = str(r.get("Nombre del Ganador", "")).strip()
                cargo = str(r.get("Cargo", "")).strip()
                            
                item = f"• {premio} ({anio}): Otorgado a {ganador}, en su calidad de {cargo}."
                lista_items.append(item)
            
            texto_recons = intro_recon + "\n" + "\n".join(lista_items)

        #LINEA DE TIEMPO
        texto_timeline = ""
        eventos = []

        # Función auxiliar para sacar el año (busca 19XX o 20XX en cualquier lado)
        def obtener_anio(texto):
            if not texto: return 9999 # Si no hay fecha, lo mandamos al final
            match = re.search(r'\b(19|20)\d{2}\b', str(texto))
            return int(match.group(0)) if match else 9999

        # --- A. Agregamos Resoluciones ---
        if reg1: eventos.append((obtener_anio(reg1), f"Creación y Registro Calificado inicial ({reg1})."))
        if reg2: eventos.append((obtener_anio(reg2), f"Renovación del Registro Calificado ({reg2})."))
        if reg3: eventos.append((obtener_anio(reg3), f"Segunda Renovación Registro Calificado ({reg3})."))

        # --- B. Agregamos Planes (P1=Viejo, P2=Medio, P3=Actual) ---
        # Solo agregamos si hay fecha válida
        if p1_fec: eventos.append((obtener_anio(p1_fec), f"Inicio Plan de Estudios {p1_nom}."))
        if p2_fec: eventos.append((obtener_anio(p2_fec), f"Actualización Curricular - Plan {p2_nom}."))
        if p3_fec: eventos.append((obtener_anio(p3_fec), f"Implementación Plan Vigente {p3_nom}."))

        # --- C. Agregamos Acreditaciones ---
        if acred1: eventos.append((obtener_anio(acred1), f"Obtención Acreditación de Alta Calidad ({acred1})."))
        if acred2: eventos.append((obtener_anio(acred2), f"Renovación Acreditación de Alta Calidad ({acred2})."))

        # --- D. Agregamos Reconocimientos (Solo los destacados) ---
        if recons_validos:
            for r in recons_validos:
                anio_r = obtener_anio(r.get("Año", ""))
                nom_r = r.get("Nombre del premio", "Premio")
                # Solo agregamos si encontramos un año válido para no ensuciar la línea
                if anio_r != 9999:
                     eventos.append((anio_r, f"Reconocimiento: {nom_r}."))

        # --- E. Ordenar y Construir Texto ---
        # Ordenamos la lista por el año (el primer elemento de la tupla)
        eventos.sort(key=lambda x: x[0])

        if eventos:
            # Creamos un "título" visual en negrita o separado
            lines = ["Hitos relevantes en la línea de tiempo del programa:"]
            
            last_year = 0
            for anio, desc in eventos:
                if anio != 9999:
                    lines.append(f"• {anio}: {desc}")
            
            texto_timeline = "\n".join(lines)

   
        # UNIÓN FINAL E INSERCIÓN
        partes = [
            texto_historia,  # 1. Creación
            parrafo_motivo,  # 2. Motivo
            parrafo_planes,  # 3. Planes
            texto_acred,     # 4. Acreditación
            texto_recons,    # 5. Reconocimientos
            texto_timeline   # 6. Línea de Tiempo (¡Aquí va!)
        ]
        
        # Unimos todo en un solo bloque de texto grande
        texto_final_completo = "\n\n".join([p for p in partes if p and p.strip()])
        
        # Insertamos en el Word en el lugar correcto
        insertar_texto_debajo_de_titulo(doc, "Historia del programa", texto_final_completo)
                
        # 1.2 GENERALIDADES DEL PROGRAMA
        v_denom = str(st.session_state.get("denom_input", "")).strip()
        v_titulo = str(st.session_state.get("titulo_input", "")).strip()
        v_nivel = str(st.session_state.get("nivel_formacion_widget", "")).strip()
        v_snies = str(st.session_state.get("snies_input", "")).strip()
        v_modalidad = str(st.session_state.get("modalidad_input", "")).strip()
        v_acuerdo = str(st.session_state.get("acuerdo_input", "")).strip()
        v_periodicidad = str(st.session_state.get("periodicidad_input", "")).strip()
        v_lugar = str(st.session_state.get("lugar_input", "")).strip()
        v_creditos = str(st.session_state.get("cred", "")).strip() 
        v_area = str(st.session_state.get("area", "")).strip()

        # Cálculo del Registro Calificado Vigente
        r1 = str(st.session_state.get("reg1", "")).strip()
        r2 = str(st.session_state.get("reg2", "")).strip()
        r3 = str(st.session_state.get("reg3", "")).strip()
        reg_final = r3 if r3 else (r2 if r2 else r1)

        # B. Crear la Lista de Datos (Ordenada tal cual la pediste)
        # ---------------------------------------------------------
        lista_datos = [
            f"● Denominación del programa: {v_denom}",
            f"● Título otorgado: {v_titulo}",
            f"● Nivel de formación: {v_nivel}",
            f"● Área de formación: {v_area}",
            f"● Modalidad de oferta: {v_modalidad}",
            f"● Acuerdo de creación: {v_acuerdo}",
            f"●Registro calificado: {reg_final}",
            f"● Créditos académicos: {v_creditos}",
            f"● Periodicidad de admisión: {v_periodicidad}",
            f"● Lugares de desarrollo: {v_lugar}",
            f"● SNIES: {v_snies}"
        ]

        # C. Función para Insertar DEBAJO de un párrafo específico
        # --------------------------------------------------------
        def insertar_lista_bajo_titulo(documento, texto_titulo, lista_items):
            """
            Busca el párrafo que contenga 'texto_titulo'.
            Si lo encuentra, inserta los items de la lista justo debajo.
            """
            for i, paragraph in enumerate(documento.paragraphs):
                # Buscamos el título (ignorando mayúsculas/minúsculas para asegurar)
                if texto_titulo.lower() in paragraph.text.lower():
                    
                    # Truco técnico: Para insertar "despues", nos paramos en el párrafo SIGUIENTE
                    # y le decimos "insertar antes de ti".
                    
                    # Verificamos si hay un párrafo siguiente
                    if i + 1 < len(documento.paragraphs):
                        p_siguiente = documento.paragraphs[i + 1]
                                                                       
                        # Estrategia Limpia: Insertamos antes del siguiente párrafo
                        for item in lista_datos:
                            p_siguiente.insert_paragraph_before(item)
                        encontrado = True
                        break # Terminamos apenas lo encontramos
                        
            if not encontrado:
                doc.add_heading("1.2. Generalidades del programa", level=2)
                for item in lista_datos:
                    doc.add_paragraph(item)

        insertar_lista_bajo_titulo(doc, "Generalidades del programa", lista_datos)
        

        # CAPÍTULO 2: REFERENTES CONCEPTUALES
        #2.1 NATURALEZA DEL PROGRAMA
       
        v_obj_nombre = str(st.session_state.get("obj_nombre_input", "")).strip()
        texto_para_pegar = "" # Contendrá la definición extensa

        if metodo_trabajo == "Automatizado (Cargar Documento Maestro)" and archivo_dm is not None:
            try:
                doc_m = Document(archivo_dm)
                t_inicio = str(st.session_state.get("inicio_def_oc", "")).strip().lower()
                t_fin = str(st.session_state.get("fin_def_oc", "")).strip().lower()
                
                p_extraidos_21 = []
                capturando_21 = False

                for p_m in doc_m.paragraphs:
                    # Usamos el texto original para el recorte final, 
                    # pero una versión limpia para la búsqueda
                    p_text_raw = p_m.text
                    p_text_low = p_text_raw.lower()
                    busqueda_ini = t_inicio.lower()
                    busqueda_fin = t_fin.lower()
                    
                    # CASO: Encontrar el inicio
                    if busqueda_ini in p_text_low and not capturando_21:
                        capturando_21 = True
                        idx_start = p_text_low.find(busqueda_ini)
                        
                        # Verificamos si el final está en este mismo párrafo
                        if busqueda_fin in p_text_low[idx_start + len(busqueda_ini):]:
                            # Si ambos están en el mismo párrafo, cortamos ambos extremos
                            idx_end = p_text_low.find(busqueda_fin, idx_start) + len(busqueda_fin)
                            p_extraidos_21.append(p_text_raw[idx_start:idx_end])
                            capturando_21 = False
                            break
                        else:
                            # Si no está el final, guardamos desde el inicio hasta el final del párrafo
                            p_extraidos_21.append(p_text_raw[idx_start:])
                        continue
                    
                    # CASO: Estamos capturando párrafos intermedios
                    if capturando_21:
                        if busqueda_fin in p_text_low:
                            # Encontramos el cierre: cortamos hasta donde termina el marcador final
                            idx_end = p_text_low.find(busqueda_fin) + len(busqueda_fin)
                            p_extraidos_21.append(p_text_raw[:idx_end])
                            capturando_21 = False
                            break
                        else:
                            # Párrafo intermedio completo
                            p_extraidos_21.append(p_text_raw)

                texto_para_pegar = "\n\n".join(p_extraidos_21)
            except Exception as e:
                st.error(f"Error en la extracción: {e}")

        # 2. INSERCIÓN EN PLACEHOLDERS {{oc}} y {{def_oc}}
        texto_nombre_completo = f"Objeto de conocimiento del programa: {v_obj_nombre}"

        for p_plan in doc.paragraphs:
            # Reemplazo del Nombre del Objeto
            if "{{oc}}" in p_plan.text:
                p_plan.text = p_plan.text.replace("{{oc}}", texto_nombre_completo)
            
            # Reemplazo de la Definición (Estricta)
            if "{{def_oc}}" in p_plan.text:
                if texto_para_pegar:
                    p_plan.text = p_plan.text.replace("{{def_oc}}", texto_para_pegar)
                    p_plan.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                else:
                    p_plan.text = p_plan.text.replace("{{def_oc}}", "")
    
    #FUNDAMENTACIÓN EPISTEMOLÓGICA                
# 1. Recuperar el texto (si no hay nada, queda vacío)
texto_final = str(st.session_state.get("fund_epi_manual", ""))

# 2. REEMPLAZO DIRECTO (Sin funciones anidadas para evitar errores)
if False:
    # Buscar en párrafos normales
    for p in doc.paragraphs:
        if "{{fundamentacion_epistemologica}}" in p.text:
            p.text = p.text.replace("{{fundamentacion_epistemologica}}", texto_final)
            p.alignment = 3 # Justificado
    
    # Buscar en tablas (esto suele ser donde se rompe si no se hace con cuidado)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if "{{fundamentacion_epistemologica}}" in celda.text:
                    # Reemplazo directo en la celda
                    for p_celda in celda.paragraphs:
                        if "{{fundamentacion_epistemologica}}" in p_celda.text:
                            p_celda.text = p_celda.text.replace("{{fundamentacion_epistemologica}}", texto_final)
                            p_celda.alignment = 3

    #GUARDAR ARCHIVO
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.success("✅ ¡Documento PEP generado!")
    st.download_button(
                label="📥 Descargar Documento PEP en Word",
                data=bio.getvalue(),
                file_name=f"PEP_Modulo1_{denom.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  )
