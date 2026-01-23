# ... dentro del bucle de secciones ...

if config["tipo"] == "ia":
    st.write(f"✍️ Redactando narrativa para: {seccion_nombre}...")
    texto_ia = redactar_seccion_ia(seccion_nombre, respuestas_finales[seccion_nombre])
    doc.add_paragraph(texto_ia)

elif config["tipo"] == "especial_pascual":
    # 1. Insertar Texto Institucional Obligatorio
    p_inst = doc.add_paragraph()
    p_inst.add_run("La fundamentación académica del Programa responde a los Lineamientos Académicos y Curriculares (LAC) de la I.U. Pascual Bravo...").bold = False
    # (Aquí pegarías el texto completo que me pasaste para que aparezca tal cual)
    doc.add_paragraph("La fundamentación académica del Programa responde a los Lineamientos Académicos y Curriculares (LAC) de la I.U. Pascual Bravo, garantizando la coherencia entre el diseño curricular, la metodología pedagógica y los estándares de calidad definidos por el Ministerio de Educación Nacional de Colombia...")
    doc.add_paragraph("Dentro de los LAC se establece la política de créditos académicos de la Universidad, siendo ésta el conjunto de lineamientos y procedimientos que rigen la asignación de créditos...")

    # 2. Subtítulo Rutas Educativas
    doc.add_heading("Rutas educativas: Certificaciones Temáticas Tempranas", level=3)
    doc.add_paragraph("Las Certificaciones Temáticas Tempranas son el resultado del agrupamiento de competencias y cursos propios del currículo...")

    # 3. Insertar Tabla de Certificaciones
    # Aquí leeríamos los datos del st.data_editor que crearemos en la interfaz
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = 'Table Grid'
    hdr_cells = tabla.rows[0].cells
    hdr_cells[0].text = 'Certificación'
    hdr_cells[1].text = 'Curso'
    hdr_cells[2].text = 'Créditos'
    
    # Llenar con los datos que el usuario meta en el editor
    for cert en st.session_state.certificaciones:
        row_cells = tabla.add_row().cells
        row_cells[0].text = cert['Nombre']
        row_cells[1].text = cert['Curso']
        row_cells[2].text = str(cert['Créditos'])
