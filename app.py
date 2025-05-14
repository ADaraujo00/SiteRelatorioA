def generate_word_report(general_data, images, sample_data):
    doc = Document()

    # Título centralizado
    title_para = doc.add_paragraph("Blender Performance evaluation")
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph() # Espaço após o título

    def remove_table_borders(table):
        tblPr = table._tbl.get_or_add_tblPr()
        borders = docx.oxml.shared.OxmlElement('w:tblBorders')
        top = docx.oxml.shared.OxmlElement('w:top')
        top.set(docx.oxml.ns.qn('w:val'), 'nil')
        left = docx.oxml.shared.OxmlElement('w:left')
        left.set(docx.oxml.ns.qn('w:val'), 'nil')
        bottom = docx.oxml.shared.OxmlElement('w:bottom')
        bottom.set(docx.oxml.ns.qn('w:val'), 'nil')
        right = docx.oxml.shared.OxmlElement('w:right')
        right.set(docx.oxml.ns.qn('w:val'), 'nil')
        insideH = docx.oxml.shared.OxmlElement('w:insideH')
        insideH.set(docx.oxml.ns.qn('w:val'), 'nil')
        insideV = docx.oxml.shared.OxmlElement('w:insideV')
        insideV.set(docx.oxml.ns.qn('w:val'), 'nil')
        borders.extend([top, left, bottom, right, insideH, insideV])
        tblPr.append(borders)

    # Primeira linha de informações gerais
    table1 = doc.add_table(rows=1, cols=4)
    cells1 = table1.rows[0].cells
    cells1[0].text = f"Product: {general_data['Product']}"
    cells1[1].text = f"Project: {general_data['Project']}"
    cells1[2].text = f"Lot: {general_data['Lot']}"
    cells1[3].text = f"Released: {general_data['Released']}"
    remove_table_borders(table1)

    doc.add_paragraph() # Espaço entre as linhas

    # Segunda linha de informações gerais
    table2 = doc.add_table(rows=1, cols=4)
    cells2 = table2.rows[0].cells
    cells2[0].text = f"Requested by: {general_data['Requested by']}"
    cells2[1].text = f"Performed by: {general_data['Performed by']}"
    cells2[2].text = f"Reviewed by: {general_data['Reviewed by']}"
    cells2[3].text = f"Approved by: {general_data['Approved by']}"
    remove_table_borders(table2)

    doc.add_paragraph() # Espaço antes da próxima seção

    doc.add_paragraph("3. SAMPLES DESCRIPTION:")

    # Adicionar as imagens em layout de grade (3x4)
    for i in range(0, len(images), 4):
        row = doc.add_paragraph()
        for j in range(4):
            index = i * 4 + j
            if index < len(images) and images[index] is not None:
                try:
                    img = Image.open(images[index])
                    width, height = img.size
                    ratio = height / width
                    new_width = 2.5
                    new_height = new_width * ratio
                    row.add_run().add_picture(images[index], width=Inches(new_width), height=Inches(new_height))
                except Exception as e:
                    doc.add_paragraph(f"Erro ao adicionar imagem {index + 1}: {e}")
            row.add_run("   ")
        doc.add_paragraph()

    # Adicionar a tabela com os detalhes das amostras
    doc.add_paragraph("\nDetalhes das Amostras:")
    table_samples = doc.add_table(rows=len(sample_data), cols=2)
    for i, (key, value) in enumerate(sample_data.items()):
        cell_left = table_samples.cell(i, 0)
        cell_right = table_samples.cell(i, 1)
        cell_left.text = f"{key}:"
        cell_left.paragraphs[0].runs[0].font.bold = True
        cell_right.text = value

    # Salvar o documento na memória
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
