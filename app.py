from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.styles import WD_STYLE_TYPE
from docx.enum.style import WD_STYLE
from docx.shared import RGBColor

def generate_word_report(report_num, general_data, images, sample_data):
    doc = Document()

    # Definir a fonte Arial tamanho 10 como padrão
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10)

    # Adicionar a imagem de cabeçalho
    try:
        doc.add_picture(HEADER_IMAGE_PATH, width=Inches(19.50 / 2.54))
        header_paragraph = doc.paragraphs[-1]
        header_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception as e:
        doc.add_paragraph(f"Erro ao adicionar imagem de cabeçalho: {e}")
        paragraph = doc.paragraphs[-1]
        paragraph.style.font.name = 'Arial'
        paragraph.style.font.size = Pt(10)

    # Adicionar o número do relatório
    paragraph = doc.add_paragraph(f"Relatório No: {report_num}")
    paragraph.style.font.name = 'Arial'
    paragraph.style.font.size = Pt(10)
    doc.add_paragraph()  # Espaço

    def remove_table_borders(table):
        tblPr = table._element.xpath('./w:tblPr')[0]
        borders = docx.oxml.shared.OxmlElement('w:tblBorders')
        for tag in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = docx.oxml.shared.OxmlElement(f'w:{tag}')
            border.set(docx.oxml.ns.qn('w:val'), 'nil')
            borders.append(border)
        tblPr.append(borders)

    # Informações Gerais em tabela com respostas abaixo
    table_row1 = doc.add_table(rows=2, cols=4)
    for row in table_row1.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.style.font.name = 'Arial'
                paragraph.style.font.size = Pt(10)
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
                    if "Product:" in cell.text or "Project:" in cell.text or "Lot:" in cell.text or "Released:" in cell.text:
                        run.font.bold = True
    cells_row1 = table_row1.rows[0].cells
    cells_row1[0].text = "Product:"
    cells_row1[1].text = "Project:"
    cells_row1[2].text = "Lot:"
    cells_row1[3].text = "Released:"
    cells_row2 = table_row1.rows[1].cells
    cells_row2[0].text = general_data['Product']
    cells_row2[1].text = general_data['Project']
    cells_row2[2].text = general_data['Lot']
    cells_row2[3].text = general_data['Released']
    remove_table_borders(table_row1)

    table_row2 = doc.add_table(rows=2, cols=4)
    for row in table_row2.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.style.font.name = 'Arial'
                paragraph.style.font.size = Pt(10)
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
                    if "Requested by:" in cell.text or "Performed by:" in cell.text or "Reviewed by:" in cell.text or "Approved by:" in cell.text:
                        run.font.bold = True
    cells_row3 = table_row2.rows[0].cells
    cells_row3[0].text = "Requested by:"
    cells_row3[1].text = "Performed by:"
    cells_row3[2].text = "Reviewed by:"
    cells_row3[3].text = "Approved by:"
    cells_row4 = table_row2.rows[1].cells
    cells_row4[0].text = general_data['Requested by']
    cells_row4[1].text = general_data['Performed by']
    cells_row4[2].text = general_data['Reviewed by']
    cells_row4[3].text = general_data['Approved by']
    remove_table_borders(table_row2)

    paragraph = doc.add_paragraph()  # Espaço antes da próxima seção
    paragraph.style.font.name = 'Arial'
    paragraph.style.font.size = Pt(10)

    paragraph = doc.add_paragraph("3. SAMPLES DESCRIPTION:")
    paragraph.style.font.name = 'Arial'
    paragraph.style.font.size = Pt(10)
    paragraph.runs[0].font.bold = True

    # Adicionar as imagens em tabela 3x4 com tamanho fixo
    image_table = doc.add_table(rows=3, cols=4)
    for i in range(3):
        for j in range(4):
            index = i * 4 + j
            cell = image_table.cell(i, j)
            for paragraph in cell.paragraphs:
                paragraph.style.font.name = 'Arial'
                paragraph.style.font.size = Pt(10)
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
            if index < len(images) and images[index] is not None:
                try:
                    cell.paragraphs[0].add_run().add_picture(
                        images[index],
                        width=Inches(IMAGE_SIZE_INCHES),
                        height=Inches(IMAGE_SIZE_INCHES)
                    )
                except Exception as e:
                    cell.text = f"Erro ao adicionar imagem {index + 1}: {e}"
                    for paragraph in cell.paragraphs:
                        paragraph.style.font.name = 'Arial'
                        paragraph.style.font.size = Pt(10)
                        for run in paragraph.runs:
                            run.font.name = 'Arial'
                            run.font.size = Pt(10)

    paragraph = doc.add_paragraph("\nDetalhes das Amostras:")
    paragraph.style.font.name = 'Arial'
    paragraph.style.font.size = Pt(10)
    paragraph.runs[0].font.bold = True
    table_samples = doc.add_table(rows=len(sample_data), cols=2)
    for i, (key, value) in enumerate(sample_data.items()):
        cell_left = table_samples.cell(i, 0)
        cell_right = table_samples.cell(i, 1)
        for paragraph in cell_left.paragraphs:
            paragraph.style.font.name = 'Arial'
            paragraph.style.font.size = Pt(10)
            for run in paragraph.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(10)
                run.font.bold = True
        cell_left.text = f"{key}:"
        for paragraph in cell_right.paragraphs:
            paragraph.style.font.name = 'Arial'
            paragraph.style.font.size = Pt(10)
            for run in paragraph.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(10)
        cell_right.text = value

    # Salvar o documento na memória
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
