import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
from PIL import Image
import docx  # Importamos a biblioteca docx explicitamente

OUTPUT_FILENAME = 'relatorio_completo.docx'
IMAGE_SIZE_CM = 4
IMAGE_SIZE_INCHES = IMAGE_SIZE_CM / 2.54
# Salve a imagem com este nome na mesma pasta do script
HEADER_IMAGE_PATH = 'electrolux_header.png'

st.title("Gerador de Relatório Completo")

report_number = st.text_input("**Número do Relatório:**", "TR00001")

st.subheader("Informações Gerais:")

col1, col2, col3, col4 = st.columns(4)
product_general = col1.text_input("**Product:**")
project = col2.text_input("**Project:**")
lot = col3.text_input("**Lot:**")
released = col4.text_input("**Released:**")

col5, col6, col7, col8 = st.columns(4)
requested_by = col5.text_input("**Requested by:**")
performed_by = col6.text_input("**Performed by:**")
reviewed_by = col7.text_input("**Reviewed by:**")
approved_by = col8.text_input("**Approved by:**")

st.subheader("3. SAMPLES DESCRIPTION:")

image_labels = [
    "Foto base frente", "Foto base costas", "Foto base cima", "Foto base baixo",
    "Produto inteiro frente com jarra blender", "Produto inteiro lado com jarra blender",
    "Produto inteiro costas com jarra blender", "Produto inteiro lado com jarra blender",
    "Jarra lado", "Jarra frente", "Jarra cima", "Jarra embaixo"
]

image_uploaders = []
cols_grid = [st.columns(4) for _ in range(3)]
for i in range(3):
    for j in range(4):
        index = i * 4 + j
        uploader = cols_grid[i][j].file_uploader(f"{image_labels[index]}:", type=[
                                                                            "jpg", "jpeg", "png"], key=f"image_{index}")
        image_uploaders.append(uploader)

st.subheader("Detalhes das Amostras:")
sample_details_input = {}
questions = ["Product", "Accessories", "Model", "Voltage",
             "Dimensions", "Supplier", "Quantity", "Volume"]

for question in questions:
    col_q, col_a = st.columns([1, 2])
    col_q.markdown(f"**{question}:**")
    sample_details_input[question] = col_a.text_input(
        "", label_visibility="collapsed", key=question)


def generate_word_report(report_num, general_data, images, sample_data):
    doc = Document()

    # Adicionar a imagem de cabeçalho
    try:
        # Ajuste a largura conforme necessário
        doc.add_picture(HEADER_IMAGE_PATH, width=Inches(6))
        header_paragraph = doc.paragraphs[-1]
        header_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception as e:
        doc.add_paragraph(f"Erro ao adicionar imagem de cabeçalho: {e}")

    # Adicionar o número do relatório
    doc.add_paragraph(f"Relatório No: {report_num}")
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
    cells_row1 = table_row1.rows[0].cells
    cells_row1[0].text = "Product:"
    cells_row1[0].paragraphs[0].runs[0].font.bold = True  # Negrito no Word
    cells_row1[1].text = "Project:"
    cells_row1[1].paragraphs[0].runs[0].font.bold = True  # Negrito no Word
    cells_row1[2].text = "Lot:"
    cells_row1[2].paragraphs[0].runs[0].font.bold = True  # Negrito no Word
    cells_row1[3].text = "Released:"
    cells_row1[3].paragraphs[0].runs[0].font.bold = True  # Negrito no Word
    cells_row2 = table_row1.rows[1].cells
    cells_row2[0].text = general_data['Product']
    cells_row2[1].text = general_data['Project']
    cells_row2[2].text = general_data['Lot']
    cells_row2[3].text = general_data['Released']
    remove_table_borders(table_row1)

    table_row2 = doc.add_table(rows=2, cols=4)
    cells_row3 = table_row2.rows[0].cells
    cells_row3[0].text = "Requested by:"
    cells_row3[0].paragraphs[0].runs[0].font.bold = True  # Negrito no Word
    cells_row3[1].text = "Performed by:"
    cells_row3[1].paragraphs[0].runs[0].font.bold = True  # Negrito no Word
    cells_row3[2].text = "Reviewed by:"
    cells_row3[1].paragraphs[0].runs[0].font.bold = True  # Negrito no Word
    cells_row3[3].text = "Approved by:"
    cells_row3[1].paragraphs[0].runs[0].font.bold = True  # Negrito no Word
    cells_row4 = table_row2.rows[1].cells
    cells_row4[0].text = general_data['Requested by']
    cells_row4[1].text = general_data['Performed by']
    cells_row4[2].text = general_data['Reviewed by']
    cells_row4[3].text = general_data['Approved by']
    remove_table_borders(table_row2)

    doc.add_paragraph()  # Espaço antes da próxima seção

    doc.add_paragraph("3. SAMPLES DESCRIPTION:")

    # Adicionar as imagens em tabela 3x4 com tamanho fixo
    image_table = doc.add_table(rows=3, cols=4)
    for i in range(3):
        for j in range(4):
            index = i * 4 + j
            cell = image_table.cell(i, j)
            if index < len(images) and images[index] is not None:
                try:
                    cell.paragraphs[0].add_run().add_picture(
                        images[index],
                        width=Inches(IMAGE_SIZE_INCHES),
                        height=Inches(IMAGE_SIZE_INCHES)
                    )
                except Exception as e:
                    cell.text = f"Erro ao adicionar imagem {index + 1}: {e}"

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


if st.button("Gerar Relatório"):
    uploaded_files = [
        uploader for uploader in image_uploaders if uploader is not None]
    general_info = {
        "Product": product_general,
        "Project": project,
        "Lot": lot,
        "Released": released,
        "Requested by": requested_by,
        "Performed by": performed_by,
        "Reviewed by": reviewed_by,
        "Approved by": approved_by
    }
    word_buffer = generate_word_report(
        report_number, general_info, uploaded_files, sample_details_input)
    if word_buffer:
        st.download_button(
            label="Baixar Relatório Word",
            data=word_buffer.getvalue(),
            file_name=OUTPUT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("""
---
**Instruções:**
1. Insira o **Número do Relatório**.
2. Preencha as **Informações Gerais**.
3. Carregue as imagens correspondentes a cada descrição.
4. Preencha os **Detalhes das Amostras**.
5. Clique em "Gerar Relatório" para baixar o documento Word.
""")
