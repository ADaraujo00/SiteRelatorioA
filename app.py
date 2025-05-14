import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
from PIL import Image

OUTPUT_FILENAME = 'relatorio_completo.docx'

st.title("Gerador de Relatório Completo")

st.subheader("Informações Gerais:")

col1, col2, col3, col4 = st.columns(4)
product_general = col1.text_input("Product:")
project = col2.text_input("Project:")
lot = col3.text_input("Lot:")
released = col4.text_input("Released:")

col5, col6, col7, col8 = st.columns(4)
requested_by = col5.text_input("Requested by:")
performed_by = col6.text_input("Performed by:")
reviewed_by = col7.text_input("Reviewed by:")
approved_by = col8.text_input("Approved by:")

st.subheader("3. SAMPLES DESCRIPTION:")

# Layout da grade de imagens (3x4)
image_uploaders = []
for i in range(3):
    cols = st.columns(4)
    row_uploaders = []
    for j in range(4):
        index = i * 4 + j
        uploader = cols[j].file_uploader(f"Imagem {index + 1}:", type=["jpg", "jpeg", "png"])
        row_uploaders.append(uploader)
    image_uploaders.extend(row_uploaders)

st.subheader("Detalhes das Amostras:")
sample_details_input = {}
questions = ["Product", "Accessories", "Model", "Voltage", "Dimensions", "Supplier", "Quantity", "Volume"]

for question in questions:
    col_q, col_a = st.columns([1, 2])
    col_q.markdown(f"**{question}:**")
    sample_details_input[question] = col_a.text_input("", label_visibility="collapsed", key=question)

def generate_word_report(general_data, images, sample_data):
    doc = Document()

    doc.add_paragraph("Informações Gerais:")
    doc.add_paragraph(f"**Product:** {general_data['Product']}")
    doc.add_paragraph(f"**Project:** {general_data['Project']}")
    doc.add_paragraph(f"**Lot:** {general_data['Lot']}")
    doc.add_paragraph(f"**Released:** {general_data['Released']}")
    doc.add_paragraph()
    doc.add_paragraph(f"**Requested by:** {general_data['Requested by']}")
    doc.add_paragraph(f"**Performed by:** {general_data['Performed by']}")
    doc.add_paragraph(f"**Reviewed by:** {general_data['Reviewed by']}")
    doc.add_paragraph(f"**Approved by:** {general_data['Approved by']}")
    doc.add_paragraph()

    doc.add_paragraph("3. SAMPLES DESCRIPTION:")

    # Adicionar as imagens em layout de grade (3x4)
    for i in range(0, len(images), 4):
        row = doc.add_paragraph()
        for j in range(4):
            if i + j < len(images) and images[i + j] is not None:
                try:
                    img = Image.open(images[i + j])
                    width, height = img.size
                    ratio = height / width
                    new_width = 2.5
                    new_height = new_width * ratio
                    row.add_run().add_picture(images[i + j], width=Inches(new_width), height=Inches(new_height))
                except Exception as e:
                    doc.add_paragraph(f"Erro ao adicionar imagem {i+j+1}: {e}")
            row.add_run("   ")
        doc.add_paragraph()

    # Adicionar a tabela com os detalhes das amostras
    doc.add_paragraph("\nDetalhes das Amostras:")
    table = doc.add_table(rows=len(sample_data), cols=2)
    for i, (key, value) in enumerate(sample_data.items()):
        cell_left = table.cell(i, 0)
        cell_right = table.cell(i, 1)
        cell_left.text = f"{key}:"
        cell_left.paragraphs[0].runs[0].font.bold = True
        cell_right.text = value

    # Salvar o documento na memória
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("Gerar Relatório"):
    uploaded_files = [uploader for uploader in image_uploaders if uploader is not None]
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
    word_buffer = generate_word_report(general_info, uploaded_files, sample_details_input)
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
1. Preencha as informações gerais.
2. Carregue até 12 imagens na grade.
3. Preencha os detalhes das amostras.
4. Clique em "Gerar Relatório" para baixar o documento Word.
""")
