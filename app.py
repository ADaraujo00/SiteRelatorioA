import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
from PIL import Image

OUTPUT_FILENAME = 'relatorio_com_multiplas_imagens.docx'

st.title("Gerador de Relatório Word")

st.subheader("Preencha as informações:")

col1, col2, col3, col4 = st.columns(4)
product = col1.text_input("Product:")
project = col2.text_input("Project:")
lot = col3.text_input("Lot:")
released = col4.text_input("Released:")

col5, col6, col7, col8 = st.columns(4)
requested_by = col5.text_input("Requested by:")
performed_by = col6.text_input("Performed by:")
reviewed_by = col7.text_input("Reviewed by:")
approved_by = col8.text_input("Approved by:")

uploaded_images = st.file_uploader("Carregar Imagens:", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

report_data = {
    "Product": product,
    "Project": project,
    "Lot": lot,
    "Released": released,
    "Requested by": requested_by,
    "Performed by": performed_by,
    "Reviewed by": reviewed_by,
    "Approved by": approved_by
}

def generate_word_report(data, images):
    doc = Document()

    doc.add_paragraph(f"**Product:** {data['Product']}")
    doc.add_paragraph(f"**Project:** {data['Project']}")
    doc.add_paragraph(f"**Lot:** {data['Lot']}")
    doc.add_paragraph(f"**Released:** {data['Released']}")
    doc.add_paragraph()
    doc.add_paragraph(f"**Requested by:** {data['Requested by']}")
    doc.add_paragraph(f"**Performed by:** {data['Performed by']}")
    doc.add_paragraph(f"**Reviewed by:** {data['Reviewed by']}")
    doc.add_paragraph(f"**Approved by:** {data['Approved by']}")
    doc.add_paragraph()

    if images:
        doc.add_paragraph("Imagens:")
        for img_file in images:
            try:
                img = Image.open(img_file)
                width, height = img.size
                ratio = height / width
                new_width = 4  # Polegadas (ajuste conforme necessário)
                new_height = new_width * ratio
                doc.add_picture(img_file, width=Inches(new_width), height=Inches(new_height))
                doc.add_paragraph() # Adiciona um espaço após cada imagem
            except Exception as e:
                doc.add_paragraph(f"Erro ao adicionar imagem: {e}")
    else:
        doc.add_paragraph("Nenhuma imagem carregada.")

    # Salvar o documento na memória
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("Gerar Relatório Word"):
    if all(report_data.values()):
        word_buffer = generate_word_report(report_data, uploaded_images)
        if word_buffer:
            st.download_button(
                label="Baixar Relatório Word",
                data=word_buffer.getvalue(),
                file_name=OUTPUT_FILENAME,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Por favor, preencha todos os campos.")

st.markdown("""
---
**Instruções:**
1. Preencha todos os campos solicitados (os primeiros quatro na primeira linha, os seguintes na segunda).
2. Carregue as imagens desejadas usando o uploader. Você pode selecionar múltiplos arquivos.
3. Clique em "Gerar Relatório Word" para baixar o documento com as informações e as imagens.
""")
