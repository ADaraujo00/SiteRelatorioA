import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO

OUTPUT_FILENAME = 'relatorio_gerado.docx'
BODY_IMAGE_PATH = 'imagem_inicial.png'  # Nome do arquivo da imagem a ser inserida no corpo

def generate_word_report(data):
    doc = Document()

    # Adicionar imagem no COMEÇO do corpo
    try:
        doc.add_picture(BODY_IMAGE_PATH, width=Inches(4.0)) # Ajuste a largura conforme necessário
        doc.add_paragraph() # Adiciona uma linha em branco após a imagem
    except FileNotFoundError:
        st.warning(f"Arquivo de imagem '{BODY_IMAGE_PATH}' não encontrado.")

    doc.add_paragraph(f"**Product:** {data['Product']}")
    doc.add_paragraph(f"**Project:** {data['Project']}")
    doc.add_paragraph(f"**Lot:** {data['Lot']}")
    doc.add_paragraph(f"**Released:** {data['Released']}")
    doc.add_paragraph()
    doc.add_paragraph(f"**Requested by:** {data['Requested by']}")
    doc.add_paragraph(f"**Performed by:** {data['Performed by']}")
    doc.add_paragraph(f"**Reviewed by:** {data['Reviewed by']}")
    doc.add_paragraph(f"**Approved by:** {data['Approved by']}")

    # Salvar o documento na memória
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

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

if st.button("Gerar Relatório Word"):
    if all(report_data.values()):
        # Gerar relatório com perguntas em negrito
        doc = Document()
        try:
            doc.add_picture(BODY_IMAGE_PATH, width=Inches(4.0))
            doc.add_paragraph()
        except FileNotFoundError:
            st.warning(f"Arquivo de imagem '{BODY_IMAGE_PATH}' não encontrado.")

        doc.add_paragraph(f"**Product:** {report_data['Product']}")
        doc.add_paragraph(f"**Project:** {report_data['Project']}")
        doc.add_paragraph(f"**Lot:** {report_data['Lot']}")
        doc.add_paragraph(f"**Released:** {report_data['Released']}")
        doc.add_paragraph()
        doc.add_paragraph(f"**Requested by:** {report_data['Requested by']}")
        doc.add_paragraph(f"**Performed by:** {report_data['Performed by']}")
        doc.add_paragraph(f"**Reviewed by:** {report_data['Reviewed by']}")
        doc.add_paragraph(f"**Approved by:** {report_data['Approved by']}")

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button(
            label="Baixar Relatório",
            data=buffer.getvalue(),
            file_name=OUTPUT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("Por favor, preencha todos os campos.")

st.markdown("""
---
**Instruções:**
1. Preencha todos os campos solicitados. Os primeiros quatro campos aparecerão na primeira linha e os quatro seguintes na segunda.
2. Certifique-se de que o arquivo de imagem `imagem_inicial.png` esteja no mesmo diretório deste script.
3. Clique em "Gerar Relatório Word" para baixar o documento com a imagem no início e as perguntas em negrito.
""")
