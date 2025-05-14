import streamlit as st
from docx import Document
from io import BytesIO

OUTPUT_FILENAME = 'relatorio_gerado.docx'

def generate_word_report(data):
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

    # Salvar o documento na memória
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.title("Gerador de Relatório Word")

st.subheader("Preencha as informações:")

product = st.text_input("Product:")
project = st.text_input("Project:")
lot = st.text_input("Lot:")
released = st.text_input("Released:")
requested_by = st.text_input("Requested by:")
performed_by = st.text_input("Performed by:")
reviewed_by = st.text_input("Reviewed by:")
approved_by = st.text_input("Approved by:")

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
        word_buffer = generate_word_report(report_data)
        if word_buffer:
            st.download_button(
                label="Baixar Relatório",
                data=word_buffer.getvalue(),
                file_name=OUTPUT_FILENAME,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Por favor, preencha todos os campos.")

st.markdown("""
---
**Instruções:**
1. Preencha todos os campos solicitados.
2. Clique em "Gerar Relatório Word" para baixar o documento.
""")
