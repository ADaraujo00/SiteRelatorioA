import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
from PIL import Image

OUTPUT_FILENAME = 'relatorio_com_tabela_alternativa.docx'

st.title("Gerador de Relatório com Tabela e Imagens")

st.subheader("Detalhes das Amostras:")

sample_details_input = {}
questions = ["Product", "Accessories", "Model", "Voltage", "Dimensions", "Supplier", "Quantity", "Volume"]

for question in questions:
    col1, col2 = st.columns([1, 2])
    col1.markdown(f"**{question}:**")
    sample_details_input[question] = col2.text_input("", label_visibility="collapsed", key=question)

st.subheader("Carregar Imagens:")
uploaded_images = st.file_uploader("Selecione as imagens:", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

def generate_word_report(details, images):
    doc = Document()

    doc.add_paragraph("Detalhes das Amostras:")
    table = doc.add_table(rows=len(details), cols=2)
    for i, (key, value) in enumerate(details.items()):
        cell_left = table.cell(i, 0)
        cell_right = table.cell(i, 1)
        cell_left.text = f"{key}:"
        cell_left.paragraphs[0].runs[0].font.bold = True
        cell_right.text = value

    doc.add_paragraph("\nImagens:")
    if images:
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
    if all(sample_details_input.values()):
        word_buffer = generate_word_report(sample_details_input, uploaded_images)
        if word_buffer:
            st.download_button(
                label="Baixar Relatório Word",
                data=word_buffer.getvalue(),
                file_name=OUTPUT_FILENAME,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Por favor, preencha todos os detalhes das amostras.")

st.markdown("""
---
**Instruções:**
1. Preencha os detalhes das amostras nos campos ao lado de cada pergunta.
2. Carregue as imagens desejadas.
3. Clique em "Gerar Relatório Word" para baixar o documento com a tabela de detalhes e as imagens.
""")
