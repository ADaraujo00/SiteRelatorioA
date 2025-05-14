import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
from PIL import Image

OUTPUT_FILENAME = 'relatorio_amostras.docx'

st.title("Descrição de Amostras")

st.subheader("Imagens das Amostras:")
image_files = []
for i in range(12):
    uploaded_file = st.file_uploader(f"Imagem {i+1}:", type=["jpg", "jpeg", "png"])
    if uploaded_file is not None:
        image_files.append(uploaded_file)
    else:
        image_files.append(None)

st.subheader("Detalhes das Amostras:")
product = st.text_input("Product:")
accessories = st.text_input("Accessories:")
model = st.text_input("Model:")
voltage = st.text_input("Voltage:")
dimensions = st.text_input("Dimensions:")
supplier = st.text_input("Supplier:")
quantity = st.text_input("Quantity:")
volume = st.text_input("Volume:")

sample_details = {
    "Product": product,
    "Accessories": accessories,
    "Model": model,
    "Voltage": voltage,
    "Dimensions": dimensions,
    "Supplier": supplier,
    "Quantity": quantity,
    "Volume": volume,
}

def generate_word_report(images, details):
    doc = Document()
    doc.add_paragraph("3. SAMPLES DESCRIPTION:")

    # Adicionar as imagens em layout de grade (3x4)
    for i in range(0, len(images), 4):
        row = doc.add_paragraph()
        for j in range(4):
            if i + j < len(images) and images[i + j] is not None:
                try:
                    img = Image.open(images[i + j])
                    # Ajustar o tamanho da imagem conforme necessário
                    width, height = img.size
                    ratio = height / width
                    new_width = 2.5  # Polegadas
                    new_height = new_width * ratio
                    row.add_run().add_picture(images[i + j], width=Inches(new_width), height=Inches(new_height))
                except Exception as e:
                    doc.add_paragraph(f"Erro ao adicionar imagem {i+j+1}: {e}")
            row.add_run("   ") # Adiciona algum espaço entre as imagens
        doc.add_paragraph() # Espaço entre as linhas de imagens

    # Adicionar a tabela com os detalhes
    table = doc.add_table(rows=len(details), cols=2)
    for i, (key, value) in enumerate(details.items()):
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
    word_buffer = generate_word_report(image_files, sample_details)
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
1. Carregue até 12 imagens. Se menos forem carregadas, o layout se ajustará.
2. Preencha os detalhes das amostras na tabela.
3. Clique em "Gerar Relatório Word" para baixar o documento.
""")
