import streamlit as st
from docx import Document
from io import BytesIO

OUTPUT_FILENAME = 'relatorio_gerado.docx'

def generate_word_report(qa_pairs):
    doc = Document()
    for question, answer in qa_pairs.items():
        doc.add_paragraph(f"**Pergunta:** {question}")
        doc.add_paragraph(f"**Resposta:** {answer}")
        doc.add_paragraph() # Adiciona um espaço entre os pares

    # Salvar o documento na memória
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.title("Gerador de Relatório Word Dinâmico")

st.subheader("Insira as Perguntas e Respostas:")

qa_pairs = {}
num_qa_pairs = st.number_input("Número de Perguntas/Respostas:", min_value=1, step=1, value=1)

for i in range(num_qa_pairs):
    question = st.text_input(f"Pergunta {i+1}:")
    answer = st.text_area(f"Resposta para a Pergunta {i+1}:")
    if question:
        qa_pairs[question] = answer

if st.button("Gerar Relatório Word"):
    if qa_pairs:
        word_buffer = generate_word_report(qa_pairs)
        if word_buffer:
            st.download_button(
                label="Baixar Relatório",
                data=word_buffer.getvalue(),
                file_name=OUTPUT_FILENAME,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("Por favor, insira pelo menos uma Pergunta e Resposta.")

st.markdown("""
---
**Instruções:**
1. Insira o número de pares de Pergunta/Resposta que você deseja adicionar.
2. Preencha os campos de cada Pergunta e sua respectiva Resposta.
3. Clique em "Gerar Relatório Word" para baixar o documento.
""")
