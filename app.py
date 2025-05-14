import streamlit as st
from docx import Document
from io import BytesIO

# Nome do arquivo de template Word (deve estar no mesmo diretório)
TEMPLATE_PATH = 'template_relatorio.docx'
OUTPUT_FILENAME = 'relatorio_gerado.docx'

def generate_word_report(data):
    try:
        doc = Document(TEMPLATE_PATH)
        for para in doc.paragraphs:
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in para.text:
                    para.text = para.text.replace(placeholder, value)

        # Salvar o documento na memória
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    except FileNotFoundError:
        st.error(f"Arquivo de template '{TEMPLATE_PATH}' não encontrado.")
        return None
    except Exception as e:
        st.error(f"Ocorreu um erro ao gerar o relatório: {e}")
        return None

st.title("Gerador de Relatório Word")

st.subheader("Insira os dados para o relatório:")

produto = st.text_input("Produto:")
lote = st.text_input("Lote:")
data = st.date_input("Data:")
observacoes = st.text_area("Observações:")

report_data = {
    "Produto": produto,
    "Lote": lote,
    "Data": data.strftime('%d/%m/%Y') if data else "",
    "Observacoes": observacoes
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
1. Certifique-se de ter um arquivo chamado `template_relatorio.docx` no mesmo diretório deste script.
2. Preencha os campos acima.
3. Clique em "Gerar Relatório Word" para baixar o documento preenchido.
""")
