import streamlit as st
import pandas as pd
import re
from PyPDF2 import PdfReader
import io

st.set_page_config(page_title="Zelus - PDF para Excel", layout="centered")
st.title("Zelus RH - Extração de Dados do PDF")
st.write("Envie um PDF com informações de colaboradores para extrair nome e data de admissão.")

uploaded_file = st.file_uploader("Escolha o arquivo PDF", type="pdf")

def extrair_dados_por_blocos(file):
    reader = PdfReader(file)
    texto = ""
    for page in reader.pages:
        texto += page.extract_text() + "\n"

    blocos = texto.split("Cód:")
    nomes, admissoes = [], []

    for bloco in blocos:
        if "Admissão:" in bloco:
            nome_match = re.search(r"[A-ZÀ-ÿ\s]{5,}[0-9]{2,3}", bloco)
            admissao_match = re.search(r"Admissão:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", bloco)

            if nome_match and admissao_match:
                nome_raw = nome_match.group(0)
                nome_limpo = re.sub(r"\d+$", "", nome_raw).strip()
                nomes.append(nome_limpo)
                admissoes.append(admissao_match.group(1))

    return pd.DataFrame({"Nome": nomes, "Admissão": admissoes})

if uploaded_file:
    try:
        df = extrair_dados_por_blocos(uploaded_file)
        if df.empty:
            st.warning("Nenhum dado foi encontrado no PDF.")
        else:
            st.success("Dados extraídos com sucesso!")
            st.dataframe(df)

            buffer = io.BytesIO()
            df.to_excel(buffer, index=False, engine='openpyxl')
            st.download_button("📥 Baixar Excel", data=buffer.getvalue(),
                               file_name="dados_extraidos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Erro ao processar o PDF: {str(e)}")
