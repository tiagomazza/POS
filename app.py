import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Processamento de POS – KENNA")

uploaded_file = st.file_uploader(
    "Carregar listagem.xls ou listagem.xlsx",
    type=["xls", "xlsx"]
)
if uploaded_file is None:
    st.info("Por favor, carregue o ficheiro listagem.xls ou listagem.xlsx.")
    st.stop()

if uploaded_file.name.endswith(".xls"):
    engine = "xlrd"
else:
    engine = "openpyxl"

listagem = pd.read_excel(uploaded_file, header=None, engine=engine)
st.write("### listagem bruta (header=None)")
st.dataframe(listagem)

# Ajustar cabeçalho
listagem.columns = listagem.iloc[5].astype(str).values
listagem = listagem.iloc[6:, :]

# Normalizar nomes de coluna
listagem.columns = (
    listagem.columns
    .astype(str)
    .str.strip()
    .str.replace("  ", " ")
)

# 🔹 Manter apenas linhas onde Descrição [Tipos de Documentos] == "Fatura"
listagem = listagem[listagem["Descrição [Tipos de Documentos]"] == "Fatura"].copy()

st.write("### listagem após ajuste e filtro (apenas Fatura)")
st.dataframe(listagem)
