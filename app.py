import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Processamento de POS – KENNA")

uploaded_file = st.file_uploader("Carregar listagem.xlsx", type=["xlsx"])
if uploaded_file is None:
    st.info("Por favor, carregue o ficheiro listagem.xlsx.")
    st.stop()

# Ler o ficheiro sem header
listagem = pd.read_excel(uploaded_file, header=None, engine="openpyxl")

st.write("### Primeiras 5 linhas do ficheiro original (sem cabeçalho)")
st.dataframe(listagem.head())

# Ajuste: escolhe a linha que contém os cabeçalhos (por exemplo linha 5, índice 4)
# Se os nomes das colunas estiverem na linha 5, faz:
listagem.columns = listagem.iloc[4].astype(str).values  # linha 5 como cabeçalho
listagem = listagem.iloc[5:, :]  # dados a partir da linha 6

# Normalizar nomes de coluna
listagem.columns = (
    listagem.columns
    .astype(str)
    .str.strip()
    .str.replace("  ", " ")
)

st.write("### Colunas atuais após ajuste")
st.write(listagem.columns.tolist())

# Exportar para xlsx (botão de download)
buffer = BytesIO()
listagem.to_excel(buffer, index=False, engine="openpyxl")
buffer.seek(0)

st.download_button(
    label="📥 Exportar listagem ajustada para .xlsx",
    data=buffer,
    file_name="listagem_ajustada.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
