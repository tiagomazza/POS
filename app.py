import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Processamento de POS – KENNA")

# ============== 1. Upload ==============
uploaded_file = st.file_uploader(
    "Carregar listagem.xls ou listagem.xlsx",
    type=["xls", "xlsx"]
)
if uploaded_file is None:
    st.info("Por favor, carregue o ficheiro listagem.xls ou listagem.xlsx.")
    st.stop()

# Escolher engine conforme extensão
if uploaded_file.name.endswith(".xls"):
    engine = "xlrd"
else:
    engine = "openpyxl"

# Ler o ficheiro bruto
listagem = pd.read_excel(uploaded_file, header=None, engine=engine)
st.write("### listagem bruta (header=None)")
st.dataframe(listagem)

# ============== 2. Ajuste de cabeçalho ==============
# Supondo que os nomes das colunas estão na linha 6 (índice 5)
listagem.columns = listagem.iloc[5].astype(str).values
listagem = listagem.iloc[6:, :]

# Normalizar nomes de coluna
listagem.columns = (
    listagem.columns
    .astype(str)
    .str.strip()
    .str.replace("  ", " ")
)

st.write("### Colunas atuais após normalização")
st.write(listagem.columns.tolist())

st.write("### listagem após ajuste de cabeçalho")
st.dataframe(listagem)

