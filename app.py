import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Processamento de POS – KENNA")

uploaded_file = st.file_uploader(
    "Carregar listagem.xls ou listagem.xlsx",
    type=["xls", "xlsx"]
)
if uploaded_file is None:
    st.info("Por favor, carregue o ficheiro *.xls ou *.xlsx.")
    st.stop()

if uploaded_file.name.endswith(".xls"):
    engine = "xlrd"
else:
    engine = "openpyxl"

listagem = pd.read_excel(uploaded_file, header=None, engine=engine)
listagem.columns = listagem.iloc[5].astype(str).values
listagem = listagem.iloc[6:, :]

listagem.columns = (
    listagem.columns
    .astype(str)
    .str.strip()
    .str.replace("  ", " ")
)

listagem = listagem[listagem["Descrição [Tipos de Documentos]"] == "Fatura"].copy()
listagem = listagem[listagem["Família [Artigos]"] == "KENNA"].copy()

st.write("### listagem após ajuste e filtros")
st.dataframe(listagem)

df_kits = listagem[
    listagem["Descrição [Artigos]"]
    .astype(str)              
    .str.contains("TORNO", case=False, na=False)
].copy()

st.write("DF na listagem Kits")
st.dataframe(df_kits)

