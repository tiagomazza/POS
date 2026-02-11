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

# ============== 3. Export limpo inicial ==============
st.write("### 📥 Exportar listagem limpa (após limpeza das colunas)")
buffer_limpo = BytesIO()
listagem.to_excel(buffer_limpo, index=False, engine="openpyxl")
buffer_limpo.seek(0)

st.download_button(
    label="📥 Download listagem_limpa.xlsx",
    data=buffer_limpo,
    file_name="listagem_limpa.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ============== 4. Filtro Fatura / KENNA ==============
mask_tipo = listagem["Descrição [Tipos de Documentos]"] == "Fatura"
mask_familia = listagem["Família [Artigos]"] == "KENNA"
listagem = listagem[mask_tipo & mask_familia].copy()

st.write("### listagem filtrada (Fatura & KENNA)")
st.dataframe(listagem)

# Remover colunas com nome NA ou vazio
listagem = listagem.loc[:, ~listagem.columns.isna() & (listagem.columns != "")]
st.write("### listagem após remover colunas vazias/NA")
st.dataframe(listagem)

# ============== 5. Separar Kits ==============
df_kits = listagem[
    listagem["Descrição [Artigos]"].notna()
    & listagem["Descrição [Artigos]"].str.contains("KIT", case=False, na=False)
].copy()

st.write("### df_kits (li
