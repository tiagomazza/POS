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

componentes_dos_kits = pd.read_excel("data/componentes_kits.xlsx")

st.write("### componentes dos kits (data/componentes_kits.xlsx)")
st.dataframe(componentes_dos_kits)

# df_kits já existe
# componentes_dos_kits já existe (lido de data/componentes_kits.xlsx)

# garantir tipos compatíveis para o join
df_kits["Número [Artigos]"] = df_kits["Número [Artigos]"].astype(str)
componentes_dos_kits["codigo_aba"] = componentes_dos_kits["codigo_aba"].astype(str)

novas_linhas = []

for idx, row in df_kits.iterrows():
    codigo = row["Número [Artigos]"]

    # procurar linha(s) correspondente(s) no componentes_dos_kits
    comp_rows = componentes_dos_kits[componentes_dos_kits["codigo_aba"] == codigo]

    for _, comp in comp_rows.iterrows():
        # iterar pelas colunas sap_1 ... sap_10
        for i in range(1, 11):
            col_sap = f"sap_{i}"
            if col_sap in comp.index:
                valor_sap = str(comp[col_sap]).strip()
                if valor_sap and valor_sap.lower() != "nan":
                    nova_obs = row.copy()
                    nova_obs["Abrev. [Artigos]"] = valor_sap
                    novas_linhas.append(nova_obs)

# criar novo df com os componentes dos kits (linhas “explodidas”)
if novas_linhas:
    df_componentes_kits = pd.DataFrame(novas_linhas).reset_index(drop=True)
else:
    df_componentes_kits = pd.DataFrame(columns=df_kits.columns)

st.write("### componentes dos kits (a partir de df_kits + componentes_dos_kits)")
st.dataframe(df_componentes_kits)
