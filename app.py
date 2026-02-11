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

nome_coluna_abrev = "Abrev. [Artigos]"
nome_coluna_artigo = "Artigo [Documentos GC Lin]"
nome_coluna_codigo = "codigo_aba"

# garantir que os tipos batem
listagem[nome_coluna_artigo] = listagem[nome_coluna_artigo].astype(str)
componentes_dos_kits[nome_coluna_codigo] = componentes_dos_kits[nome_coluna_codigo].astype(str)

novas_linhas = []

for idx, row in listagem.iterrows():
    valor_artigo = row[nome_coluna_artigo]
    # índices das linhas em componentes_dos_kits com o mesmo codigo_aba
    idx_comp = componentes_dos_kits.index[componentes_dos_kits[nome_coluna_codigo] == valor_artigo]

    if len(idx_comp) > 0:
        for j in idx_comp:
            linha_comp = componentes_dos_kits.loc[j]
            # colunas 2:21 em R -> índices 1 a 20 em pandas
            for col_idx in range(1, 21):
                col_name = componentes_dos_kits.columns[col_idx]
                novo_valor = str(linha_comp[col_name])
                if pd.notna(novo_valor) and novo_valor.strip() != "":
                    nova_linha = row.copy()
                    nova_linha[nome_coluna_abrev] = novo_valor
                    novas_linhas.append(nova_linha)

# criar df_componentes_kits
if novas_linhas:
    df_componentes_kits = pd.concat(novas_linhas, axis=1).T.reset_index(drop=True)
else:
    df_componentes_kits = pd.DataFrame(columns=listagem.columns)

df_componentes_kits["Abrev. [Artigos]"] = (
    df_componentes_kits["Abrev. [Artigos]"]
    .replace("", pd.NA)
    .replace(" ", pd.NA)
)
df_componentes_kits = df_componentes_kits.dropna(subset=["Abrev. [Artigos]"])
st.write("### df_componentes_kits (equivalente ao loop em R)")
st.dataframe(df_componentes_kits)
