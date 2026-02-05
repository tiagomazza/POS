import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Processamento de POS – KENNA")

# 1. Upload do listagem.xls / .xlsx
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

# Ler o ficheiro
listagem = pd.read_excel(uploaded_file, header=None, engine=engine)

# Ajustar cabeçalho (supondo que os nomes das colunas estão na linha 5)
listagem.columns = listagem.iloc[4].astype(str).values
listagem = listagem.iloc[5:, :]

# Normalizar nomes de coluna
listagem.columns = (
    listagem.columns
    .astype(str)
    .str.strip()
    .str.replace("  ", " ")
)

# Debug: mostrar colunas atuais
st.write("### Colunas atuais")
st.write(listagem.columns.tolist())

# Filtrar apenas "Fatura" e "KENNA"
mask_tipo = listagem["Descrição [Tipos de Documentos]"] == "Fatura"
mask_familia = listagem["Família [Artigos]"] == "KENNA"
listagem = listagem[mask_tipo & mask_familia].copy()

# Remover colunas com nome NA ou vazio
listagem = listagem.loc[:, ~listagem.columns.isna() & (listagem.columns != "")]

# Separar kits
df_kits = listagem[
    listagem["Descrição [Artigos]"].notna()
    & listagem["Descrição [Artigos]"].str.contains("KIT", case=False, na=False)
].copy()

# Ler componentes_kits.xlsx da pasta data
componentes_kits = pd.read_excel("data/componentes_kits.xlsx")
nome_coluna_abrev = "Abrev. [Artigos]"
nome_coluna_artigo = "Artigo [Documentos GC Lin]"
nome_coluna_codigo = "codigo_aba"

# Loop para expandir componentes dos kits
novas_linhas = []

for idx, row in listagem.iterrows():
    valor_artigo = row[nome_coluna_artigo]
    idx_comp = componentes_kits[componentes_kits[nome_coluna_codigo] == valor_artigo].index

    for j in idx_comp:
        linha_comp = componentes_kits.loc[j]
        for col_idx in range(1, 21):  # colunas 2 a 21 (índice 1 a 20)
            col_name = componentes_kits.columns[col_idx]
            novo_valor = str(linha_comp[col_name])
            if pd.notna(novo_valor) and novo_valor.strip() != "":
                nova_linha = row.copy()
                nova_linha[nome_coluna_abrev] = novo_valor
                novas_linhas.append(nova_linha)

# Criar df_componentes_kits
if novas_linhas:
    df_componentes_kits = pd.concat(novas_linhas, axis=1).T.reset_index(drop=True)
else:
    df_componentes_kits = pd.DataFrame(columns=listagem.columns)

# Ler preço_custo.xlsx
preco_custo = pd.read_excel("data/preço_custo.xlsx")
preco_custo["sap"] = preco_custo["sap"].astype(str)

# Fazer join com preço de custo
if not df_componentes_kits.empty:
    df_componentes_kits = df_componentes_kits.merge(
        preco_custo[["sap", "preço_custo"]],
        left_on=nome_coluna_abrev,
        right_on="sap",
        how="left",
    )
    df_componentes_kits["Úl.Pr.Cmp. [Artigos]"] = df_componentes_kits["preço_custo"]
    df_componentes_kits = df_componentes_kits.drop(columns=["sap", "preço_custo"], errors="ignore")

# Remover kits da listagem original
mask_sem_kit_desc = (
    listagem["Descrição [Artigos]"].isna()
    | ~listagem["Descrição [Artigos]"].str.contains("KIT", case=False, na=True)
)
mask_sem_kit_abrev = (
    listagem["Abrev. [Artigos]"].isna()
    | ~listagem["Abrev. [Artigos]"].str.contains("KIT", case=False, na=True)
)
listagem = listagem[mask_sem_kit_desc & mask_sem_kit_abrev].copy()

# Kits sem correspondência em componentes_kits
kits_sem_corresp = df_kits.merge(
    componentes_kits[[nome_coluna_codigo]],
    left_on="Artigo [Documentos GC Lin]",
    right_on=nome_coluna_codigo,
    how="left",
    indicator=True,
)
kits_sem_corresp = kits_sem_corresp[kits_sem_corresp["_merge"] == "left_only"]
kits_sem_corresp = (
    kits_sem_corresp.groupby("Artigo [Documentos GC Lin]", as_index=False)
    .size()
    .rename(columns={"size": "qtd"})
)

# Adicionar componentes dos kits à listagem
if not df_componentes_kits.empty:
    listagem = pd.concat([listagem, df_componentes_kits], ignore_index=True)

# Ler revenda.xlsx
revenda_lista = pd.read_excel("data/revenda.xlsx")
revenda_lista["revenda"] = revenda_lista["revenda"].astype(str)

# Adicionar coluna revenda
listagem["revenda"] = None

# Fazer join com revenda
if "revenda" in listagem.columns:
    listagem = listagem.merge(
        revenda_lista[["revenda"]],
        left_on="Número [Clientes]",
        right_on="revenda",
        how="left",
        indicator=True,
    )
    listagem = listagem[listagem["_merge"] == "left_only"].drop(columns=["revenda", "_merge"], errors="ignore")
else:
    listagem = listagem[listagem["revenda"].isna()].drop(columns=["revenda"], errors="ignore")

# Limpar linhas totalmente vazias
listagem = listagem.dropna(how="all").copy()

# Tratar Abrev. [Artigos] e Úl.Pr.Cmp. [Artigos]
if "Abrev. [Artigos]" in listagem.columns:
    listagem["Abrev. [Artigos]"] = (
        listagem["Abrev. [Artigos]"].astype(str).str.slice(0, 7)
    )

if "Úl.Pr.Cmp. [Artigos]" in listagem.columns:
    listagem["Úl.Pr.Cmp. [Artigos]"] = pd.to_numeric(
        listagem["Úl.Pr.Cmp. [Artigos]"], errors="coerce"
    )
    listagem["Úl.Pr.Cmp. [Artigos]"] = listagem["Úl.Pr.Cmp. [Artigos]"].fillna(0.0)

# Ler POS_ABA.xls (se precisares de algo específico dela, ajusta aqui)
pos_aba = pd.read_excel("data/POS_ABA.xls")  # opcional

# Criar POS
POS = listagem.assign(
    **{
        "Distributor SAP Acct #": 70465299,
        "Customer Ship To Country": "PT",
        "Customer Ship To Zip Code": listagem["Cód.Postal [Clientes]"],
        "SAP Material Master No.": listagem["Abrev. [Artigos]"],
        "ANSI Catalog No./Grade Item Number": "",
        "Qty Sold": listagem["Quant [Documentos GC Lin]"],
        "Invoice Date": listagem["Data"],
        "Deal Registration ID": "",
        "Total Distributor Cost": listagem["Úl.Pr.Cmp. [Artigos]"].round(2),
    }
)

POS = POS.dropna(subset=["Customer Ship To Zip Code"])

# Permitir download do POS final
buffer = BytesIO()
POS.to_excel(buffer, index=False, engine="openpyxl")
buffer.seek(0)

st.write("### POS pronto para download")
st.dataframe(POS)

st.download_button(
    label="📥 Download POS_pronta.xlsx",
    data=buffer,
    file_name="POS_pronta.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
