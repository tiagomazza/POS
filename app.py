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
    .str.contains("KIT", case=False, na=False)
].copy()

st.write("DF na listagem Kits")
st.dataframe(df_kits)

# 1) Ler revenda.xlsx
revenda = pd.read_excel("data/revenda.xlsx")
revenda["revenda"] = revenda["revenda"].astype(str)
listagem["Número [Clientes]"] = listagem["Número [Clientes]"].astype(str)

# 2) Juntar para identificar quais são de revenda
merged = listagem.merge(
    revenda[["revenda"]],
    left_on="Número [Clientes]",
    right_on="revenda",
    how="left",
    indicator=True,
)

# 3) Dataframe com as observações RETIRADAS (clientes que estão na lista de revenda)
retirados_revenda = merged[merged["_merge"] == "both"].drop(
    columns=["revenda", "_merge"],
    errors="ignore"
)

st.write("### Observações retiradas (Número [Clientes] presente em revenda)")
st.dataframe(retirados_revenda)

# 4) Atualizar df listagem mantendo apenas quem NÃO está em revenda
listagem = merged[merged["_merge"] == "left_only"].drop(
    columns=["revenda", "_merge"],
    errors="ignore"
)

st.write("### listagem após remover clientes de revenda")
st.dataframe(listagem)

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
    .replace("nan", pd.NA)
    .replace("nan", pd.NA)
)
df_componentes_kits = df_componentes_kits.dropna(subset=["Abrev. [Artigos]"])
st.write("### df_componentes_kits (equivalente ao loop em R)")
st.dataframe(df_componentes_kits)

# 1) Filtrar apenas linhas com KIT na listagem
kits_listagem = listagem[
    listagem["Descrição [Artigos]"]
    .astype(str)
    .str.contains("KIT", case=False, na=False)
].copy()

# 2) Garantir tipos compatíveis para o match
kits_listagem["Artigo [Documentos GC Lin]"] = kits_listagem["Artigo [Documentos GC Lin]"].astype(str)
componentes_dos_kits["codigo_aba"] = componentes_dos_kits["codigo_aba"].astype(str)

# 3) Anti-join: kits da listagem que NÃO têm correspondência em componentes_dos_kits
kits_sem_corresp = kits_listagem.merge(
    componentes_dos_kits[["codigo_aba"]],
    left_on="Artigo [Documentos GC Lin]",
    right_on="codigo_aba",
    how="left",
    indicator=True,
)

kits_sem_corresp = kits_sem_corresp[kits_sem_corresp["_merge"] == "left_only"].drop(
    columns=["codigo_aba", "_merge"],
    errors="ignore"
)

st.write("### Kits na listagem sem correspondência em componentes_dos_kits")
st.dataframe(kits_sem_corresp)

# 1) Remover da listagem todas as linhas que contenham "KIT"
mask_sem_kit_desc = (
    listagem["Descrição [Artigos]"].isna()
    | ~listagem["Descrição [Artigos]"].astype(str).str.contains("KIT", case=False, na=True)
)
mask_sem_kit_abrev = (
    listagem["Abrev. [Artigos]"].isna()
    | ~listagem["Abrev. [Artigos]"].astype(str).str.contains("KIT", case=False, na=True)
)

listagem = listagem[mask_sem_kit_desc & mask_sem_kit_abrev].copy()

st.write("### listagem após remover observações com KIT")
st.dataframe(listagem)

preco_custo = pd.read_excel("data/preço_custo.xlsx")
# garantir que sap é string
preco_custo["sap"] = preco_custo["sap"].astype(str)
df_componentes_kits["Abrev. [Artigos]"] = df_componentes_kits["Abrev. [Artigos]"].astype(str)

# 2) Fazer o merge (left join) para trazer preço_custo
df_componentes_kits = df_componentes_kits(
    preco_custo[["sap", "preço_custo"]],
    left_on="Abrev. [Artigos]",
    right_on="sap",
    how="left",
)

# 3) Copiar o valor de preço_custo para a coluna Úl.Pr.Cmp.
#    (se ainda não existir, será criada agora)
df_componentes_kits["Úl.Pr.Cmp."] = df_componentes_kits["preço_custo"]

# 4) (Opcional) limpar colunas auxiliares
df_componentes_kits = df_componentes_kits.drop(columns=["sap", "preço_custo"], errors="ignore")

# 5) (Opcional) garantir numérico
df_componentes_kits["Úl.Pr.Cmp."] = pd.to_numeric(df_componentes_kits["Úl.Pr.Cmp."], errors="coerce").fillna(0.0)

st.write("### kits após join com preço_custo")
st.dataframe(df_componentes_kits)

# 2) Adicionar as observações de df_componentes_kits
if not df_componentes_kits.empty:
    listagem = pd.concat([listagem, df_componentes_kits], ignore_index=True)

st.write("### listagem após adicionar df_componentes_kits")
st.dataframe(listagem)


# garantir que a coluna existe; ajusta o nome se for "Úl.Pr.Cmp. [Artigos]"
col_custo = "Úl.Pr.Cmp."

# se ainda não for numérico, opcional:
listagem[col_custo] = pd.to_numeric(listagem[col_custo], errors="coerce")

# df apenas com linhas SEM valor em Úl.Pr.Cmp.
df_sem_custo = listagem[listagem[col_custo].isna()].copy()

st.write("### Observações sem Úl.Pr.Cmp.")
st.dataframe(df_sem_custo)