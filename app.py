import streamlit as st
import pandas as pd
from io import BytesIO

st.title("POS – KENNAMETAL")

# =========================
# Botão de debug
# =========================
debug = st.checkbox("🐞")

# =========================
# 1. Upload e leitura base
# =========================
uploaded_file = st.file_uploader(
    "Carregar ficheiros *.xls ou *.xlsx",
    type=["xls", "xlsx"]
)
if uploaded_file is None:
    st.info("Por favor, carregue o ficheiro *.xls ou *.xlsx.")
    st.stop()

engine = "xlrd" if uploaded_file.name.endswith(".xls") else "openpyxl"

listagem = pd.read_excel(uploaded_file, header=None, engine=engine)
listagem.columns = listagem.iloc[5].astype(str).values
listagem = listagem.iloc[6:, :]

# normalizar nomes de coluna
listagem.columns = (
    listagem.columns
    .astype(str)
    .str.strip()
    .str.replace("  ", " ")
)

# =========================
# 1.1. Validação de colunas necessárias
# =========================
colunas_necessarias = [
    "Descrição [Tipos de Documentos]",
    "Família [Artigos]",
    "Descrição [Artigos]",
    "Artigo [Documentos GC Lin]",
    "Abrev. [Artigos]",
    "Número [Clientes]",
    "Cód.Postal [Clientes]",
    "Quant [Documentos GC Lin]",
    "Data",
]

colunas_presentes = listagem.columns.astype(str).tolist()
faltantes = [c for c in colunas_necessarias if c not in colunas_presentes]

if faltantes:
    st.error(
        "O ficheiro carregado não contém todas as colunas necessárias para o processamento.\n\n"
        "Colunas em falta:\n- " + "\n- ".join(faltantes)
    )
    st.stop()
else:
    if debug:
        st.success("✔ O ficheiro contém todas as colunas necessárias para o processo.")
        st.write("Colunas detectadas:")
        st.write(listagem.columns.tolist())

# =========================
# 2. Filtros base (Fatura / KENNA) e limpeza
# =========================
listagem = listagem[listagem["Descrição [Tipos de Documentos]"] == "Fatura"].copy()
listagem = listagem.drop(columns=["Descrição [Tipos de Documentos]"], errors="ignore")
listagem = listagem[listagem["Família [Artigos]"] == "KENNA"].copy()
listagem = listagem.dropna(axis=1, how="all")

if debug:
    st.write("### 🧹 listagem após limpeza inicial (Fatura / KENNA)")
    st.dataframe(listagem)

# =========================
# 3. Remover clientes de revenda
# =========================
revenda = pd.read_excel("data/revenda.xlsx")
revenda["revenda"] = revenda["revenda"].astype(str)
listagem["Número [Clientes]"] = listagem["Número [Clientes]"].astype(str)

merged = listagem.merge(
    revenda[["revenda"]],
    left_on="Número [Clientes]",
    right_on="revenda",
    how="left",
    indicator=True,
)

retirados_revenda = merged[merged["_merge"] == "both"].drop(
    columns=["revenda", "_merge"],
    errors="ignore"
)

if debug:
    st.write("### 🔎 Revendas encontradas")
    st.dataframe(retirados_revenda)

listagem = merged[merged["_merge"] == "left_only"].drop(
    columns=["revenda", "_merge"],
    errors="ignore"
)

if debug:
    st.write("### 🟢 listagem após remover clientes de revenda")
    st.dataframe(listagem)

# =========================
# 4. Identificar KITS na listagem
# =========================
df_kits = listagem[
    listagem["Descrição [Artigos]"]
    .astype(str)
    .str.contains("KIT", case=False, na=False)
].copy()

if debug:
    st.write("### 🔎 Kits encontrados na listagem")
    st.dataframe(df_kits)

# =========================
# 5. Ler componentes dos kits e decompor
# =========================
componentes_dos_kits = pd.read_excel("data/componentes_kits.xlsx")

if debug:
    st.write("### 🧩 componentes dos kits (data/componentes_kits.xlsx)")
    st.dataframe(componentes_dos_kits)

nome_coluna_abrev = "Abrev. [Artigos]"
nome_coluna_artigo = "Artigo [Documentos GC Lin]"
nome_coluna_codigo = "codigo_aba"

listagem[nome_coluna_artigo] = listagem[nome_coluna_artigo].astype(str)
componentes_dos_kits[nome_coluna_codigo] = componentes_dos_kits[nome_coluna_codigo].astype(str)

novas_linhas = []

for idx, row in listagem.iterrows():
    valor_artigo = row[nome_coluna_artigo]
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

if novas_linhas:
    df_componentes_kits = pd.concat(novas_linhas, axis=1).T.reset_index(drop=True)
else:
    df_componentes_kits = pd.DataFrame(columns=listagem.columns)

# limpar Abrev vazios
df_componentes_kits["Abrev. [Artigos]"] = (
    df_componentes_kits["Abrev. [Artigos]"].replace("nan", pd.NA)
)
df_componentes_kits = df_componentes_kits.dropna(subset=["Abrev. [Artigos]"])

if debug:
    st.write("### 💥 Kits decompostos em componentes")
    st.dataframe(df_componentes_kits)

# =========================
# 6. Kits sem correspondência em componentes_dos_kits
# =========================
kits_listagem = df_kits.copy()
kits_listagem[nome_coluna_artigo] = kits_listagem[nome_coluna_artigo].astype(str)
componentes_dos_kits[nome_coluna_codigo] = componentes_dos_kits[nome_coluna_codigo].astype(str)

kits_sem_corresp = kits_listagem.merge(
    componentes_dos_kits[[nome_coluna_codigo]],
    left_on=nome_coluna_artigo,
    right_on=nome_coluna_codigo,
    how="left",
    indicator=True,
)

kits_sem_corresp = kits_sem_corresp[kits_sem_corresp["_merge"] == "left_only"].drop(
    columns=[nome_coluna_codigo, "_merge"],
    errors="ignore"
)

# ❌ Sempre visível
st.write("### ❌ Kits na listagem sem correspondência")
st.dataframe(kits_sem_corresp)

# =========================
# 7. Trazer preço de custo dos componentes dos kits
# =========================
preco_custo = pd.read_excel("data/preço_custo.xlsx")
preco_custo["sap"] = preco_custo["sap"].astype(str)
df_componentes_kits["Abrev. [Artigos]"] = df_componentes_kits["Abrev. [Artigos]"].astype(str)

if debug:
    st.write("### 🧩 tabela de preço de custo (preço_custo.xlsx)")
    st.dataframe(preco_custo)

df_componentes_kits = df_componentes_kits.merge(
    preco_custo[["sap", "preço_custo"]],
    left_on="Abrev. [Artigos]",
    right_on="sap",
    how="left",
)

df_componentes_kits["Úl.Pr.Cmp."] = df_componentes_kits["preço_custo"]
df_componentes_kits = df_componentes_kits.drop(columns=["sap", "preço_custo"], errors="ignore")
df_componentes_kits["Úl.Pr.Cmp."] = pd.to_numeric(
    df_componentes_kits["Úl.Pr.Cmp."], errors="coerce"
).fillna(0.0)

if debug:
    st.write("### 💰 Componentes de kits com preço de custo")
    st.dataframe(df_componentes_kits)

# =========================
# 8. Remover linhas de KIT da listagem base e adicionar componentes
# =========================
mask_sem_kit_desc = (
    listagem["Descrição [Artigos]"].isna()
    | ~listagem["Descrição [Artigos]"].astype(str).str.contains("KIT", case=False, na=True)
)
mask_sem_kit_abrev = (
    listagem["Abrev. [Artigos]"].isna()
    | ~listagem["Abrev. [Artigos]"].astype(str).str.contains("KIT", case=False, na=True)
)

listagem = listagem[mask_sem_kit_desc & mask_sem_kit_abrev].copy()

if debug:
    st.write("### 🟢 listagem após remover linhas de KIT")
    st.dataframe(listagem)

# alinhar nome de custo dos componentes com o da listagem
df_componentes_kits["Úl.Pr.Cmp. [Artigos]"] = df_componentes_kits["Úl.Pr.Cmp."]

# adicionar componentes de kits à listagem
if not df_componentes_kits.empty:
    listagem = pd.concat([listagem, df_componentes_kits], ignore_index=True)

if debug:
    st.write("### 🟢 listagem final com kits decompostos adicionados")
    st.dataframe(listagem)

# =========================
# 9. Identificar artigos sem último preço de compra
# =========================
listagem["Úl.Pr.Cmp. [Artigos]"] = pd.to_numeric(
    listagem["Úl.Pr.Cmp. [Artigos]"], errors="coerce"
)

listagem_sem_custos = listagem[listagem["Úl.Pr.Cmp. [Artigos]"].isna()].copy()

# 🟡 Sempre visível
st.write("### 🟡 Artigos sem último preço de compra")
st.dataframe(listagem_sem_custos)

# =========================
# 10. Construir POS final
# =========================
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

cols_pos = [
    "Distributor SAP Acct #",
    "Customer Ship To Country",
    "Customer Ship To Zip Code",
    "SAP Material Master No.",
    "ANSI Catalog No./Grade Item Number",
    "Qty Sold",
    "Invoice Date",
    "Deal Registration ID",
    "Total Distributor Cost",
]

POS = POS[cols_pos].copy()
POS = POS.dropna(subset=["Customer Ship To Zip Code"])

# ❇️ Sempre visível
st.write("### ❇️ POS terminada")
st.dataframe(POS)

# =========================
# 11. Exportar POS em XLSX
# =========================
buffer_pos = BytesIO()
POS.to_excel(buffer_pos, index=False, engine="openpyxl")
buffer_pos.seek(0)

st.download_button(
    label="📥 Download_POS.xlsx",
    data=buffer_pos,
    file_name="POS_pronta.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
debug = st.checkbox("🐞 debug (mostrar passos intermédios)")