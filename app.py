import streamlit as st
import pandas as pd
from io import BytesIO

st.title("POS – KENNAMETAL")

# =========================
# Botão de debug
# =========================
debug = st.checkbox("👾")

# =========================
# 1. Upload e leitura base
# =========================
uploaded_file = st.file_uploader(
    "",
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
# 2. Filtros base
# =========================
listagem = listagem[listagem["Descrição [Tipos de Documentos]"] == "Fatura"].copy()
listagem = listagem.drop(columns=["Descrição [Tipos de Documentos]"], errors="ignore")
listagem = listagem[listagem["Família [Artigos]"] == "KENNA"].copy()
listagem = listagem.dropna(axis=1, how="all")

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

listagem = merged[merged["_merge"] == "left_only"].drop(
    columns=["revenda", "_merge"],
    errors="ignore"
)

# =========================
# 4. Identificar KITS
# =========================
df_kits = listagem[
    listagem["Descrição [Artigos]"]
    .astype(str)
    .str.contains("KIT", case=False, na=False)
].copy()

# =========================
# 5. Decompor KITS
# =========================
componentes_dos_kits = pd.read_excel("data/componentes_kits.xlsx")

nome_coluna_abrev = "Abrev. [Artigos]"
nome_coluna_artigo = "Artigo [Documentos GC Lin]"
nome_coluna_codigo = "codigo_aba"

listagem[nome_coluna_artigo] = listagem[nome_coluna_artigo].astype(str)
componentes_dos_kits[nome_coluna_codigo] = componentes_dos_kits[nome_coluna_codigo].astype(str)

novas_linhas = []

for _, row in listagem.iterrows():
    valor_artigo = row[nome_coluna_artigo]
    idx_comp = componentes_dos_kits.index[
        componentes_dos_kits[nome_coluna_codigo] == valor_artigo
    ]

    if len(idx_comp) > 0:
        for j in idx_comp:
            linha_comp = componentes_dos_kits.loc[j]
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

df_componentes_kits["Abrev. [Artigos]"] = (
    df_componentes_kits["Abrev. [Artigos]"].replace("nan", pd.NA)
)
df_componentes_kits = df_componentes_kits.dropna(subset=["Abrev. [Artigos]"])

# =========================
# 6. Preço de custo
# =========================
preco_custo = pd.read_excel("data/preço_custo.xlsx")
preco_custo["sap"] = preco_custo["sap"].astype(str)
df_componentes_kits["Abrev. [Artigos]"] = df_componentes_kits["Abrev. [Artigos]"].astype(str)

df_componentes_kits = df_componentes_kits.merge(
    preco_custo[["sap", "preço_custo"]],
    left_on="Abrev. [Artigos]",
    right_on="sap",
    how="left",
)

df_componentes_kits["Úl.Pr.Cmp."] = df_componentes_kits["preço_custo"]
df_componentes_kits = df_componentes_kits.drop(columns=["sap", "preço_custo"], errors="ignore")

# ⚠️ IMPORTANTE: NÃO fazemos fillna aqui (como pediste)

# =========================
# 7. Remover kits e juntar componentes
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

df_componentes_kits["Úl.Pr.Cmp. [Artigos]"] = df_componentes_kits["Úl.Pr.Cmp."]

if not df_componentes_kits.empty:
    listagem = pd.concat([listagem, df_componentes_kits], ignore_index=True)

# =========================
# 8. POS final
# =========================
listagem["Úl.Pr.Cmp. [Artigos]"] = pd.to_numeric(
    listagem["Úl.Pr.Cmp. [Artigos]"], errors="coerce"
)

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
        # ✅ AQUI é onde tratamos os NaN
        "Total Distributor Cost": listagem["Úl.Pr.Cmp. [Artigos]"]
            .fillna(0)
            .round(2),
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

st.write("### ❇️ POS terminada")
st.dataframe(POS)

# =========================
# 9. Exportar
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