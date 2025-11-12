# DASHBOARD COMERCIAL BRASFORMA — v25
# Upgrades:
# - Leitura das bases direto do GitHub (raw) com fallback para upload local
# - Suporte a repositório privado via GITHUB_TOKEN (opcional)
# - Correção robusta da agregação de Metas por hierarquia (Regional/Representante)
# - Mantém v24: alíquotas efetivas por SKU no simulador, filtros, logo, etc.

import os
import io
import requests
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from PIL import Image
from fpdf import FPDF

# ==========================
# CONFIG GERAL
# ==========================
LOGO_PATHS = ["images.png", "logo_brasforma.png"]
PAGE_TITLE = "Brasforma — Dashboard Comercial v25"
st.set_page_config(page_title=PAGE_TITLE, layout="wide", page_icon=LOGO_PATHS[0])

def _load_logo():
    for p in LOGO_PATHS:
        try: return Image.open(p)
        except: continue
    return None

APP_LOGO = _load_logo()
with st.container():
    c1, c2 = st.columns([1, 8])
    if APP_LOGO is not None:
        c1.image(APP_LOGO, use_container_width=True)
    c2.markdown("## **Dashboard Comercial — Brasforma**")
    c2.caption("Vendas • Metas & Forecast • Rentabilidade • Operação • Fiscal")

# ==========================
# ENTRADAS (GITHUB)
# ==========================
st.sidebar.header("Origem dos dados (GitHub ou upload)")
st.sidebar.caption("Dica: use URLs RAW do GitHub. Para privado, informe o token abaixo.")

# Campos para GitHub (você pode setar defaults aqui)
RAW_MAIN_URL   = st.sidebar.text_input("URL RAW — Base principal (.xlsx, aba 'Carteira de Vendas')", "")
RAW_TAX_URL    = st.sidebar.text_input("URL RAW — Impostos (.xls/.xlsx)", "")
RAW_GOALS_URL  = st.sidebar.text_input("URL RAW — Metas_Brasforma.xlsx (aba 'Metas')", "")
GITHUB_TOKEN   = st.sidebar.text_input("GITHUB_TOKEN (opcional p/ privado)", type="password")

# Fallback para upload manual
file_main  = st.sidebar.file_uploader("OU envie a base principal (.xlsx)", type=["xlsx"])
file_taxes = st.sidebar.file_uploader("OU envie a base de impostos (.xls ou .xlsx)", type=["xls","xlsx"])
file_goals = st.sidebar.file_uploader("OU envie Metas_Brasforma.xlsx", type=["xlsx"])

# ==========================
# HELPERS
# ==========================
def _to_num(x):
    if pd.isna(x): return 0.0
    s = str(x).replace("R$","").replace("%","").replace(".","").replace(" ","").replace("\u00a0","").replace(",",".")
    try: return float(s)
    except: return pd.to_numeric(x, errors="coerce")

def _detect_qty(df: pd.DataFrame):
    for c in ["Qtde","QTD","Quantidade","QTD. TOTAL","Qtd","QTD_TOTAL"]:
        if c in df.columns: return c
    return None

@st.cache_data(show_spinner=False)
def fetch_bytes_from_github(raw_url: str, token: str|None) -> bytes:
    if not raw_url: return b""
    headers = {}
    if token:
        headers["Authorization"] = f"token {token}"
    r = requests.get(raw_url, headers=headers, timeout=30)
    r.raise_for_status()
    return r.content

def get_io_source(raw_url: str, uploaded):
    """Prioridade: GitHub se URL válida → upload se existir → None"""
    if raw_url:
        try:
            b = fetch_bytes_from_github(raw_url, GITHUB_TOKEN or os.getenv("GITHUB_TOKEN"))
            if b: return io.BytesIO(b)
        except Exception as e:
            st.sidebar.warning(f"GitHub falhou: {e}")
    if uploaded: return uploaded
    return None

# ==========================
# LOADERS
# ==========================
@st.cache_data(show_spinner=False)
def load_data(file_like, sheet_name="Carteira de Vendas"):
    df = pd.read_excel(file_like, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]
    ren = {"Transacao":"Transação","Observacao":"Observação","Numero do Pedido":"Pedido"}
    for k,v in ren.items():
        if k in df.columns and v not in df.columns: df.rename(columns={k:v}, inplace=True)

    for c in ["Valor Pedido R$","TICKET MÉDIO","Custo"]:
        if c in df.columns: df[c] = df[c].apply(_to_num).fillna(0.0)

    for c in ["Data / Mês","Data do Pedido","Data da Entrega","Data Final","Data Inserção"]:
        if c in df.columns: df[c] = pd.to_datetime(df[c], errors="coerce")

    if "Data / Mês" in df.columns:
        df["Ano"] = df["Data / Mês"].dt.year
        df["Mes"] = df["Data / Mês"].dt.month
        df["Ano-Mes"] = df["Data / Mês"].dt.to_period("M").astype(str)

    if {"Data do Pedido","Data da Entrega"}.issubset(df.columns):
        df["LeadTime (dias)"] = (df["Data da Entrega"] - df["Data do Pedido"]).dt.days

    if "Atrasado / No prazo" in df.columns:
        df["AtrasadoFlag"] = df["Atrasado / No prazo"].astype(str).str.contains("atras", case=False, na=False)

    qty_col = _detect_qty(df)
    if qty_col and "Custo" in df.columns:
        df["Custo Total"] = df["Custo"]*pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    else:
        df["Custo Total"] = df.get("Custo", 0)

    if {"Valor Pedido R$","Custo Total"}.issubset(df.columns):
        df["Lucro Bruto"] = df["Valor Pedido R$"] - df["Custo Total"]
        df["Margem %"] = np.where(df["Valor Pedido R$"] != 0, df["Lucro Bruto"]/df["Valor Pedido R$"]*100, 0)

    if "Observação" in df.columns and "Família" not in df.columns:
        df["Família"] = df["Observação"]

    if {"Pedido","ITEM"}.issubset(df.columns):
        df["PedidoItemKey"] = df["Pedido"].astype(str) + "||" + df["ITEM"].astype(str)

    return df

@st.cache_data(show_spinner=False)
def load_taxes(file_like):
    """Retorna (valores por Pedido+ITEM, alíquota efetiva por ITEM)."""
    df_tx = pd.read_excel(file_like)
    df_tx.columns = [c.strip() for c in df_tx.columns]
    ren = {
        "Item":"ITEM", "imposto":"Imposto", "Alíquota":"Aliquota",
        "Valor":"Valor_Total", "Valor do Imposto":"Valor_Imposto",
        "Tipo de Operação":"Tipo_Operacao", "Base de Cálculo":"Base_Calculo",
        "Pedido":"Pedido"
    }
    for k,v in ren.items():
        if k in df_tx.columns and v not in df_tx.columns:
            df_tx.rename(columns={k:v}, inplace=True)

    if "Aliquota" in df_tx.columns:
        s = (df_tx["Aliquota"].astype(str).str.replace("%","",regex=False)
             .str.replace(".","",regex=False).str.replace(",",".",regex=False))
        df_tx["Aliquota"] = pd.to_numeric(s, errors="coerce").fillna(0.0)
        if df_tx["Aliquota"].gt(1).mean() > 0.5:
            df_tx["Aliquota"] = df_tx["Aliquota"]/100.0

    for col in ["Valor_Imposto","Base_Calculo"]:
        if col in df_tx.columns:
            df_tx[col] = pd.to_numeric(df_tx[col], errors="coerce").fillna(0.0)

    # A) valores em R$ por Pedido+ITEM
    piv_val = df_tx.pivot_table(
        index=["Pedido","ITEM"], columns="Imposto", values="Valor_Imposto",
        aggfunc="sum", fill_value=0
    )
    piv_val.columns = [f"VALOR_IMPOSTO_{c.upper()}" for c in piv_val.columns]
    piv_val = piv_val.reset_index()
    piv_val["VALOR_IMPOSTO_TOTAL"] = piv_val.drop(columns=["Pedido","ITEM"]).sum(axis=1)

    # B) alíquota efetiva (sum imposto / sum base) por ITEM
    eff = (df_tx.groupby(["ITEM","Imposto"], dropna=False)
           .agg(VALOR_IMPOSTO=("Valor_Imposto","sum"), BASE_CALC=("Base_Calculo","sum"))
           .reset_index())
    eff["ALIQUOTA_EFETIVA"] = np.where(eff["BASE_CALC"]>0, eff["VALOR_IMPOSTO"]/eff["BASE_CALC"], 0.0)
    eff_w = eff.pivot_table(index="ITEM", columns="Imposto", values="ALIQUOTA_EFETIVA", aggfunc="mean", fill_value=0)
    eff_w.columns = [f"ALIQUOTA_EFETIVA_{c.upper()}" for c in eff_w.columns]
    eff_w = eff_w.reset_index()
    # padroniza col de OUTROS (se não existir)
    if "ALIQUOTA_EFETIVA_OUTROS" not in eff_w.columns:
        eff_w["ALIQUOTA_EFETIVA_OUTROS"] = 0.0
    return piv_val, eff_w

@st.cache_data(show_spinner=False)
def load_goals(file_like):
    dfm = pd.read_excel(file_like, sheet_name="Metas", engine="openpyxl")
    dfm.columns = [c.strip() for c in dfm.columns]
    base_cols = {"Ano","Mes","Representante","Meta_Faturamento"}
    miss = base_cols - set(dfm.columns)
    if miss: raise ValueError(f"Metas: colunas ausentes {miss}")
    dfm["Meta_Faturamento"] = dfm["Meta_Faturamento"].apply(_to_num)
    return dfm

# ==========================
# CARREGAMENTO (GitHub → upload)
# ==========================
src_main  = get_io_source(RAW_MAIN_URL,  file_main)
src_taxes = get_io_source(RAW_TAX_URL,   file_taxes)
src_goals = get_io_source(RAW_GOALS_URL, file_goals)

if src_main is None:
    st.warning("Forneça **URL RAW** da base principal ou faça **upload** do .xlsx.")
    st.stop()

df = load_data(src_main)
st.sidebar.success(f"Base principal carregada ({len(df):,} linhas).")

df_tx_pedido_item, df_tx_eff_item = (pd.DataFrame(), pd.DataFrame())
if src_taxes is not None:
    try:
        df_tx_pedido_item, df_tx_eff_item = load_taxes(src_taxes)
        if not df_tx_pedido_item.empty and {"Pedido","ITEM"}.issubset(df.columns):
            df = df.merge(df_tx_pedido_item, on=["Pedido","ITEM"], how="left")
            st.sidebar.success("Base de impostos integrada (R$ por pedido/item).")
        if not df_tx_eff_item.empty and "ITEM" in df.columns:
            df = df.merge(df_tx_eff_item, on="ITEM", how="left")
            st.sidebar.success("Alíquotas efetivas por SKU aplicáveis no simulador.")
    except Exception as e:
        st.sidebar.warning(f"Falha ao integrar impostos: {e}")

df_goals = pd.DataFrame()
if src_goals is not None:
    try:
        df_goals = load_goals(src_goals)
        st.sidebar.success("Metas carregadas.")
    except Exception as e:
        st.sidebar.warning(f"Metas_Brasforma.xlsx inválido: {e}")

# ==========================
# FILTROS
# ==========================
st.sidebar.header("Filtros")
if "Data / Mês" in df.columns:
    dmin, dmax = df["Data / Mês"].min(), df["Data / Mês"].max()
else:
    dmin = dmax = None

periodo = st.sidebar.date_input("Período (data principal)", [dmin, dmax] if dmin is not None and dmax is not None else [])
regional = st.sidebar.multiselect("Regional", sorted(df.get("Regional", pd.Series(dtype=str)).dropna().unique()))
representante = st.sidebar.multiselect("Representante", sorted(df.get("Representante", pd.Series(dtype=str)).dropna().unique()))
uf = st.sidebar.multiselect("UF", sorted(df.get("UF", pd.Series(dtype=str)).dropna().unique()))
status_pf = st.sidebar.multiselect("Status Prod./Fat.", sorted(df.get("Status de Produção / Faturamento", pd.Series(dtype=str)).dropna().unique()))
transacao = st.sidebar.multiselect("Transação", sorted(df.get("Transação", pd.Series(dtype=str)).dropna().unique()))
cliente_ct = st.sidebar.text_input("Cliente (contém)")
sku_ct = st.sidebar.text_input("SKU/Item (contém)")
only_neg = st.sidebar.checkbox("Mostrar apenas linhas com margem negativa", value=False)
flag_dev = st.sidebar.checkbox("Subtrair devoluções do faturamento", value=False)

def apply_filters(data: pd.DataFrame) -> pd.DataFrame:
    out = data.copy()
    if periodo and len(periodo)==2 and "Data / Mês" in out.columns:
        out = out[(out["Data / Mês"] >= pd.to_datetime(periodo[0])) & (out["Data / Mês"] <= pd.to_datetime(periodo[1]))]
    if regional: out = out[out.get("Regional","").isin(regional)]
    if representante: out = out[out.get("Representante","").isin(representante)]
    if uf: out = out[out.get("UF","").isin(uf)]
    if status_pf: out = out[out.get("Status de Produção / Faturamento","").isin(status_pf)]
    if transacao: out = out[out.get("Transação","").isin(transacao)]
    if cliente_ct: out = out[out.get("Nome Cliente","").astype(str).str.contains(cliente_ct, case=False, na=False)]
    if sku_ct: out = out[out.get("ITEM","").astype(str).str.contains(sku_ct, case=False, na=False)]
    if only_neg and "Margem %" in out.columns: out = out[out["Margem %"] < 0]
    if flag_dev and {"Transação","Valor Pedido R$"}.issubset(out.columns):
        mask = out["Transação"].astype(str).str.contains("DEVOL", case=False, na=False)
        out.loc[mask, "Valor Pedido R$"] *= -1
    return out

flt = apply_filters(df)

# ==========================
# VIS UTILS
# ==========================
def money(x):
    try: return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    except: return x

def kpi_block(col, label, value, help_txt=None):
    if help_txt: col.metric(label, money(value), help=help_txt)
    else: col.metric(label, money(value))
def series_12m(flt_df: pd.DataFrame, value_col: str):
    """
    Série mensal robusta para Altair.
    - Normaliza datetime e numérico
    - Remove NaT/NaN
    - Renomeia colunas para nomes simples (Data, Valor)
    - Retorna um placeholder se não houver dados suficientes
    """
    # pré-validações
    if flt_df is None or not isinstance(flt_df, pd.DataFrame):
        return alt.Chart(pd.DataFrame({"Data": [], "Valor": []})).mark_line()

    if "Data / Mês" not in flt_df.columns or value_col not in flt_df.columns:
        return alt.Chart(pd.DataFrame({"Data": [], "Valor": []})).mark_line()

    # normalizações
    ser = flt_df[["Data / Mês", value_col]].copy()
    ser["Data / Mês"] = pd.to_datetime(ser["Data / Mês"], errors="coerce")
    ser[value_col] = pd.to_numeric(ser[value_col], errors="coerce")

    # drop lixo
    ser = ser.dropna(subset=["Data / Mês"])
    if ser.empty:
        return alt.Chart(pd.DataFrame({"Data": [], "Valor": []})).mark_line()

    # resample mensal (MS = month start)
    ser = (ser.groupby(pd.Grouper(key="Data / Mês", freq="MS"))[value_col]
              .sum()
              .reset_index()
              .rename(columns={"Data / Mês": "Data", value_col: "Valor"}))

    # últimos 12
    ser = ser.sort_values("Data").tail(12)

    if ser["Valor"].isna().all() or ser.empty:
        return alt.Chart(pd.DataFrame({"Data": [], "Valor": []})).mark_line()

    # chart robusto
    chart = (
        alt.Chart(ser)
        .mark_line(point=True)
        .encode(
            x=alt.X("yearmonth(Data):T", title="Mês"),
            y=alt.Y("Valor:Q", title="R$"),
            tooltip=[
                alt.Tooltip("yearmonth(Data):T", title="Mês"),
                alt.Tooltip("Valor:Q", title="Valor", format=",.2f"),
            ],
        )
        .properties(height=220)
    )
    return chart


# ==========================
# ABAS
# ==========================
tabs = st.tabs([
    "Diretoria – Metas & Forecast",
    "Visão Executiva",
    "Clientes – RFM",
    "Rentabilidade",
    "Clientes",
    "Produtos",
    "Representantes",
    "Geografia",
    "Operacional",
    "Pareto/ABC",
    "SEBASTIAN",
    "Simulador de Vendas",
    "Exportar"
])

# -------- 0) Diretoria – Metas & Forecast
with tabs[0]:
    st.subheader("Diretoria – Metas & Forecast (mensal)")
    if len(flt) and {"Ano","Mes"}.issubset(flt.columns):
        ano_ref = int(flt["Ano"].max())
        mes_ref = int(flt[flt["Ano"]==ano_ref]["Mes"].max())
    else:
        ano_ref = mes_ref = None

    if {"Ano","Mes","Valor Pedido R$"}.issubset(flt.columns) and ano_ref and mes_ref:
        realizado_mes = flt[(flt["Ano"]==ano_ref)&(flt["Mes"]==mes_ref)]["Valor Pedido R$"].sum()
    else:
        realizado_mes = flt.get("Valor Pedido R$", pd.Series(dtype=float)).sum()

    meta_mes = np.nan
    if not df_goals.empty and ano_ref and mes_ref:
        meta_mes = df_goals[(df_goals["Ano"]==ano_ref)&(df_goals["Mes"]==mes_ref)]["Meta_Faturamento"].sum()

    forecast_mes = realizado_mes
    c1,c2,c3,c4 = st.columns(4)
    kpi_block(c1, "Meta do mês (total)", meta_mes if pd.notna(meta_mes) else 0)
    kpi_block(c2, "Realizado no mês", realizado_mes)
    kpi_block(c3, "Forecast do mês", forecast_mes)
    ating = (realizado_mes/meta_mes*100) if (meta_mes and meta_mes>0) else 0
    c4.metric("Atingimento projetado", f"{ating:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))

    st.markdown("#### Metas por hierarquia comercial — mês de referência")
    opt = st.radio("Nível de análise", ["Regional","Representante"], horizontal=True, label_visibility="collapsed")
    group_col = "Regional" if opt=="Regional" else "Representante"

    if group_col in df.columns:
        base_mes = flt[(flt["Ano"]==ano_ref)&(flt["Mes"]==mes_ref)] if (ano_ref and mes_ref) else flt
        tb_real = base_mes.groupby(group_col, dropna=False)["Valor Pedido R$"].sum().rename("Realizado").reset_index()

        # Metas: se não houver a coluna na planilha de metas (Ex.: falta "Regional"), mapeamos do DF principal
        if not df_goals.empty and ano_ref and mes_ref:
            metas_mes = df_goals[(df_goals["Ano"]==ano_ref)&(df_goals["Mes"]==mes_ref)].copy()
            if group_col not in metas_mes.columns:
                # cria mapeamento Representante -> Regional pelo modo
                if {"Representante","Regional"}.issubset(df.columns):
                    map_rep_reg = (df.groupby("Representante")["Regional"].agg(lambda s: s.mode().iloc[0] if len(s.mode()) else None).to_dict())
                    metas_mes[group_col] = metas_mes["Representante"].map(map_rep_reg)
                else:
                    metas_mes[group_col] = metas_mes.get("Representante")  # fallback
            tb_meta = metas_mes.groupby(group_col, dropna=False)["Meta_Faturamento"].sum().reset_index()
            tb = tb_real.merge(tb_meta, on=group_col, how="left")
        else:
            tb = tb_real.copy()
            tb["Meta_Faturamento"] = np.nan

        tb["Atingimento Atual (%)"] = np.where(tb["Meta_Faturamento"]>0, tb["Realizado"]/tb["Meta_Faturamento"]*100, 0)
        st.dataframe(tb.sort_values("Realizado", ascending=False), use_container_width=True)
        st.altair_chart(alt.Chart(tb).mark_bar().encode(
            x=alt.X("Realizado:Q", title="R$"), y=alt.Y(f"{group_col}:N", sort="-x"),
            tooltip=["Realizado","Meta_Faturamento","Atingimento Atual (%)"]
        ).properties(height=320), use_container_width=True)
    else:
        st.info(f"Coluna '{group_col}' não encontrada na base.")

# -------- 1) Visão Executiva
with tabs[1]:
    st.subheader("Visão Executiva")
    fat_total = flt.get("Valor Pedido R$", pd.Series(dtype=float)).sum()
    pedidos   = flt.get("Pedido", pd.Series(dtype=object)).nunique()
    clientes  = flt.get("Nome Cliente", pd.Series(dtype=object)).nunique()
    skus      = flt.get("ITEM", pd.Series(dtype=object)).nunique()
    lucro     = flt.get("Lucro Bruto", pd.Series(dtype=float)).sum()
    margem    = (lucro/fat_total*100) if fat_total else 0
    pos_lin   = (flt.get("Lucro Bruto", pd.Series(dtype=float))>0).mean()*100 if len(flt) else 0
    a,b,c,d = st.columns(4)
    kpi_block(a,"Faturamento", fat_total)
    b.metric("Pedidos", f"{pedidos:,}".replace(",", "."))
    c.metric("Clientes", f"{clientes:,}".replace(",", "."))
    d.metric("SKUs", f"{skus:,}".replace(",", "."))
    a,b,c,d = st.columns(4)
    kpi_block(a,"Lucro Bruto", lucro)
    b.metric("Margem bruta (%)", f"{margem:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))
    c.metric("% Linhas rentáveis", f"{pos_lin:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))
    if "VALOR_IMPOSTO_TOTAL" in flt.columns:
        kpi_block(d,"Impostos (R$) — base fiscal", flt["VALOR_IMPOSTO_TOTAL"].sum(),
                  help_txt="Somatório por Pedido+ITEM integrados")

    st.markdown("#### Séries — últimos 12 meses")
    g1,g2 = st.columns(2)
    g1.altair_chart(series_12m(flt, "Valor Pedido R$"), use_container_width=True)
    if "Lucro Bruto" in flt.columns:
        g2.altair_chart(series_12m(flt, "Lucro Bruto"), use_container_width=True)

# -------- 2) Clientes – RFM
with tabs[2]:
    st.subheader("Clientes — RFM")
    base = flt.copy()
    ref_date = base["Data do Pedido"].max() if "Data do Pedido" in base.columns else base.get("Data / Mês").max()
    if ref_date is None:
        st.info("Sem datas para calcular RFM.")
    else:
        agg = base.groupby("Nome Cliente").agg(
            Ultima_Compra=("Data do Pedido","max") if "Data do Pedido" in base.columns else ("Data / Mês","max"),
            Frequencia=("Pedido","nunique"), Valor=("Valor Pedido R$","sum")
        ).reset_index()
        agg["RecenciaDias"] = (pd.to_datetime(ref_date)-pd.to_datetime(agg["Ultima_Compra"])).dt.days
        def qscore(s):
            try: return pd.qcut(s, q=3, labels=[3,2,1]).astype(int)
            except: return pd.Series(np.where(s>=s.median(),1,3), index=s.index)
        agg["R_Score"]=qscore(-agg["RecenciaDias"]); agg["F_Score"]=qscore(agg["Frequencia"]); agg["M_Score"]=qscore(agg["Valor"])
        agg["Score"]=agg["R_Score"]+agg["F_Score"]+agg["M_Score"]
        st.dataframe(agg.sort_values(["Score","Valor"], ascending=[False,False]), use_container_width=True)
        st.altair_chart(alt.Chart(agg).mark_circle(size=80).encode(
            x="Frequencia:Q", y="Valor:Q", color="Score:Q",
            tooltip=["Nome Cliente","Frequencia","Valor","RecenciaDias","Score"]
        ).properties(height=420), use_container_width=True)

# -------- 3) Rentabilidade
with tabs[3]:
    st.subheader("Rentabilidade")
    A,B,C,D = st.columns(4)
    kpi_block(A,"Lucro Bruto (total)", flt.get("Lucro Bruto", pd.Series(dtype=float)).sum())
    kpi_block(B,"Margem Bruta (%)", (flt["Lucro Bruto"].sum()/flt["Valor Pedido R$"].sum()*100) if flt["Valor Pedido R$"].sum() else 0)
    C.metric("% Linhas negativas", f"{(flt['Lucro Bruto']<0).mean()*100:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))
    if "VALOR_IMPOSTO_TOTAL" in flt.columns:
        kpi_block(D,"Impostos consolidados (R$)", flt["VALOR_IMPOSTO_TOTAL"].sum())

    st.markdown("#### Top clientes por lucro bruto")
    tb = (flt.groupby("Nome Cliente", dropna=False)
          .agg(Faturamento=("Valor Pedido R$","sum"), Lucro=("Lucro Bruto","sum"))
          .reset_index().sort_values("Lucro", ascending=False).head(30))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb).mark_bar().encode(x="Lucro:Q", y=alt.Y("Nome Cliente:N", sort="-x")), use_container_width=True)

# -------- 4) Clientes
with tabs[4]:
    st.subheader("Clientes — ranking")
    tb = (flt.groupby("Nome Cliente", dropna=False)
          .agg(Faturamento=("Valor Pedido R$","sum"), Pedidos=("Pedido","nunique"), Lucro=("Lucro Bruto","sum"))
          .reset_index().sort_values("Faturamento", ascending=False))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb.head(40)).mark_bar().encode(x="Faturamento:Q", y=alt.Y("Nome Cliente:N", sort="-x")), use_container_width=True)

# -------- 5) Produtos
with tabs[5]:
    st.subheader("Produtos — ranking")
    base = flt.copy()
    if "Família" not in base.columns and "Observação" in base.columns:
        base["Família"] = base["Observação"]
    tb = (base.groupby(["ITEM","Família"], dropna=False)
          .agg(Faturamento=("Valor Pedido R$","sum"), Lucro=("Lucro Bruto","sum"))
          .reset_index().sort_values("Faturamento", ascending=False))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb.head(40)).mark_bar().encode(
        x="Faturamento:Q", y=alt.Y("ITEM:N", sort="-x"), color="Família:N"), use_container_width=True)

# -------- 6) Representantes
with tabs[6]:
    st.subheader("Representantes — ranking")
    tb = (flt.groupby("Representante", dropna=False)
          .agg(Faturamento=("Valor Pedido R$","sum"), Clientes=("Nome Cliente","nunique"), Pedidos=("Pedido","nunique"))
          .reset_index().sort_values("Faturamento", ascending=False))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb.head(40)).mark_bar().encode(x="Faturamento:Q", y=alt.Y("Representante:N", sort="-x")), use_container_width=True)

# -------- 7) Geografia
with tabs[7]:
    st.subheader("Geografia — faturamento por UF")
    tb = (flt.groupby("UF", dropna=False)["Valor Pedido R$"].sum().reset_index().rename(columns={"Valor Pedido R$":"Faturamento"}))
    st.dataframe(tb.sort_values("Faturamento", ascending=False), use_container_width=True)
    st.altair_chart(alt.Chart(tb).mark_bar().encode(x="Faturamento:Q", y=alt.Y("UF:N", sort="-x")), use_container_width=True)

# -------- 8) Operacional
with tabs[8]:
    st.subheader("Operacional — prazos e status")
    if "LeadTime (dias)" in flt.columns:
        dsc = flt["LeadTime (dias)"].describe()[["count","mean","50%","min","max"]].rename({"50%":"median"})
        st.write(dsc)
        st.altair_chart(alt.Chart(flt.dropna(subset=["LeadTime (dias)"])).mark_bar().encode(
            x=alt.X("LeadTime (dias):Q", bin=alt.Bin(maxbins=30)), y="count()"
        ).properties(height=250), use_container_width=True)
    if "Atrasado / No prazo" in flt.columns:
        tb = flt.groupby("Atrasado / No prazo").size().reset_index(name="Qtd")
        st.dataframe(tb, use_container_width=True)
        st.altair_chart(alt.Chart(tb).mark_bar().encode(x="Qtd:Q", y=alt.Y("Atrasado / No prazo:N", sort="-x")), use_container_width=True)

# -------- 9) Pareto/ABC
with tabs[9]:
    st.subheader("Pareto/ABC")
    tb = (flt.groupby("Nome Cliente", dropna=False)["Valor Pedido R$"].sum()
          .reset_index().sort_values("Valor Pedido R$", ascending=False))
    tb["%Acum"] = tb["Valor Pedido R$"].cumsum()/tb["Valor Pedido R$"].sum()*100
    tb["Classe"] = np.where(tb["%Acum"]<=80,"A", np.where(tb["%Acum"]<=95,"B","C"))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb).mark_line(point=True).encode(x=alt.X("row_number()"), y="%Acum:Q"), use_container_width=True)

# -------- 10) SEBASTIAN
with tabs[10]:
    st.subheader("SEBASTIAN — visão tática (12 meses + período)")
    base = flt.copy()
    if "Data / Mês" in base.columns:
        g_mes = (base.groupby(pd.Grouper(key="Data / Mês", freq="MS"))
                 .agg(Faturamento=("Valor Pedido R$","sum"), Pedidos=("Pedido","nunique"), Clientes=("Nome Cliente","nunique"))
                 .reset_index().sort_values("Data / Mês"))
        st.markdown("##### Histórico de pedidos (12 meses)")
        st.altair_chart(alt.Chart(g_mes.tail(12)).mark_line(point=True).encode(x="yearmonth(Data / Mês):T", y="Pedidos:Q"),
                        use_container_width=True)
        st.markdown("##### Histórico de faturamento (12 meses)")
        st.altair_chart(alt.Chart(g_mes.tail(12)).mark_line(point=True).encode(x="yearmonth(Data / Mês):T", y="Faturamento:Q"),
                        use_container_width=True)

    if "Representante" in base.columns:
        rep = (base.groupby("Representante", dropna=False)
               .agg(Faturamento=("Valor Pedido R$","sum"), Clientes=("Nome Cliente","nunique"), Pedidos=("Pedido","nunique"))
               .reset_index().sort_values("Faturamento", ascending=False))
        st.markdown("##### Desempenho individual (período do filtro)")
        st.dataframe(rep, use_container_width=True)
        st.altair_chart(alt.Chart(rep.head(40)).mark_bar().encode(x="Faturamento:Q", y=alt.Y("Representante:N", sort="-x")), use_container_width=True)

# -------- 11) Simulador de Vendas
with tabs[11]:
    st.subheader("Simulador de Vendas — multi-SKU com MC e impostos (alíquota efetiva por SKU)")
    if not {"ITEM","Valor Pedido R$"}.issubset(df.columns):
        st.info("Base sem colunas ITEM/Valor Pedido R$.")
    else:
        all_skus = sorted(df["ITEM"].dropna().unique().tolist())
        sel_skus = st.multiselect("Selecione SKUs para simular", all_skus[:50])
        if sel_skus:
            qty_col = _detect_qty(df)
            hist = df[df["ITEM"].isin(sel_skus)].copy()
            hist["_Q"] = pd.to_numeric(hist[qty_col], errors="coerce").fillna(0) if qty_col else 1
            hist["_Fat"] = pd.to_numeric(hist["Valor Pedido R$"], errors="coerce").fillna(0)
            hist["_CustoTot"] = pd.to_numeric(hist.get("Custo Total", 0), errors="coerce").fillna(0)
            g = (hist.groupby("ITEM").agg(Qtd=("_Q","sum"), Fat=("_Fat","sum"), CustoTot=("_CustoTot","sum")).reset_index())
            g["PrecoMed"] = np.where(g["Qtd"]>0, g["Fat"]/g["Qtd"], 0)
            g["CustoMed"] = np.where(g["Qtd"]>0, g["CustoTot"]/g["Qtd"], 0)

            ali_cols = ["ALIQUOTA_EFETIVA_ICMS","ALIQUOTA_EFETIVA_PIS","ALIQUOTA_EFETIVA_COFINS","ALIQUOTA_EFETIVA_OUTROS"]
            for c in ali_cols:
                if c not in df.columns: df[c] = np.nan
            ali_eff = df.groupby("ITEM")[ali_cols].mean().reset_index()
            g = g.merge(ali_eff, on="ITEM", how="left")

            st.markdown("##### Histórico consolidado por SKU")
            st.dataframe(g.assign(PrecoMed=g["PrecoMed"].map(money), CustoMed=g["CustoMed"].map(money),
                                  Fat=g["Fat"].map(money), CustoTot=g["CustoTot"].map(money)),
                         use_container_width=True)

            st.markdown("##### Parâmetros globais (override se faltar alíquota por SKU)")
            icms_pct = st.number_input("ICMS (%)", value=18.0, step=0.5)
            pis_pct  = st.number_input("PIS (%)", value=1.65, step=0.05)
            cof_pct  = st.number_input("COFINS (%)", value=7.6, step=0.1)
            out_pct  = st.number_input("Outros impostos (%)", value=0.0, step=0.1)
            frete_pct= st.number_input("Frete (% faturamento)", value=0.0, step=0.5)
            com_pct  = st.number_input("Comissão (% faturamento)", value=0.0, step=0.5)
            mc_alvo  = st.number_input("Margem de Contribuição mínima (%)", value=20.0, step=0.5)

            def fback(v, default_pct): return float(v) if pd.notna(v) and v>=0 else (default_pct/100.0)

            rows = []
            for _, r in g.iterrows():
                sku = r["ITEM"]
                qtd = st.number_input(f"Qtd — {sku}", value=int(max(r["Qtd"], 100)), min_value=0, step=10)
                adj_p = st.number_input(f"Ajuste preço (%) — {sku}", value=0.0, step=1.0)
                adj_c = st.number_input(f"Ajuste custo (%) — {sku}", value=0.0, step=1.0)
                pu_manual = st.number_input(f"Preço manual (0=hist.) — {sku}", value=0.0, step=1.0)

                pu = pu_manual if pu_manual>0 else r["PrecoMed"]*(1+adj_p/100)
                cu = r["CustoMed"]*(1+adj_c/100)

                ali_icms = fback(r.get("ALIQUOTA_EFETIVA_ICMS", np.nan), icms_pct)
                ali_pis  = fback(r.get("ALIQUOTA_EFETIVA_PIS", np.nan),  pis_pct)
                ali_cof  = fback(r.get("ALIQUOTA_EFETIVA_COFINS", np.nan), cof_pct)
                ali_out  = fback(r.get("ALIQUOTA_EFETIVA_OUTROS", np.nan), out_pct)
                td = (frete_pct + com_pct)/100.0

                fat = pu*qtd
                custo_tot = cu*qtd
                t_imp = ali_icms + ali_pis + ali_cof + ali_out
                receita_liq = fat*(1 - t_imp)
                desp_var = fat*td
                mc_val = receita_liq - custo_tot - desp_var
                mc_pct = (mc_val/receita_liq*100) if receita_liq>0 else 0

                M = mc_alvo/100.0
                A = (1 - M)*(1 - t_imp) - td
                preco_min = np.nan if A<=0 else cu/A
                desc_max = np.nan if (pd.isna(preco_min) or pu==0) else (1 - preco_min/pu)*100

                rows.append([sku, qtd, pu, cu, fat, custo_tot, mc_val, mc_pct, preco_min, desc_max,
                             ali_icms, ali_pis, ali_cof, ali_out])

            sim = pd.DataFrame(rows, columns=[
                "ITEM","Qtd","Preço Unit.","Custo Unit.","Faturamento","Custo Total",
                "MC (R$)","MC (%)","Preço mín. (MC alvo)","Desconto máx. (%)",
                "Aliq_ICMS","Aliq_PIS","Aliq_COFINS","Aliq_Outros"
            ])
            st.markdown("##### Resultado da simulação")
            st.dataframe(sim.style.format({
                "Preço Unit.":"{:,.2f}".format,"Custo Unit.":"{:,.2f}".format,
                "Faturamento":"{:,.2f}".format,"Custo Total":"{:,.2f}".format,
                "MC (R$)":"{:,.2f}".format,"MC (%)":"{:,.1f}%".format,
                "Preço mín. (MC alvo)":"{:,.2f}".format,"Desconto máx. (%)":"{:,.1f}%".format,
                "Aliq_ICMS":"{:,.2%}".format,"Aliq_PIS":"{:,.2%}".format,
                "Aliq_COFINS":"{:,.2%}".format,"Aliq_Outros":"{:,.2%}".format
            }), use_container_width=True)

            # Mini DRE
            fat_tot = sim["Faturamento"].sum()
            custo_tot = sim["Custo Total"].sum()
            if fat_tot>0:
                t_imp_eff = ((sim["Aliq_ICMS"]+sim["Aliq_PIS"]+sim["Aliq_COFINS"]+sim["Aliq_Outros"])*sim["Faturamento"]).sum()/fat_tot
            else:
                t_imp_eff = (icms_pct+pis_pct+cof_pct+out_pct)/100.0
            td = (frete_pct+com_pct)/100.0
            receita_liq = fat_tot*(1 - t_imp_eff)
            desp_var = fat_tot*td
            mc_val = receita_liq - custo_tot - desp_var
            mc_p = (mc_val/receita_liq*100) if receita_liq>0 else 0
            c1,c2,c3,c4 = st.columns(4)
            kpi_block(c1,"Faturamento bruto", fat_tot)
            kpi_block(c2,"Receita líquida", receita_liq)
            kpi_block(c3,"MC (R$)", mc_val)
            c4.metric("MC (%)", f"{mc_p:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))

            # Exports
            st.download_button("Baixar simulação (CSV)",
                               data=sim.to_csv(index=False).encode("utf-8"),
                               file_name="simulacao_venda_brasforma.csv", mime="text/csv")

            def build_pdf(df_sim: pd.DataFrame):
                pdf = FPDF(); pdf.set_auto_page_break(auto=True, margin=10); pdf.add_page()
                pdf.set_font("helvetica","B",14); pdf.cell(0,10,"Simulação de Vendas — Brasforma", ln=1)
                pdf.set_font("helvetica","",10)
                pdf.cell(0,6,f"Faturamento: {money(fat_tot)}  |  Receita líquida: {money(receita_liq)}  |  MC: {money(mc_val)} ({mc_p:.1f}%)", ln=1)
                pdf.ln(2)
                headers = ["ITEM","Qtd","Preço","Custo","Fat.","Custo Tot.","MC R$","MC %","Preço mín.","Desc máx %"]
                widths  = [35,15,18,18,22,22,22,18,22,18]
                pdf.set_font("helvetica","B",9)
                for h,w in zip(headers, widths): pdf.cell(w,6,h,1,0,"C")
                pdf.ln(6); pdf.set_font("helvetica","",9)
                for _,row in df_sim.iterrows():
                    vals=[str(row["ITEM"]), int(row["Qtd"]), f"{row['Preço Unit.']:.2f}", f"{row['Custo Unit.']:.2f}",
                          f"{row['Faturamento']:.2f}", f"{row['Custo Total']:.2f}", f"{row['MC (R$)']:.2f}",
                          f"{row['MC (%)']:.1f}", ("" if pd.isna(row['Preço mín. (MC alvo)']) else f"{row['Preço mín. (MC alvo)']:.2f}"),
                          ("" if pd.isna(row['Desconto máx. (%)']) else f"{row['Desconto máx. (%)']:.1f}")]
                    for v,w in zip(vals,widths): pdf.cell(w,6,str(v),1)
                    pdf.ln(6)
                return pdf.output(dest="S").encode("latin1","ignore")

            try:
                st.download_button("Baixar simulação (PDF)", data=build_pdf(sim),
                                   file_name="simulacao_venda_brasforma.pdf", mime="application/pdf")
            except Exception as e:
                st.warning(f"PDF: {e}")

# -------- 12) Exportar
with tabs[12]:
    st.subheader("Exportar — CSV filtrado")
    st.dataframe(flt.head(200), use_container_width=True)
    st.download_button("Baixar CSV filtrado",
                       data=flt.to_csv(index=False).encode("utf-8"),
                       file_name="brasforma_filtro.csv", mime="text/csv")

st.success("✅ v25 ativo: ingestão GitHub + metas por hierarquia corrigidas + fiscal por SKU preservado.")
