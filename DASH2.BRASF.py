# DASHBOARD COMERCIAL BRASFORMA — v24
# Upgrades:
# - Integração fiscal com cálculo de ALÍQUOTA EFETIVA por SKU (ponderada pela Base de Cálculo)
# - Override automático no Simulador: ICMS/PIS/COFINS/Outros por ITEM quando disponível
# - Mantém todas as abas e features do v23 (Metas, SEBASTIAN, etc.)

import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path
from PIL import Image
from fpdf import FPDF

# ==============================================================
# CONFIGURAÇÃO GERAL
# ==============================================================

LOGO_PATHS = ["images.png", "logo_brasforma.png"]
PAGE_TITLE = "Brasforma — Dashboard Comercial v24"

st.set_page_config(page_title=PAGE_TITLE, layout="wide", page_icon=LOGO_PATHS[0])

def _load_logo():
    for p in LOGO_PATHS:
        try:
            return Image.open(p)
        except Exception:
            continue
    return None

APP_LOGO = _load_logo()

with st.container():
    c1, c2 = st.columns([1, 8])
    with c1:
        if APP_LOGO is not None:
            st.image(APP_LOGO, use_container_width=True)
    with c2:
        st.markdown("## **Dashboard Comercial — Brasforma**")
        st.caption("Vendas • Metas & Forecast • Rentabilidade • Operação • Fiscal")

# ==============================================================
# UPLOADS
# ==============================================================

st.sidebar.header("Bases de Dados")
file_main = st.sidebar.file_uploader("Base principal (.xlsx) — aba 'Carteira de Vendas'", type=["xlsx"])
file_taxes = st.sidebar.file_uploader("Base de Impostos (.xls ou .xlsx) — opcional", type=["xls", "xlsx"])
file_goals = st.sidebar.file_uploader("Metas_Brasforma.xlsx — opcional (aba 'Metas')", type=["xlsx"])

# ==============================================================
# HELPERS
# ==============================================================

def _to_num(x):
    if pd.isna(x): 
        return 0.0
    s = str(x).replace("R$", "").replace("%", "").replace(".", "").replace(" ", "").replace("\u00a0","").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return pd.to_numeric(x, errors="coerce")

def _detect_qty(df: pd.DataFrame):
    candidates = ["Qtde", "QTD", "Quantidade", "QTD. TOTAL", "Qtd", "QTD_TOTAL"]
    for c in candidates:
        if c in df.columns:
            return c
    return None

# ------------------ LOAD DATA ------------------

@st.cache_data(show_spinner=False)
def load_data(path, sheet_name="Carteira de Vendas"):
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    df.columns = [c.strip() for c in df.columns]

    # normalizações simples
    ren = {"Transacao": "Transação", "Observacao": "Observação", "Numero do Pedido": "Pedido"}
    for k, v in ren.items():
        if k in df.columns and v not in df.columns:
            df.rename(columns={k: v}, inplace=True)

    num_cols = ["Valor Pedido R$", "TICKET MÉDIO", "Custo"]
    for c in num_cols:
        if c in df.columns:
            df[c] = df[c].apply(_to_num).fillna(0.0)

    date_cols = ["Data / Mês", "Data do Pedido", "Data da Entrega", "Data Final", "Data Inserção"]
    for c in date_cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    if "Data / Mês" in df.columns:
        df["Ano"] = df["Data / Mês"].dt.year
        df["Mes"] = df["Data / Mês"].dt.month
        df["Ano-Mes"] = df["Data / Mês"].dt.to_period("M").astype(str)

    # LeadTime
    if {"Data do Pedido","Data da Entrega"}.issubset(df.columns):
        df["LeadTime (dias)"] = (df["Data da Entrega"] - df["Data do Pedido"]).dt.days

    # Atrasado flag
    if "Atrasado / No prazo" in df.columns:
        df["AtrasadoFlag"] = df["Atrasado / No prazo"].astype(str).str.contains("atras", case=False, na=False)

    qty_col = _detect_qty(df)
    if qty_col and "Custo" in df.columns:
        df["Custo Total"] = df["Custo"] * pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    else:
        df["Custo Total"] = df.get("Custo", 0)

    if {"Valor Pedido R$", "Custo Total"}.issubset(df.columns):
        df["Lucro Bruto"] = df["Valor Pedido R$"] - df["Custo Total"]
        df["Margem %"] = np.where(df["Valor Pedido R$"] != 0, df["Lucro Bruto"] / df["Valor Pedido R$"] * 100, 0)

    # Família pela coluna I (Observação), quando existir
    if "Observação" in df.columns and "Família" not in df.columns:
        df["Família"] = df["Observação"]

    # chave potencial
    if "Pedido" in df.columns and "ITEM" in df.columns:
        df["PedidoItemKey"] = df["Pedido"].astype(str) + "||" + df["ITEM"].astype(str)

    return df

@st.cache_data(show_spinner=False)
def load_taxes(path):
    """
    Lê a planilha fiscal e constrói:
      (A) Resumo por Pedido+ITEM com impostos em R$
      (B) ALÍQUOTA EFETIVA por SKU (ITEM) ponderada pela Base de Cálculo
    Colunas esperadas (pelo print enviado): 
      'Pedido', 'ITEM', 'imposto', 'Alíquota', 'Base de Cálculo', 'Valor do Imposto', 'Tipo de Operação'
    """
    try:
        df_tx = pd.read_excel(path)
        df_tx.columns = [c.strip() for c in df_tx.columns]
        ren = {
            "Item":"ITEM", "imposto":"Imposto", "Alíquota":"Aliquota",
            "Valor":"Valor_Total", "Valor do Imposto":"Valor_Imposto",
            "Tipo de Operação":"Tipo_Operacao", "Base de Cálculo":"Base_Calculo"
        }
        for k,v in ren.items():
            if k in df_tx.columns and v not in df_tx.columns:
                df_tx.rename(columns={k:v}, inplace=True)

        # numéricos
        if "Aliquota" in df_tx.columns:
            s = (df_tx["Aliquota"].astype(str).str.replace("%","",regex=False)
                 .str.replace(".", "", regex=False).str.replace(",",".", regex=False))
            df_tx["Aliquota"] = pd.to_numeric(s, errors="coerce").fillna(0.0)
            if df_tx["Aliquota"].gt(1).mean() > 0.5:
                df_tx["Aliquota"] = df_tx["Aliquota"]/100.0

        for col in ["Valor_Imposto", "Base_Calculo"]:
            if col in df_tx.columns:
                df_tx[col] = pd.to_numeric(df_tx[col], errors="coerce").fillna(0.0)

        # A) Consolidação por Pedido+ITEM → valores em R$
        piv_val = df_tx.pivot_table(
            index=["Pedido","ITEM"],
            columns="Imposto",
            values="Valor_Imposto",
            aggfunc="sum", fill_value=0
        )
        piv_val.columns = [f"VALOR_IMPOSTO_{c.upper()}" for c in piv_val.columns]
        piv_val = piv_val.reset_index()
        piv_val["VALOR_IMPOSTO_TOTAL"] = piv_val.drop(columns=["Pedido","ITEM"]).sum(axis=1)

        # B) Alíquota efetiva por SKU: sum(Valor_Imposto) / sum(Base_Calculo) para cada Imposto
        # Primeiro, somar por ITEM + Imposto
        eff = (df_tx.groupby(["ITEM","Imposto"], dropna=False)
                     .agg(VALOR_IMPOSTO=("Valor_Imposto","sum"),
                          BASE_CALC=("Base_Calculo","sum"))
                     .reset_index())
        eff["ALIQUOTA_EFETIVA"] = np.where(eff["BASE_CALC"]>0, eff["VALOR_IMPOSTO"]/eff["BASE_CALC"], 0.0)
        # Pivotar por imposto para wide
        eff_w = eff.pivot_table(index="ITEM", columns="Imposto", values="ALIQUOTA_EFETIVA", aggfunc="mean", fill_value=0)
        eff_w.columns = [f"ALIQUOTA_EFETIVA_{c.upper()}" for c in eff_w.columns]
        eff_w = eff_w.reset_index()

        # Resultado final: duas tabelas
        return piv_val, eff_w
    except Exception as e:
        st.warning(f"Falha ao ler impostos: {e}")
        return pd.DataFrame(), pd.DataFrame()

@st.cache_data(show_spinner=False)
def load_goals(path):
    try:
        dfm = pd.read_excel(path, sheet_name="Metas", engine="openpyxl")
        dfm.columns = [c.strip() for c in dfm.columns]
        base_cols = {"Ano","Mes","Representante","Meta_Faturamento"}
        miss = base_cols - set(dfm.columns)
        if miss:
            raise ValueError(f"Colunas ausentes na aba 'Metas': {miss}")
        dfm["Meta_Faturamento"] = dfm["Meta_Faturamento"].apply(_to_num)
        return dfm
    except Exception as e:
        st.warning(f"Metas_Brasforma.xlsx não lida: {e}")
        return pd.DataFrame()

# ==============================================================
# CARREGAR BASES
# ==============================================================

if file_main is None:
    st.warning("Envie a **base principal** para iniciar.")
    st.stop()

df = load_data(file_main)
st.sidebar.success(f"Base principal carregada ({len(df):,} linhas).")

df_tx_pedido_item = pd.DataFrame()
df_tx_eff_item = pd.DataFrame()

if file_taxes is not None:
    df_tx_pedido_item, df_tx_eff_item = load_taxes(file_taxes)
    # Merge (A) por Pedido+ITEM para expor valores em R$ de impostos na base principal
    if not df_tx_pedido_item.empty and {"Pedido","ITEM"}.issubset(df.columns):
        df = df.merge(df_tx_pedido_item, on=["Pedido","ITEM"], how="left")
        st.sidebar.success("Base de impostos integrada (valores em R$ por pedido/item).")
    # Merge (B) por ITEM para expor alíquotas efetivas por SKU (override no simulador)
    if not df_tx_eff_item.empty and {"ITEM"}.issubset(df.columns):
        df = df.merge(df_tx_eff_item, on="ITEM", how="left")
        st.sidebar.success("Alíquotas efetivas por SKU disponíveis (override no simulador).")

df_goals = pd.DataFrame()
if file_goals is not None:
    df_goals = load_goals(file_goals)
    if not df_goals.empty:
        st.sidebar.success("Metas carregadas.")

# ==============================================================
# FILTROS GLOBAIS (inclui Transação e Devoluções)
# ==============================================================

st.sidebar.header("Filtros")

if "Data / Mês" in df.columns:
    dmin, dmax = df["Data / Mês"].min(), df["Data / Mês"].max()
else:
    dmin, dmax = None, None

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
    if flag_dev and "Transação" in out.columns and "Valor Pedido R$" in out.columns:
        mask = out["Transação"].astype(str).str.contains("DEVOL", case=False, na=False)
        out.loc[mask, "Valor Pedido R$"] *= -1
    return out

flt = apply_filters(df)

# ==============================================================
# FUNÇÕES DE VISUALIZAÇÃO
# ==============================================================

def money(x): 
    try:
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    except Exception:
        return x

def kpi_block(col, label, value, help_txt=None):
    if help_txt:
        col.metric(label, money(value), help=help_txt)
    else:
        col.metric(label, money(value))

def series_12m(flt_df, value_col):
    if "Data / Mês" not in flt_df.columns: 
        return alt.Chart(pd.DataFrame())
    ser = (flt_df.dropna(subset=["Data / Mês"])
                  .groupby(pd.Grouper(key="Data / Mês", freq="MS"))[value_col]
                  .sum().reset_index())
    ser = ser.sort_values("Data / Mês").tail(12)
    return alt.Chart(ser).mark_line(point=True).encode(
        x=alt.X("yearmonth(Data / Mês):T", title="Mês"),
        y=alt.Y(f"{value_col}:Q", title="R$"),
        tooltip=[alt.Tooltip("yearmonth(Data / Mês):T","Mês"), alt.Tooltip(f"{value_col}:Q","Valor", format=",.2f")]
    ).properties(height=220)

# ==============================================================
# ABAS
# ==============================================================

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

# ------------------ 0) Diretoria – Metas & Forecast ------------------

with tabs[0]:
    st.subheader("Diretoria – Metas & Forecast (mensal)")

    if len(flt) and "Ano" in flt.columns and "Mes" in flt.columns:
        ano_ref = int(flt["Ano"].max())
        mes_ref = int(flt[flt["Ano"]==ano_ref]["Mes"].max())
    else:
        ano_ref, mes_ref = None, None

    if {"Ano","Mes","Valor Pedido R$"}.issubset(flt.columns) and ano_ref and mes_ref:
        realizado_mes = flt[(flt["Ano"]==ano_ref) & (flt["Mes"]==mes_ref)]["Valor Pedido R$"].sum()
    else:
        realizado_mes = flt.get("Valor Pedido R$", pd.Series(dtype=float)).sum()

    meta_mes = np.nan
    if not df_goals.empty and ano_ref and mes_ref:
        meta_mes = df_goals[(df_goals["Ano"]==ano_ref) & (df_goals["Mes"]==mes_ref)]["Meta_Faturamento"].sum()

    forecast_mes = realizado_mes

    c1, c2, c3, c4 = st.columns(4)
    kpi_block(c1, "Meta do mês (total)", meta_mes if pd.notna(meta_mes) else 0)
    kpi_block(c2, "Realizado no mês", realizado_mes)
    kpi_block(c3, "Forecast do mês", forecast_mes)
    ating = (realizado_mes/meta_mes*100) if (meta_mes and meta_mes>0) else 0
    c4.metric("Atingimento projetado", f"{ating:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))

    st.markdown("#### Metas por hierarquia comercial — mês de referência")
    opt = st.radio("Nível de análise", ["Regional","Representante"], horizontal=True, label_visibility="collapsed")
    group_col = "Regional" if opt == "Regional" else "Representante"
    if group_col in flt.columns:
        tb = flt[(flt["Ano"]==ano_ref) & (flt["Mes"]==mes_ref)] if (ano_ref and mes_ref) else flt
        tb = tb.groupby(group_col, dropna=False)["Valor Pedido R$"].sum().rename("Realizado").reset_index()

        if not df_goals.empty and ano_ref and mes_ref and "Representante" in df_goals.columns:
            metas_g = (df_goals[(df_goals["Ano"]==ano_ref) & (df_goals["Mes"]==mes_ref)]
                       .groupby(group_col, dropna=False)["Meta_Faturamento"].sum().reset_index())
            tb = tb.merge(metas_g, on=group_col, how="left")
        else:
            tb["Meta_Faturamento"] = np.nan

        tb["Atingimento Atual (%)"] = np.where(tb["Meta_Faturamento"]>0, tb["Realizado"]/tb["Meta_Faturamento"]*100, 0)
        st.dataframe(tb.sort_values("Realizado", ascending=False), use_container_width=True)

        ch = alt.Chart(tb).mark_bar().encode(
            x=alt.X("Realizado:Q", title="R$"),
            y=alt.Y(f"{group_col}:N", sort="-x", title=group_col),
            tooltip=["Realizado","Meta_Faturamento","Atingimento Atual (%)"]
        ).properties(height=320)
        st.altair_chart(ch, use_container_width=True)
    else:
        st.info(f"Coluna '{group_col}' não encontrada na base.")

# ------------------ 1) Visão Executiva ------------------

with tabs[1]:
    st.subheader("Visão Executiva")

    fat_total = flt.get("Valor Pedido R$", pd.Series(dtype=float)).sum()
    pedidos = flt.get("Pedido", pd.Series(dtype=object)).nunique()
    clientes = flt.get("Nome Cliente", pd.Series(dtype=object)).nunique()
    skus = flt.get("ITEM", pd.Series(dtype=object)).nunique()
    lucro = flt.get("Lucro Bruto", pd.Series(dtype=float)).sum()
    margem = (lucro/fat_total*100) if fat_total else 0
    linhas_pos = (flt.get("Lucro Bruto", pd.Series(dtype=float))>0).mean()*100 if len(flt) else 0

    a,b,c,d = st.columns(4)
    kpi_block(a, "Faturamento", fat_total)
    b.metric("Pedidos", f"{pedidos:,}".replace(",", "."))
    c.metric("Clientes", f"{clientes:,}".replace(",", "."))
    d.metric("SKUs", f"{skus:,}".replace(",", "."))

    a,b,c,d = st.columns(4)
    kpi_block(a, "Lucro Bruto", lucro)
    b.metric("Margem bruta (%)", f"{margem:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))
    c.metric("% Linhas rentáveis", f"{linhas_pos:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))
    if "VALOR_IMPOSTO_TOTAL" in flt.columns:
        kpi_block(d, "Impostos (R$) — base fiscal", flt["VALOR_IMPOSTO_TOTAL"].sum(), help_txt="Somatório por linha consolidada")

    st.markdown("#### Séries — últimos 12 meses")
    g1, g2 = st.columns(2)
    with g1: st.altair_chart(series_12m(flt, "Valor Pedido R$"), use_container_width=True)
    with g2:
        if "Lucro Bruto" in flt.columns:
            st.altair_chart(series_12m(flt, "Lucro Bruto"), use_container_width=True)

# ------------------ 2) Clientes – RFM ------------------

with tabs[2]:
    st.subheader("Clientes — RFM")
    base = flt.copy()
    ref_date = base["Data do Pedido"].max() if "Data do Pedido" in base.columns else base.get("Data / Mês").max()
    if ref_date is None:
        st.info("Sem datas para calcular RFM.")
    else:
        agg = base.groupby("Nome Cliente").agg(
            Ultima_Compra=("Data do Pedido","max") if "Data do Pedido" in base.columns else ("Data / Mês","max"),
            Frequencia=("Pedido","nunique"),
            Valor=("Valor Pedido R$","sum")
        ).reset_index()
        agg["RecenciaDias"] = (pd.to_datetime(ref_date) - pd.to_datetime(agg["Ultima_Compra"])).dt.days
        def qscore(s):
            try:
                return pd.qcut(s, q=3, labels=[3,2,1]).astype(int)
            except Exception:
                return pd.Series(np.where(s >= s.median(),1,3), index=s.index)
        agg["R_Score"] = qscore(-agg["RecenciaDias"])
        agg["F_Score"] = qscore(agg["Frequencia"])
        agg["M_Score"] = qscore(agg["Valor"])
        agg["Score"] = agg["R_Score"] + agg["F_Score"] + agg["M_Score"]
        st.dataframe(agg.sort_values(["Score","Valor"], ascending=[False,False]), use_container_width=True)
        st.markdown("#### Dispersão — Frequência × Valor")
        chart = alt.Chart(agg).mark_circle(size=80).encode(
            x="Frequencia:Q", y="Valor:Q", color="Score:Q",
            tooltip=["Nome Cliente","Frequencia","Valor","RecenciaDias","Score"]
        ).properties(height=420)
        st.altair_chart(chart, use_container_width=True)

# ------------------ 3) Rentabilidade ------------------

with tabs[3]:
    st.subheader("Rentabilidade")
    colA, colB, colC, colD = st.columns(4)
    kpi_block(colA, "Lucro Bruto (total)", flt.get("Lucro Bruto", pd.Series(dtype=float)).sum())
    kpi_block(colB, "Margem Bruta (%)", (flt["Lucro Bruto"].sum()/flt["Valor Pedido R$"].sum()*100) if flt["Valor Pedido R$"].sum() else 0)
    colC.metric("% Linhas negativas", f"{(flt['Lucro Bruto']<0).mean()*100:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))
    if "VALOR_IMPOSTO_TOTAL" in flt.columns:
        kpi_block(colD, "Impostos consolidados (R$)", flt["VALOR_IMPOSTO_TOTAL"].sum())

    st.markdown("#### Top clientes por lucro bruto")
    tb = (flt.groupby("Nome Cliente", dropna=False)
          .agg(Faturamento=("Valor Pedido R$","sum"),
               Lucro=("Lucro Bruto","sum"))
          .reset_index()
          .sort_values("Lucro", ascending=False)
          .head(30))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb).mark_bar().encode(x="Lucro:Q", y=alt.Y("Nome Cliente:N", sort="-x")), use_container_width=True)

# ------------------ 4) Clientes ------------------

with tabs[4]:
    st.subheader("Clientes — ranking")
    tb = (flt.groupby("Nome Cliente", dropna=False)
          .agg(Faturamento=("Valor Pedido R$","sum"),
               Pedidos=("Pedido","nunique"),
               Lucro=("Lucro Bruto","sum"))
          .reset_index()
          .sort_values("Faturamento", ascending=False))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb.head(40)).mark_bar().encode(x="Faturamento:Q", y=alt.Y("Nome Cliente:N", sort="-x")), use_container_width=True)

# ------------------ 5) Produtos ------------------

with tabs[5]:
    st.subheader("Produtos — ranking")
    base = flt.copy()
    if "Família" not in base.columns and "Observação" in base.columns:
        base["Família"] = base["Observação"]
    tb = (base.groupby(["ITEM","Família"], dropna=False)
          .agg(Faturamento=("Valor Pedido R$","sum"),
               Lucro=("Lucro Bruto","sum"))
          .reset_index()
          .sort_values("Faturamento", ascending=False))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb.head(40)).mark_bar().encode(x="Faturamento:Q", y=alt.Y("ITEM:N", sort="-x"), color="Família:N"), use_container_width=True)

# ------------------ 6) Representantes ------------------

with tabs[6]:
    st.subheader("Representantes — ranking")
    tb = (flt.groupby("Representante", dropna=False)
          .agg(Faturamento=("Valor Pedido R$","sum"),
               Clientes=("Nome Cliente","nunique"),
               Pedidos=("Pedido","nunique"))
          .reset_index()
          .sort_values("Faturamento", ascending=False))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb.head(40)).mark_bar().encode(x="Faturamento:Q", y=alt.Y("Representante:N", sort="-x")), use_container_width=True)

# ------------------ 7) Geografia ------------------

with tabs[7]:
    st.subheader("Geografia — faturamento por UF")
    tb = (flt.groupby("UF", dropna=False)["Valor Pedido R$"].sum().reset_index().rename(columns={"Valor Pedido R$":"Faturamento"}))
    st.dataframe(tb.sort_values("Faturamento", ascending=False), use_container_width=True)
    st.altair_chart(alt.Chart(tb).mark_bar().encode(x="Faturamento:Q", y=alt.Y("UF:N", sort="-x")), use_container_width=True)

# ------------------ 8) Operacional ------------------

with tabs[8]:
    st.subheader("Operacional — prazos e status")
    if "LeadTime (dias)" in flt.columns:
        d = flt["LeadTime (dias)"].describe()[["count","mean","50%","min","max"]].rename({"50%":"median"})
        st.write(d)
        st.altair_chart(alt.Chart(flt.dropna(subset=["LeadTime (dias)"])).mark_bar().encode(
            x=alt.X("LeadTime (dias):Q", bin=alt.Bin(maxbins=30)), y="count()"
        ).properties(height=250), use_container_width=True)
    if "Atrasado / No prazo" in flt.columns:
        tb = flt.groupby("Atrasado / No prazo").size().reset_index(name="Qtd")
        st.dataframe(tb, use_container_width=True)
        st.altair_chart(alt.Chart(tb).mark_bar().encode(x="Qtd:Q", y=alt.Y("Atrasado / No prazo:N", sort="-x")), use_container_width=True)

# ------------------ 9) Pareto/ABC ------------------

with tabs[9]:
    st.subheader("Pareto/ABC")
    tb = (flt.groupby("Nome Cliente", dropna=False)["Valor Pedido R$"].sum()
           .reset_index().sort_values("Valor Pedido R$", ascending=False))
    tb["%Acum"] = tb["Valor Pedido R$"].cumsum()/tb["Valor Pedido R$"].sum()*100
    tb["Classe"] = np.where(tb["%Acum"]<=80, "A", np.where(tb["%Acum"]<=95,"B","C"))
    st.dataframe(tb, use_container_width=True)
    st.altair_chart(alt.Chart(tb).mark_line(point=True).encode(x=alt.X("row_number()"), y="%Acum:Q"), use_container_width=True)

# ------------------ 10) SEBASTIAN ------------------

with tabs[10]:
    st.subheader("SEBASTIAN — visão tática (12 meses + período)")

    base = flt.copy()
    if "Data / Mês" in base.columns:
        g_mes = (base.groupby(pd.Grouper(key="Data / Mês", freq="MS"))
                      .agg(Faturamento=("Valor Pedido R$","sum"),
                           Pedidos=("Pedido","nunique"),
                           Clientes=("Nome Cliente","nunique"))
                      .reset_index()
                      .sort_values("Data / Mês"))
        st.markdown("##### Histórico de pedidos (12 meses)")
        st.altair_chart(alt.Chart(g_mes.tail(12)).mark_line(point=True).encode(
            x="yearmonth(Data / Mês):T", y="Pedidos:Q"
        ), use_container_width=True)
        st.markdown("##### Histórico de faturamento (12 meses)")
        st.altair_chart(alt.Chart(g_mes.tail(12)).mark_line(point=True).encode(
            x="yearmonth(Data / Mês):T", y="Faturamento:Q"
        ), use_container_width=True)

    if "Representante" in base.columns:
        rep = (base.groupby("Representante", dropna=False)
                    .agg(Faturamento=("Valor Pedido R$","sum"),
                         Clientes=("Nome Cliente","nunique"),
                         Pedidos=("Pedido","nunique"))
                    .reset_index()
                    .sort_values("Faturamento", ascending=False))
        st.markdown("##### Desempenho individual (período do filtro)")
        st.dataframe(rep, use_container_width=True)
        st.altair_chart(alt.Chart(rep.head(40)).mark_bar().encode(
            x="Faturamento:Q", y=alt.Y("Representante:N", sort="-x")
        ), use_container_width=True)

# ------------------ 11) Simulador de Vendas ------------------

with tabs[11]:
    st.subheader("Simulador de Vendas — multi-SKU com MC e impostos (alíquota efetiva por SKU)")

    if "ITEM" not in df.columns or "Valor Pedido R$" not in df.columns:
        st.info("Base sem colunas ITEM/Valor Pedido R$.")
    else:
        all_skus = sorted(df["ITEM"].dropna().unique().tolist())
        sel_skus = st.multiselect("Selecione SKUs para simular", all_skus[:50])

        if sel_skus:
            qty_col = _detect_qty(df)
            hist = df[df["ITEM"].isin(sel_skus)].copy()
            if qty_col is None:
                hist["_Q"] = 1
            else:
                hist["_Q"] = pd.to_numeric(hist[qty_col], errors="coerce").fillna(0)
            hist["_Fat"] = pd.to_numeric(hist["Valor Pedido R$"], errors="coerce").fillna(0)
            hist["_CustoTot"] = pd.to_numeric(hist.get("Custo Total", 0), errors="coerce").fillna(0)

            g = (hist.groupby("ITEM")
                     .agg(Qtd=("_Q","sum"), Fat=("_Fat","sum"), CustoTot=("_CustoTot","sum"))
                     .reset_index())
            g["PrecoMed"] = np.where(g["Qtd"]>0, g["Fat"]/g["Qtd"], 0)
            g["CustoMed"] = np.where(g["Qtd"]>0, g["CustoTot"]/g["Qtd"], 0)

            # Captura alíquotas efetivas por SKU (se vieram da base fiscal)
            ali_cols = ["ALIQUOTA_EFETIVA_ICMS", "ALIQUOTA_EFETIVA_PIS", "ALIQUOTA_EFETIVA_COFINS", "ALIQUOTA_EFETIVA_IPI", "ALIQUOTA_EFETIVA_OUTROS"]
            for c in ali_cols:
                if c not in df.columns:
                    df[c] = np.nan
            ali_eff = df.groupby("ITEM")[ali_cols].mean().reset_index()

            g = g.merge(ali_eff, on="ITEM", how="left")

            st.markdown("##### Histórico consolidado por SKU")
            st.dataframe(g.assign(PrecoMed=g["PrecoMed"].map(money),
                                  CustoMed=g["CustoMed"].map(money),
                                  Fat=g["Fat"].map(money),
                                  CustoTot=g["CustoTot"].map(money)), use_container_width=True)

            st.markdown("##### Parâmetros globais (override se faltar alíquota por SKU)")
            icms_pct = st.number_input("ICMS (%)", value=18.0, step=0.5)
            pis_pct = st.number_input("PIS (%)", value=1.65, step=0.05)
            cofins_pct = st.number_input("COFINS (%)", value=7.6, step=0.1)
            outros_pct = st.number_input("Outros impostos (%)", value=0.0, step=0.1)
            frete_pct = st.number_input("Frete (% faturamento)", value=0.0, step=0.5)
            com_pct = st.number_input("Comissão (% faturamento)", value=0.0, step=0.5)
            mc_alvo = st.number_input("Margem de Contribuição mínima (%)", value=20.0, step=0.5)

            rows = []
            for _, r in g.iterrows():
                sku = r["ITEM"]
                qtd = st.number_input(f"Quantidade simulada — {sku}", value=int(max(r["Qtd"], 100)), min_value=0, step=10)
                adj_preco = st.number_input(f"Ajuste preço vs histórico (%) — {sku}", value=0.0, step=1.0)
                adj_custo = st.number_input(f"Ajuste custo vs histórico (%) — {sku}", value=0.0, step=1.0)
                preco_manual = st.number_input(f"Preço unitário manual (0=usar histórico) — {sku}", value=0.0, step=1.0)

                pu_hist = r["PrecoMed"] * (1 + adj_preco/100)
                cu_hist = r["CustoMed"] * (1 + adj_custo/100)
                pu = preco_manual if preco_manual>0 else pu_hist
                cu = cu_hist

                # Alíquotas efetivas por SKU (fallback para globais)
                ali_icms = r.get("ALIQUOTA_EFETIVA_ICMS", np.nan)
                ali_pis  = r.get("ALIQUOTA_EFETIVA_PIS", np.nan)
                ali_cof  = r.get("ALIQUOTA_EFETIVA_COFINS", np.nan)
                ali_outros = r.get("ALIQUOTA_EFETIVA_OUTROS", np.nan)

                def fallback(val, default_pct):
                    return float(val) if pd.notna(val) and val>=0 else (default_pct/100.0)

                ali_icms  = fallback(ali_icms, icms_pct)
                ali_pis   = fallback(ali_pis,  pis_pct)
                ali_cof   = fallback(ali_cof,  cofins_pct)
                ali_outros= fallback(ali_outros, outros_pct)

                td = (frete_pct + com_pct)/100.0

                fat = pu * qtd
                custo_tot = cu * qtd
                t_imp = ali_icms + ali_pis + ali_cof + ali_outros

                receita_liq = fat * (1 - t_imp)
                desp_var = fat * td
                mc_val = receita_liq - custo_tot - desp_var
                mc_pct = (mc_val/receita_liq*100) if receita_liq>0 else 0

                M = mc_alvo/100.0
                A = (1 - M) * (1 - t_imp) - td
                preco_min = np.nan if A<=0 else cu / A
                desc_max = np.nan if (pd.isna(preco_min) or pu==0) else (1 - preco_min/pu) * 100

                rows.append([sku, qtd, pu, cu, fat, custo_tot, mc_val, mc_pct, preco_min, desc_max,
                             ali_icms, ali_pis, ali_cof, ali_outros])

            sim = pd.DataFrame(rows, columns=[
                "ITEM","Qtd","Preço Unit.","Custo Unit.","Faturamento","Custo Total",
                "MC (R$)","MC (%)","Preço mín. (MC alvo)","Desconto máx. (%)",
                "Aliq_ICMS","Aliq_PIS","Aliq_COFINS","Aliq_Outros"
            ])
            st.markdown("##### Resultado da simulação")
            st.dataframe(sim.style.format({
                "Preço Unit.":"{:,.2f}".format, "Custo Unit.":"{:,.2f}".format,
                "Faturamento":"{:,.2f}".format, "Custo Total":"{:,.2f}".format,
                "MC (R$)":"{:,.2f}".format, "MC (%)":"{:,.1f}%".format,
                "Preço mín. (MC alvo)":"{:,.2f}".format, "Desconto máx. (%)":"{:,.1f}%".format,
                "Aliq_ICMS":"{:,.2%}".format, "Aliq_PIS":"{:,.2%}".format,
                "Aliq_COFINS":"{:,.2%}".format, "Aliq_Outros":"{:,.2%}".format
            }), use_container_width=True)

            st.markdown("##### Mini DRE")
            fat_tot = sim["Faturamento"].sum()
            custo_tot = sim["Custo Total"].sum()
            # média ponderada das alíquotas efetivas simuladas por faturamento
            if fat_tot > 0:
                t_imp_eff = ((sim["Aliq_ICMS"]+sim["Aliq_PIS"]+sim["Aliq_COFINS"]+sim["Aliq_Outros"]) * sim["Faturamento"]).sum() / fat_tot
            else:
                t_imp_eff = (icms_pct+pis_pct+cofins_pct+outros_pct)/100.0
            td = (frete_pct+com_pct)/100.0
            receita_liq = fat_tot * (1 - t_imp_eff)
            desp_var = fat_tot * td
            mc_val = receita_liq - custo_tot - desp_var
            mc_p = (mc_val/receita_liq*100) if receita_liq>0 else 0
            c1,c2,c3,c4 = st.columns(4)
            kpi_block(c1, "Faturamento bruto", fat_tot)
            kpi_block(c2, "Receita líquida", receita_liq)
            kpi_block(c3, "MC (R$)", mc_val)
            c4.metric("MC (%)", f"{mc_p:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))

            # Export CSV
            csv_bytes = sim.to_csv(index=False).encode("utf-8")
            st.download_button("Baixar simulação (CSV)", data=csv_bytes, file_name="simulacao_venda_brasforma.csv", mime="text/csv")

            # Export PDF simples
            def build_pdf(df_sim: pd.DataFrame):
                pdf = FPDF()
                pdf.set_auto_page_break(auto=True, margin=10)
                pdf.add_page()
                pdf.set_font("helvetica","B",14); pdf.cell(0,10,"Simulação de Vendas — Brasforma", ln=1)
                pdf.set_font("helvetica","",10)
                pdf.cell(0,6,f"Faturamento: {money(fat_tot)}  |  Receita líquida: {money(receita_liq)}  |  MC: {money(mc_val)} ({mc_p:.1f}%)", ln=1)
                pdf.ln(2)
                headers = ["ITEM","Qtd","Preço","Custo","Fat.","Custo Tot.","MC R$","MC %","Preço mín.","Desc máx %"]
                widths = [35,15,18,18,22,22,22,18,22,18]
                pdf.set_font("helvetica","B",9)
                for h,w in zip(headers, widths):
                    pdf.cell(w,6,h,1,0,"C")
                pdf.ln(6)
                pdf.set_font("helvetica","",9)
                for _,row in df_sim.iterrows():
                    vals = [
                        str(row["ITEM"]), int(row["Qtd"]), f"{row['Preço Unit.']:.2f}", f"{row['Custo Unit.']:.2f}",
                        f"{row['Faturamento']:.2f}", f"{row['Custo Total']:.2f}",
                        f"{row['MC (R$)']:.2f}", f"{row['MC (%)']:.1f}",
                        ("" if pd.isna(row['Preço mín. (MC alvo)']) else f"{row['Preço mín. (MC alvo)']:.2f}"),
                        ("" if pd.isna(row['Desconto máx. (%)']) else f"{row['Desconto máx. (%)']:.1f}")
                    ]
                    for v,w in zip(vals,widths):
                        pdf.cell(w,6,str(v),1)
                    pdf.ln(6)
                return pdf.output(dest="S").encode("latin1","ignore")

            try:
                pdf_bytes = build_pdf(sim)
                st.download_button("Baixar simulação (PDF)", data=pdf_bytes, file_name="simulacao_venda_brasforma.pdf", mime="application/pdf")
            except Exception as e:
                st.warning(f"PDF: {e}")

# ------------------ 12) Exportar ------------------

with tabs[12]:
    st.subheader("Exportar — CSV filtrado")
    st.dataframe(flt.head(200), use_container_width=True)
    st.download_button(
        "Baixar CSV filtrado",
        data=flt.to_csv(index=False).encode("utf-8"),
        file_name="brasforma_filtro.csv",
        mime="text/csv"
    )

st.success("✅ v24 carregado: alíquotas efetivas por SKU integradas e aplicadas no simulador, sem perder nenhuma função.")
