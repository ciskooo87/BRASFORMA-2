# streamlit_app_brasforma_v19.py
# Brasforma – Dashboard Comercial v19
# - Mantém TODAS as abas anteriores
# - Aba SEBASTIAN reforçada:
#   * Indicadores avançados (ticket, ativação, churn, concentração Top 5)
#   * Pipeline por status (R$ e quantidade)
#   * Família = coluna I (Observação)

import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path
from fpdf import FPDF  # para geração de PDF

st.set_page_config(page_title="Brasforma – Dashboard Comercial v19", layout="wide")

# ---------------- Utils ----------------
def to_num(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan

def fmt_money(v):
    if pd.isna(v):
        return "-"
    return ("R$ " + f"{v:,.2f}").replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_int(v):
    if pd.isna(v):
        return "-"
    return f"{int(v):,}".replace(",", ".")

def fmt_pct_safe(v, decimals=1):
    if pd.isna(v):
        return "-"
    return f"{v:.{decimals}f}%".replace(".", ",")

def display_table(df, money_cols=None, pct_cols=None, int_cols=None, max_rows=500):
    money_cols = money_cols or []
    pct_cols = pct_cols or []
    int_cols = int_cols or []
    view = df.copy().head(max_rows)
    for c in view.columns:
        if c in money_cols:
            view[c] = view[c].apply(fmt_money)
        elif c in pct_cols:
            view[c] = view[c].apply(lambda x: fmt_pct_safe(x, 1))
        elif c in int_cols:
            view[c] = view[c].apply(fmt_int)
    st.dataframe(view, use_container_width=True)

def build_simulation_pdf(sim_df, faturamento_sim, margem_contrib, margem_contrib_pct,
                         icms_pct, pis_pct, cofins_pct, outros_pct,
                         frete_pct, comissao_pct, margem_target_pct):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, "Simulação de Vendas - Brasforma", ln=True)
    pdf.ln(2)

    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, f"Faturamento simulado (total): {fmt_money(faturamento_sim)}", ln=True)
    pdf.cell(0, 6, f"Margem de Contribuição (R$): {fmt_money(margem_contrib)}", ln=True)
    mc_pct_txt = fmt_pct_safe(margem_contrib_pct, 1) if not pd.isna(margem_contrib_pct) else "-"
    mc_alvo_txt = fmt_pct_safe(margem_target_pct, 1)
    pdf.cell(0, 6, f"Margem de Contribuição (%): {mc_pct_txt} | MC alvo: {mc_alvo_txt}", ln=True)
    pdf.ln(3)

    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, "Parâmetros globais da simulação:", ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(0, 5, f"ICMS: {fmt_pct_safe(icms_pct,1)} | PIS: {fmt_pct_safe(pis_pct,2)} | COFINS: {fmt_pct_safe(cofins_pct,2)} | Outros: {fmt_pct_safe(outros_pct,1)}", ln=True)
    pdf.cell(0, 5, f"Frete: {fmt_pct_safe(frete_pct,1)} | Comissão: {fmt_pct_safe(comissao_pct,1)}", ln=True)
    pdf.ln(3)

    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, "Resumo por SKU:", ln=True)
    pdf.ln(1)

    headers = ["SKU", "Qtd", "Preço", "Custo", "Fat.", "Lucro", "Marg%"]
    widths = [45, 12, 23, 23, 23, 23, 15]

    pdf.set_font("Helvetica", "B", 8)
    for h, w in zip(headers, widths):
        pdf.cell(w, 6, h, border=1, align="C")
    pdf.ln(6)

    pdf.set_font("Helvetica", "", 7)
    for _, row in sim_df.iterrows():
        sku = str(row.get("SKU", ""))[:22]
        qtd = row.get("Qtd Simulada", 0)
        preco = row.get("Preço Unitário Simulado", 0.0)
        custo = row.get("Custo Unitário Simulado", 0.0)
        fat = row.get("Faturamento Simulado", 0.0)
        lucro = row.get("Lucro Bruto Simulado", 0.0)
        marg = row.get("Margem Bruta %", np.nan)

        pdf.cell(widths[0], 5, sku, border=1)
        pdf.cell(widths[1], 5, f"{int(qtd)}", border=1, align="R")
        pdf.cell(widths[2], 5, f"{preco:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), border=1, align="R")
        pdf.cell(widths[3], 5, f"{custo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), border=1, align="R")
        pdf.cell(widths[4], 5, f"{fat:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."), border=1, align="R")
        pdf.cell(widths[5], 5, f"{lucro:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."), border=1, align="R")
        if pd.isna(marg):
            marg_txt = "-"
        else:
            marg_txt = f"{marg:.1f}%".replace(".", ",")
        pdf.cell(widths[6], 5, marg_txt, border=1, align="R")
        pdf.ln(5)

    return pdf.output(dest="S").encode("latin-1")

# ---------------- Load & prep ----------------
@st.cache_data(show_spinner=False)
def load_data(path: str, sheet_name="Carteira de Vendas"):
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        st.error("Falha ao abrir Excel. Verifique .xlsx e dependência openpyxl.")
        st.exception(e)
        st.stop()
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = [c.strip() for c in df.columns]

    for col in ["Data / Mês","Data Final","Data do Pedido","Data da Entrega","Data Inserção"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in ["Valor Pedido R$","TICKET MÉDIO","Quant. Pedidos","Custo"]:
        if col in df.columns:
            if col == "Quant. Pedidos":
                df[col] = pd.to_numeric(df[col], errors="coerce")
            else:
                df[col] = df[col].apply(to_num)

    if "Data / Mês" in df.columns:
        df["Ano"] = df["Data / Mês"].dt.year
        df["Mes"] = df["Data / Mês"].dt.month
        df["Ano-Mes"] = df["Data / Mês"].dt.to_period("M").astype(str)

    if "Data do Pedido" in df.columns and "Data da Entrega" in df.columns:
        df["LeadTime (dias)"] = (df["Data da Entrega"] - df["Data do Pedido"]).dt.days

    if "Atrasado / No prazo" in df.columns:
        df["AtrasadoFlag"] = df["Atrasado / No prazo"].astype(str).str.contains("Atras", case=False, na=False)

    qty_candidates = ["Qtde","QTDE","Quantidade","Quantidade Pedido","Qtd","QTD","Quant.","Quant","Qde","QTD.","QTD PEDIDA","QTD PEDIDO","QTD SOLICITADA","QTD Solicitada"]
    qty_col = None
    for c in qty_candidates:
        if c in df.columns:
            qty_col = c
            break
    if qty_col is None:
        try:
            qty_col = df.columns[12]
        except Exception:
            qty_col = None

    if "Custo" in df.columns:
        if qty_col is not None:
            df[qty_col] = df[qty_col].apply(to_num)
            df["Custo Total"] = df["Custo"].apply(to_num) * df[qty_col]
        else:
            df["Custo Total"] = df["Custo"].apply(to_num)
    else:
        df["Custo Total"] = np.nan

    if "Valor Pedido R$" in df.columns:
        df["Lucro Bruto"] = df["Valor Pedido R$"] - df["Custo Total"]
        df["Margem %"] = np.where(df["Valor Pedido R$"]>0, 100*df["Lucro Bruto"]/df["Valor Pedido R$"], np.nan)

    if "Pedido" in df.columns and "ITEM" in df.columns:
        df["PedidoItemKey"] = df["Pedido"].astype(str) + "||" + df["ITEM"].astype(str)

    return df, qty_col

# ---- Estimativa de elasticidade histórica por SKU ----
@st.cache_data(show_spinner=False)
def estimate_elasticities(df_source: pd.DataFrame, qty_col: str):
    if qty_col is None:
        return pd.DataFrame(columns=["SKU","Elasticidade","N_Obs","R2"])

    df2 = df_source.copy()
    if "ITEM" not in df2.columns or "Valor Pedido R$" not in df2.columns:
        return pd.DataFrame(columns=["SKU","Elasticidade","N_Obs","R2"])

    if "Ano-Mes" not in df2.columns:
        if "Data / Mês" in df2.columns:
            df2["Ano-Mes"] = pd.to_datetime(df2["Data / Mês"], errors="coerce").dt.to_period("M").astype(str)
        else:
            return pd.DataFrame(columns=["SKU","Elasticidade","N_Obs","R2"])

    df2[qty_col] = df2[qty_col].apply(to_num)

    mask = (
        df2[qty_col].notna() & (df2[qty_col] > 0) &
        df2["Valor Pedido R$"].notna() & (df2["Valor Pedido R$"] > 0) &
        df2["ITEM"].notna() & df2["Ano-Mes"].notna()
    )
    df2 = df2[mask].copy()
    if df2.empty:
        return pd.DataFrame(columns=["SKU","Elasticidade","N_Obs","R2"])

    df2["PrecoUnit"] = df2["Valor Pedido R$"] / df2[qty_col]

    grp = df2.groupby(["ITEM","Ano-Mes"], as_index=False).agg(
        PrecoMed=("PrecoUnit","mean"),
        Qtd=(qty_col,"sum"),
    )

    results = []
    for sku, g in grp.groupby("ITEM"):
        g = g[(g["PrecoMed"] > 0) & (g["Qtd"] > 0)]
        if len(g) < 3 or g["PrecoMed"].nunique() < 2 or g["Qtd"].nunique() < 2:
            continue

        log_p = np.log(g["PrecoMed"].values)
        log_q = np.log(g["Qtd"].values)
        try:
            beta = np.polyfit(log_p, log_q, 1)
            slope = float(beta[0])
            pred = np.polyval(beta, log_p)
            ss_res = float(np.sum((log_q - pred)**2))
            ss_tot = float(np.sum((log_q - log_q.mean())**2))
            r2 = 1 - ss_res/ss_tot if ss_tot > 0 else np.nan
        except Exception:
            continue

        if not np.isfinite(slope):
            continue

        e = slope
        if e > 0:
            e = -0.3
        if e < -5:
            e = -5.0

        results.append({"SKU": sku, "Elasticidade": e, "N_Obs": len(g), "R2": r2})

    if not results:
        return pd.DataFrame(columns=["SKU","Elasticidade","N_Obs","R2"])
    return pd.DataFrame(results).sort_values("Elasticidade")

DEFAULT_DATA = "Dashboard - Comite Semanal - Brasforma IA (1).xlsx"
ALT_DATA = "Dashboard - Comite Semanal - Brasforma (1).xlsx"

st.sidebar.title("Fonte de dados")
uploaded = st.sidebar.file_uploader("Envie a base (.xlsx)", type=["xlsx"], accept_multiple_files=False)
data_path = uploaded if uploaded is not None else (DEFAULT_DATA if Path(DEFAULT_DATA).exists() else ALT_DATA)
st.sidebar.caption(f"Arquivo em uso: **{data_path}**")

df, qty_col = load_data(data_path)

st.sidebar.title("Filtros")
if "Data / Mês" in df.columns:
    min_date = pd.to_datetime(df["Data / Mês"]).min()
    max_date = pd.to_datetime(df["Data / Mês"]).max()
    d_ini, d_fim = st.sidebar.date_input("Período (Data / Mês)", value=(min_date, max_date))
else:
    d_ini = d_fim = None

reg = st.sidebar.multiselect("Regional", sorted(df["Regional"].dropna().unique()) if "Regional" in df.columns else [])
rep = st.sidebar.multiselect("Representante", sorted(df["Representante"].dropna().unique()) if "Representante" in df.columns else [])
uf  = st.sidebar.multiselect("UF", sorted(df["UF"].dropna().unique()) if "UF" in df.columns else [])
stat = st.sidebar.multiselect("Status Prod./Fat.", sorted(df["Status de Produção / Faturamento"].dropna().unique()) if "Status de Produção / Faturamento" in df.columns else [])
cliente = st.sidebar.text_input("Cliente (contém)")
item = st.sidebar.text_input("SKU/Item (contém)")
show_neg = st.sidebar.checkbox("Mostrar apenas linhas com margem negativa", value=False)

def apply_filters(_df):
    flt = _df.copy()
    if "Data / Mês" in flt.columns and d_ini is not None:
        flt = flt[(flt["Data / Mês"] >= pd.to_datetime(d_ini)) & (flt["Data / Mês"] <= pd.to_datetime(d_fim))]
    if reg:
        flt = flt[flt["Regional"].isin(reg)]
    if rep:
        flt = flt[flt["Representante"].isin(rep)]
    if uf:
        flt = flt[flt["UF"].isin(uf)]
    if stat:
        flt = flt[flt["Status de Produção / Faturamento"].isin(stat)]
    if cliente:
        flt = flt[flt["Nome Cliente"].astype(str).str.contains(cliente, case=False, na=False)]
    if item:
        flt = flt[flt["ITEM"].astype(str).str.contains(item, case=False, na=False)]
    if show_neg and "Lucro Bruto" in flt.columns:
        flt = flt[flt["Lucro Bruto"] < 0]
    return flt

flt = apply_filters(df)

def calc_kpis(_df):
    fat = _df["Valor Pedido R$"].sum() if "Valor Pedido R$" in _df.columns else np.nan
    n_ped = _df["Pedido"].nunique() if "Pedido" in _df.columns else len(_df)
    n_cli = _df["Nome Cliente"].nunique() if "Nome Cliente" in _df.columns else np.nan
    n_sku = _df["ITEM"].nunique() if "ITEM" in _df.columns else np.nan
    ticket = (fat / n_ped) if (n_ped and n_ped>0) else np.nan
    lucro = _df["Lucro Bruto"].sum() if "Lucro Bruto" in _df.columns else np.nan
    margem_w = 100*(lucro/fat) if (pd.notna(lucro) and fat and fat>0) else np.nan
    pct_rentavel = 100.0*(_df["Lucro Bruto"]>0).mean() if "Lucro Bruto" in _df.columns and len(_df)>0 else np.nan
    return fat, n_ped, n_cli, n_sku, ticket, lucro, margem_w, pct_rentavel

fat, n_ped, n_cli, n_sku, ticket, lucro, margem_w, pct_rentavel = calc_kpis(flt)

tabs = st.tabs([
    "Visão Executiva","Clientes – RFM","Rentabilidade","Clie
