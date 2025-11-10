# streamlit_app_brasforma_v18.py
# Brasforma – Dashboard Comercial v18
# - Mantém TODAS as abas anteriores (v17)
# - Upgrade: na aba SEBASTIAN, faturamento por família usa explicitamente a COLUNA I (Observação) da planilha

import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path
from fpdf import FPDF  # para geração de PDF

st.set_page_config(page_title="Brasforma – Dashboard Comercial v18", layout="wide")

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
    "Visão Executiva","Clientes – RFM","Rentabilidade","Clientes","Produtos",
    "Representantes","Geografia","Operacional","Pareto/ABC","SEBASTIAN",
    "Simulador de Vendas","Exportar"
])
(tab_exec, tab_rfm, tab_profit, tab_cli, tab_sku,
 tab_rep, tab_geo, tab_ops, tab_pareto, tab_seb, tab_sim, tab_export) = tabs

# ---------------- Visão Executiva ----------------
with tab_exec:
    st.subheader("KPIs Executivos")
    c1, c2, c3 = st.columns(3)
    c1.metric("Faturamento", fmt_money(fat))
    c2.metric("Pedidos", fmt_int(n_ped))
    c3.metric("Ticket Médio", fmt_money(ticket) if pd.notna(ticket) else "-")
    c4, c5, c6 = st.columns(3)
    c4.metric("Lucro Bruto", fmt_money(lucro))
    c5.metric("Margem Bruta (pond.)", fmt_pct_safe(margem_w) if pd.notna(margem_w) else "-")
    c6.metric("% Itens Rentáveis", fmt_pct_safe(pct_rentavel) if pd.notna(pct_rentavel) else "-")

    st.markdown("### KPI gráficos")
    if {"Ano-Mes","Valor Pedido R$","Lucro Bruto"}.issubset(flt.columns):
        serie = flt.groupby("Ano-Mes", as_index=False).agg({
            "Valor Pedido R$":"sum",
            "Lucro Bruto":"sum"
        }).sort_values("Ano-Mes")
        mg = flt.groupby("Ano-Mes", as_index=False).apply(
            lambda d: pd.Series({
                "Margem %": (100*d["Lucro Bruto"].sum()/d["Valor Pedido R$"].sum()) if d["Valor Pedido R$"].sum()>0 else np.nan
            })
        ).reset_index(drop=True)
        serie = serie.merge(mg, on="Ano-Mes", how="left")
        if len(serie) > 12:
            serie = serie.tail(12)

        k1, k2, k3 = st.columns(3)
        with k1:
            st.caption("Faturamento – últimos 12 meses")
            st.altair_chart(
                alt.Chart(serie).mark_area(opacity=0.4).encode(
                    x=alt.X("Ano-Mes:N", sort=None, title=None),
                    y=alt.Y("Valor Pedido R$:Q", title=None),
                    tooltip=[alt.Tooltip("Ano-Mes:N"), alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
                ),
                use_container_width=True
            )
        with k2:
            st.caption("Lucro Bruto – últimos 12 meses")
            st.altair_chart(
                alt.Chart(serie).mark_area(opacity=0.4).encode(
                    x=alt.X("Ano-Mes:N", sort=None, title=None),
                    y=alt.Y("Lucro Bruto:Q", title=None),
                    tooltip=[alt.Tooltip("Ano-Mes:N"), alt.Tooltip("Lucro Bruto:Q", format=",.0f")]
                ),
                use_container_width=True
            )
        with k3:
            st.caption("Margem Bruta (%) – últimos 12 meses")
            st.altair_chart(
                alt.Chart(serie).mark_line(point=True).encode(
                    x=alt.X("Ano-Mes:N", sort=None, title=None),
                    y=alt.Y("Margem %:Q", title=None),
                    tooltip=[alt.Tooltip("Ano-Mes:N"), alt.Tooltip("Margem %:Q", format=",.1f")]
                ),
                use_container_width=True
            )

    if "Lucro Bruto" in flt.columns and len(flt) > 0:
        pos = int((flt["Lucro Bruto"] > 0).sum())
        neg = int((flt["Lucro Bruto"] < 0).sum())
        donut_df = pd.DataFrame({"Categoria": ["Rentáveis","Negativos"], "Qtd": [pos, neg]})
        cdon1, cdon2 = st.columns([2,1])
        with cdon1:
            st.caption("Composição de linhas – rentáveis vs negativas")
            st.altair_chart(
                alt.Chart(donut_df).mark_arc(innerRadius=60).encode(
                    theta="Qtd:Q",
                    color="Categoria:N",
                    tooltip=["Categoria","Qtd"]
                ).properties(height=300),
                use_container_width=True
            )
        with cdon2:
            tot = pos + neg
            st.metric("% Linhas Rentáveis", fmt_pct_safe(100*pos/tot) if tot>0 else "-")

# ---------------- RFM ----------------
def compute_rfm(_df, ref_date=None):
    base = _df.dropna(subset=["Nome Cliente"]) if "Nome Cliente" in _df.columns else _df.copy()
    if ref_date is None:
        if "Data do Pedido" in base.columns and base["Data do Pedido"].notna().any():
            ref_date = pd.to_datetime(base["Data do Pedido"]).max()
        elif "Data / Mês" in base.columns and base["Data / Mês"].notna().any():
            ref_date = pd.to_datetime(base["Data / Mês"]).max()
        else:
            ref_date = pd.Timestamp.today().normalize()

    if "Data do Pedido" in base.columns and base["Data do Pedido"].notna().any():
        last_buy = base.groupby("Nome Cliente")["Data do Pedido"].max().rename("UltimaCompra")
    else:
        last_buy = base.groupby("Nome Cliente")["Data / Mês"].max().rename("UltimaCompra")

    freq = base.groupby("Nome Cliente")["Pedido"].nunique().rename("Frequencia") if "Pedido" in base.columns else base.groupby("Nome Cliente").size().rename("Frequencia")
    val = base.groupby("Nome Cliente")["Valor Pedido R$"].sum().rename("Valor") if "Valor Pedido R$" in base.columns else None

    rfm = pd.concat([last_buy, freq, val], axis=1)
    rfm["RecenciaDias"] = (pd.to_datetime(ref_date) - pd.to_datetime(rfm["UltimaCompra"])).dt.days

    def safe_qcut(s, labels):
        try:
            return pd.qcut(s.rank(method="first"), q=len(labels), labels=labels)
        except Exception:
            return pd.Series([labels[len(labels)//2]]*len(s), index=s.index)

    rfm["R_Score"] = safe_qcut(-rfm["RecenciaDias"].fillna(rfm["RecenciaDias"].max()), labels=[1,2,3])
    rfm["F_Score"] = safe_qcut(rfm["Frequencia"].fillna(0), labels=[1,2,3])
    rfm["M_Score"] = safe_qcut(rfm["Valor"].fillna(0), labels=[1,2,3])
    rfm["Score"] = rfm[["R_Score","F_Score","M_Score"]].astype(int).sum(axis=1)

    def seg(row):
        r,f,m = int(row["R_Score"]), int(row["F_Score"]), int(row["M_Score"])
        if r>=3 and f>=3 and m>=3: return "Campeões"
        if f>=3 and r>=2: return "Leais"
        if r==1 and m>=2: return "Em risco"
        if r==1 and f==1: return "Perdidos"
        return "Oportunidades"

    rfm["Segmento"] = rfm.apply(seg, axis=1)
    rfm = rfm.sort_values(["Score","Valor","Frequencia"], ascending=[False,False,False]).reset_index()
    rfm.rename(columns={"index":"Nome Cliente"}, inplace=True)
    return rfm

with tab_rfm:
    st.subheader("Clientes – RFM (Recência, Frequência, Valor)")
    ref_date = pd.to_datetime(d_fim) if d_fim is not None else None
    rfm = compute_rfm(flt, ref_date=ref_date)
    segs = sorted(rfm["Segmento"].unique())
    pick = st.multiselect("Segmentos", segs, default=segs)
    view = rfm[rfm["Segmento"].isin(pick)]
    c1, c2, c3 = st.columns(3)
    c1.metric("Clientes avaliados", fmt_int(len(view)))
    c2.metric("Mediana de Recência (dias)", fmt_int(np.nanmedian(view["RecenciaDias"])) if len(view)>0 else "-")
    c3.metric("Mediana de Valor (R$)", fmt_money(np.nanmedian(view["Valor"])) if len(view)>0 else "-")
    cols = ["Nome Cliente","RecenciaDias","Frequencia","Valor","R_Score","F_Score","M_Score","Score","Segmento"]
    display_table(view[cols], money_cols=["Valor"], int_cols=["RecenciaDias","Frequencia","Score"])
    try:
        scat = alt.Chart(view.reset_index(drop=True)).mark_circle(size=70).encode(
            x=alt.X("Frequencia:Q", title="Frequência"),
            y=alt.Y("Valor:Q", title="Valor (R$)"),
            color=alt.Color("Segmento:N"),
            tooltip=["Nome Cliente","Frequencia", alt.Tooltip("Valor:Q", format=",.0f"), "RecenciaDias","Segmento"]
        ).properties(height=420)
        st.altair_chart(scat, use_container_width=True)
    except Exception:
        pass

# ---------------- Rentabilidade ----------------
with tab_profit:
    st.subheader("Rentabilidade – Lucro e Margem")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Lucro Bruto Total", fmt_money(flt["Lucro Bruto"].sum()) if "Lucro Bruto" in flt.columns else "-")
    if "Valor Pedido R$" in flt.columns and flt["Valor Pedido R$"].sum()>0:
        margem_total = 100.0*flt["Lucro Bruto"].sum()/flt["Valor Pedido R$"].sum()
        c2.metric("Margem Bruta Total", fmt_pct_safe(margem_total))
    else:
        c2.metric("Margem Bruta Total", "-")
    c3.metric("Ticket de Margem", fmt_money(flt["Lucro Bruto"].sum()/flt["Pedido"].nunique()) if "Pedido" in flt.columns and flt["Pedido"].nunique()>0 else "-")
    c4.metric("% Linhas Negativas", fmt_pct_safe(100.0*(flt["Lucro Bruto"]<0).mean()) if "Lucro Bruto" in flt.columns and len(flt)>0 else "-")

    if {"Nome Cliente","Lucro Bruto"}.issubset(flt.columns):
        st.markdown("#### Top 20 – **Clientes** por Lucro Bruto")
        top_cli = flt.groupby("Nome Cliente", as_index=False)["Lucro Bruto"].sum().sort_values("Lucro Bruto", ascending=False).head(20)
        display_table(top_cli, money_cols=["Lucro Bruto"])
        st.altair_chart(
            alt.Chart(top_cli).mark_bar().encode(
                x=alt.X("Lucro Bruto:Q", title="Lucro Bruto (R$)"),
                y=alt.Y("Nome Cliente:N", sort="-x"),
                tooltip=["Nome Cliente", alt.Tooltip("Lucro Bruto:Q", format=",.0f")]
            ).properties(height=440),
            use_container_width=True
        )

    if {"ITEM","Lucro Bruto"}.issubset(flt.columns):
        st.markdown("#### Top 20 – **SKUs** por Lucro Bruto")
        top_sku = flt.groupby("ITEM", as_index=False)["Lucro Bruto"].sum().sort_values("Lucro Bruto", ascending=False).head(20)
        display_table(top_sku, money_cols=["Lucro Bruto"])
        st.altair_chart(
            alt.Chart(top_sku).mark_bar().encode(
                x=alt.X("Lucro Bruto:Q", title="Lucro Bruto (R$)"),
                y=alt.Y("ITEM:N", sort="-x"),
                tooltip=["ITEM", alt.Tooltip("Lucro Bruto:Q", format=",.0f")]
            ).properties(height=440),
            use_container_width=True
        )

    if {"Representante","Lucro Bruto","Valor Pedido R$"}.issubset(flt.columns):
        st.markdown("#### Margem por Representante")
        por_rep = flt.groupby("Representante", as_index=False).agg({"Lucro Bruto":"sum","Valor Pedido R$":"sum"})
        por_rep["Margem %"] = np.where(por_rep["Valor Pedido R$"]>0, 100.0*por_rep["Lucro Bruto"]/por_rep["Valor Pedido R$"], np.nan)
        por_rep = por_rep.sort_values("Lucro Bruto", ascending=False).head(20)
        display_table(por_rep, money_cols=["Lucro Bruto","Valor Pedido R$"], pct_cols=["Margem %"])

    if {"Nome Cliente","Valor Pedido R$","Lucro Bruto"}.issubset(flt.columns):
        st.markdown("#### Dispersão – Valor x Margem (%) por Cliente")
        disp = flt.groupby("Nome Cliente", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum"})
        disp["Margem %"] = np.where(disp["Valor Pedido R$"]>0, 100.0*disp["Lucro Bruto"]/disp["Valor Pedido R$"], np.nan)
        st.altair_chart(
            alt.Chart(disp).mark_circle(size=70).encode(
                x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
                y=alt.Y("Margem %:Q", title="Margem (%)"),
                tooltip=["Nome Cliente", alt.Tooltip("Valor Pedido R$:Q", format=",.0f"), alt.Tooltip("Margem %:Q", format=",.1f")]
            ).properties(height=420),
            use_container_width=True
        )

    if {"UF","Lucro Bruto","Valor Pedido R$"}.issubset(flt.columns):
        st.markdown("#### Margem por UF")
        por_uf = flt.groupby("UF", as_index=False).agg({"Lucro Bruto":"sum","Valor Pedido R$":"sum"})
        por_uf["Margem %"] = np.where(por_uf["Valor Pedido R$"]>0, 100.0*por_uf["Lucro Bruto"]/por_uf["Valor Pedido R$"], np.nan)
        display_table(por_uf.sort_values("Margem %", ascending=False), money_cols=["Lucro Bruto","Valor Pedido R$"], pct_cols=["Margem %"])

    if "Lucro Bruto" in flt.columns:
        st.markdown("#### Auditoria – Linhas com Margem Negativa")
        neg = flt[flt["Lucro Bruto"] < 0].copy()
        st.caption(f"{len(neg):,}".replace(",", ".") + " linhas com margem negativa no filtro atual.")
        cols_show = [c for c in ["Nome Cliente","Pedido","ITEM","Representante","UF","Valor Pedido R$","Custo","Custo Total","Lucro Bruto","Margem %","Data do Pedido","Data / Mês"] if c in neg.columns]
        display_table(neg[cols_show], money_cols=["Valor Pedido R$","Custo","Custo Total","Lucro Bruto"], pct_cols=["Margem %"])

# ---------------- Clientes ----------------
with tab_cli:
    st.subheader("Clientes – Faturamento")
    if {"Nome Cliente","Valor Pedido R$"}.issubset(flt.columns):
        top_cli = flt.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        display_table(top_cli.head(50), money_cols=["Valor Pedido R$"])

# ---------------- Produtos ----------------
with tab_sku:
    st.subheader("Produtos – Faturamento")
    if {"ITEM","Valor Pedido R$"}.issubset(flt.columns):
        top_sku = flt.groupby("ITEM", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        display_table(top_sku.head(100), money_cols=["Valor Pedido R$"])

# ---------------- Representantes ----------------
with tab_rep:
    st.subheader("Representantes – Faturamento")
    if {"Representante","Valor Pedido R$"}.issubset(flt.columns):
        por_rep_fat = flt.groupby("Representante", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        display_table(por_rep_fat.head(100), money_cols=["Valor Pedido R$"])

# ---------------- Geografia ----------------
with tab_geo:
    st.subheader("Geografia – Faturamento por UF")
    if {"UF","Valor Pedido R$"}.issubset(flt.columns):
        por_uf_fat = flt.groupby("UF", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        display_table(por_uf_fat, money_cols=["Valor Pedido R$"])

# ---------------- Operacional ----------------
with tab_ops:
    st.subheader("Operacional – Lead Time & Atraso")
    c1, c2 = st.columns(2)
    if "LeadTime (dias)" in flt.columns:
        with c1:
            lt = flt["LeadTime (dias)"].dropna()
            if len(lt)>0:
                desc = pd.Series(lt).describe()[["count","mean","50%","min","max"]].rename({"50%":"mediana"})
                display_table(desc.to_frame("LeadTime (dias)").T, int_cols=["count","min","max"])
    if "Atrasado / No prazo" in flt.columns and "Pedido" in flt.columns:
        with c2:
            atrasos = flt.groupby("Atrasado / No prazo", as_index=False)["Pedido"].nunique().rename(columns={"Pedido":"Qtde Pedidos"})
            display_table(atrasos, int_cols=["Qtde Pedidos"])

# ---------------- Pareto / ABC ----------------
with tab_pareto:
    st.subheader("Pareto 80/20 e Curva ABC (Faturamento)")
    if "Valor Pedido R$" in flt.columns:
        if "Nome Cliente" in flt.columns:
            g = flt.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
            g["%Acum"] = 100 * g["Valor Pedido R$"].cumsum() / g["Valor Pedido R$"].sum()
            g["Classe"] = g["%Acum"].apply(lambda p: "A" if p<=80 else ("B" if p<=95 else "C"))
            display_table(g.head(200), money_cols=["Valor Pedido R$"], pct_cols=["%Acum"])
        if "ITEM" in flt.columns:
            s = flt.groupby("ITEM", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
            s["%Acum"] = 100 * s["Valor Pedido R$"].cumsum() / s["Valor Pedido R$"].sum()
            s["Classe"] = s["%Acum"].apply(lambda p: "A" if p<=80 else ("B" if p<=95 else "C"))
            display_table(s.head(300), money_cols=["Valor Pedido R$"], pct_cols=["%Acum"])

# ---------------- SEBASTIAN – cockpit do representante ----------------
with tab_seb:
    st.subheader("SEBASTIAN – Desempenho Individual do Representante")

    if "Representante" not in df.columns or "Valor Pedido R$" not in df.columns:
        st.warning("A aba SEBASTIAN exige pelo menos as colunas 'Representante' e 'Valor Pedido R$'.")
    else:
        reps_all = sorted(df["Representante"].dropna().unique())
        if not reps_all:
            st.info("Nenhum representante encontrado na base.")
        else:
            rep_sel = st.selectbox("Selecione o representante / vendedor / gerente", reps_all)

            if "Data do Pedido" in df.columns and df["Data do Pedido"].notna().any():
                date_col = "Data do Pedido"
            else:
                date_col = "Data / Mês" if "Data / Mês" in df.columns else None

            if date_col is None or d_ini is None:
                st.warning("Para a visão SEBASTIAN funcionar, é necessário ter uma coluna de data (Data do Pedido ou Data / Mês) e o filtro de período configurado.")
            else:
                d_ini_ts = pd.to_datetime(d_ini)
                d_fim_ts = pd.to_datetime(d_fim)

                df_rep_all = df[df["Representante"] == rep_sel].copy()
                df_rep_all = df_rep_all[df_rep_all[date_col].notna()]

                df_rep_period = df_rep_all[
                    (df_rep_all[date_col] >= d_ini_ts) &
                    (df_rep_all[date_col] <= d_fim_ts)
                ].copy()

                if "Status de Produção / Faturamento" in df_rep_period.columns:
                    status_series = df_rep_period["Status de Produção / Faturamento"].astype(str)
                    is_faturado = status_series.str.contains("fatur", case=False, na=False)
                    is_aberto = status_series.str.contains("abert|pend|prod", case=False, na=False) & ~is_faturado
                else:
                    is_faturado = pd.Series([True]*len(df_rep_period), index=df_rep_period.index)
                    is_aberto = pd.Series([False]*len(df_rep_period), index=df_rep_period.index)

                val_faturado = df_rep_period.loc[is_faturado, "Valor Pedido R$"].sum()
                val_total_ped = df_rep_period["Valor Pedido R$"].sum()

                if "Pedido" in df_rep_period.columns:
                    qtd_ped_total = df_rep_period["Pedido"].nunique()
                    qtd_ped_aberto = df_rep_period.loc[is_aberto, "Pedido"].nunique()
                else:
                    qtd_ped_total = len(df_rep_period)
                    qtd_ped_aberto = len(df_rep_period[is_aberto])

                val_ped_aberto = df_rep_period.loc[is_aberto, "Valor Pedido R$"].sum()

                if "Nome Cliente" in df_rep_period.columns:
                    clientes_faturados = df_rep_period.loc[is_faturado, "Nome Cliente"].nunique()
                    clientes_pedidos = df_rep_period["Nome Cliente"].nunique()
                else:
                    clientes_faturados = np.nan
                    clientes_pedidos = np.nan

                if "Nome Cliente" in df_rep_all.columns:
                    clientes_ativos = clientes_pedidos

                    grp_dates = df_rep_all.groupby("Nome Cliente")[date_col]
                    first_buy = grp_dates.min()
                    last_buy = grp_dates.max()

                    novos_mask = (first_buy >= d_ini_ts) & (first_buy <= d_fim_ts)
                    clientes_novos = first_buy[novos_mask].index.tolist()

                    janela_previa_ini = d_ini_ts - pd.DateOffset(months=12)
                    prev_mask = (last_buy >= janela_previa_ini) & (last_buy < d_ini_ts)
                    clientes_prev = set(last_buy[prev_mask].index)

                    clientes_period = set(df_rep_period["Nome Cliente"].unique())
                    clientes_perdidos = sorted(clientes_prev - clientes_period)

                    n_clientes_novos = len(clientes_novos)
                    n_clientes_perdidos = len(clientes_perdidos)
                else:
                    clientes_ativos = np.nan
                    n_clientes_novos = np.nan
                    n_clientes_perdidos = np.nan
                    clientes_novos = []
                    clientes_perdidos = []

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("1a) Faturamento no período (faturado)", fmt_money(val_faturado))
                c2.metric("1b) Valor total de pedidos no período", fmt_money(val_total_ped))
                c3.metric("1c) Pedidos abertos (R$)", fmt_money(val_ped_aberto))
                c4.metric("1c) Pedidos abertos (qtd)", fmt_int(qtd_ped_aberto))

                c5, c6, c7, c8 = st.columns(4)
                c5.metric("1d) Clientes com faturamento", fmt_int(clientes_faturados))
                c6.metric("1e) Clientes com pedidos", fmt_int(clientes_pedidos))
                c7.metric("1f) Clientes ativos no período", fmt_int(clientes_ativos))
                c8.metric("1g) Clientes novos no período", fmt_int(n_clientes_novos))

                c9, _, _, _ = st.columns(4)
                c9.metric("1h) Clientes perdidos/inativados (últimos 12m)", fmt_int(n_clientes_perdidos))

                with st.expander("Clientes novos e clientes perdidos (lista resumida)"):
                    col_n, col_p = st.columns(2)
                    if clientes_novos:
                        col_n.markdown("**Clientes novos no período**")
                        col_n.write(", ".join(sorted(clientes_novos[:50])) + (" ..." if len(clientes_novos) > 50 else ""))
                    else:
                        col_n.caption("Sem clientes novos no período com esse representante.")

                    if clientes_perdidos:
                        col_p.markdown("**Clientes perdidos/inativos (tinham compra nos 12m anteriores e sumiram no período)**")
                        col_p.write(", ".join(clientes_perdidos[:50]) + (" ..." if len(clientes_perdidos) > 50 else ""))
                    else:
                        col_p.caption("Sem clientes claramente perdidos na janela analisada.")

                st.markdown("---")
                st.markdown("### Histórico dos últimos 12 meses")

                janela_12_ini = d_fim_ts - pd.DateOffset(months=12)
                df_rep_12m = df_rep_all[
                    (df_rep_all[date_col] >= janela_12_ini) &
                    (df_rep_all[date_col] <= d_fim_ts)
                ].copy()

                if "Ano-Mes" not in df_rep_12m.columns:
                    df_rep_12m["Ano-Mes"] = df_rep_12m[date_col].dt.to_period("M").astype(str)

                col_h1, col_h2 = st.columns(2)

                with col_h1:
                    if "Pedido" in df_rep_12m.columns:
                        hist_ped = df_rep_12m.groupby("Ano-Mes", as_index=False)["Pedido"].nunique().rename(columns={"Pedido":"Qtde Pedidos"})
                        st.caption("2) Histórico de pedidos (últimos 12 meses)")
                        st.altair_chart(
                            alt.Chart(hist_ped).mark_bar().encode(
                                x=alt.X("Ano-Mes:N", title=None, sort=None),
                                y=alt.Y("Qtde Pedidos:Q", title="Pedidos"),
                                tooltip=["Ano-Mes","Qtde Pedidos"]
                            ).properties(height=280),
                            use_container_width=True
                        )
                    else:
                        st.caption("2) Histórico de pedidos indisponível (coluna 'Pedido' ausente).")

                with col_h2:
                    if "Valor Pedido R$" in df_rep_12m.columns:
                        hist_fat = df_rep_12m.groupby("Ano-Mes", as_index=False)["Valor Pedido R$"].sum()
                        st.caption("3) Histórico de faturamento (últimos 12 meses)")
                        st.altair_chart(
                            alt.Chart(hist_fat).mark_line(point=True).encode(
                                x=alt.X("Ano-Mes:N", title=None, sort=None),
                                y=alt.Y("Valor Pedido R$:Q", title="Faturamento (R$)"),
                                tooltip=["Ano-Mes", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
                            ).properties(height=280),
                            use_container_width=True
                        )
                    else:
                        st.caption("3) Histórico de faturamento indisponível (coluna 'Valor Pedido R$' ausente).")

                st.markdown("### Faturamento por família de produtos – últimos 12 meses")

                col_familia = None
                try:
                    col_familia_global = df.columns[8]  # coluna I da planilha original
                    if col_familia_global in df_rep_12m.columns:
                        col_familia = col_familia_global
                except Exception:
                    col_familia = None

                if col_familia and "Valor Pedido R$" in df_rep_12m.columns:
                    fat_fam = (
                        df_rep_12m
                        .groupby(col_familia, as_index=False)["Valor Pedido R$"]
                        .sum()
                        .sort_values("Valor Pedido R$", ascending=False)
                    )
                    display_table(fat_fam, money_cols=["Valor Pedido R$"])
                    st.altair_chart(
                        alt.Chart(fat_fam.head(20)).mark_bar().encode(
                            x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
                            y=alt.Y(f"{col_familia}:N", sort="-x", title="Família (coluna I – Observação)"),
                            tooltip=[col_familia, alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
                        ).properties(height=400),
                        use_container_width=True
                    )
                else:
                    st.info("Coluna I (Observação) não foi encontrada na base para uso como família.")

                st.markdown("### Faturamento por cliente – comparação")

                if "Nome Cliente" in df_rep_12m.columns and "Valor Pedido R$" in df_rep_12m.columns:
                    fat_cli_12m = df_rep_12m.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().rename(columns={"Valor Pedido R$":"Fat 12m"})
                else:
                    fat_cli_12m = pd.DataFrame(columns=["Nome Cliente","Fat 12m"])

                if "Nome Cliente" in df_rep_period.columns and "Valor Pedido R$" in df_rep_period.columns:
                    fat_cli_periodo = df_rep_period.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().rename(columns={"Valor Pedido R$":"Fat Período"})
                else:
                    fat_cli_periodo = pd.DataFrame(columns=["Nome Cliente","Fat Período"])

                fat_cli_merge = pd.merge(fat_cli_12m, fat_cli_periodo, on="Nome Cliente", how="outer").fillna(0)
                fat_cli_merge["Δ Período vs 12m (R$)"] = fat_cli_merge["Fat Período"] - fat_cli_merge["Fat 12m"]/12.0

                display_table(
                    fat_cli_merge.sort_values("Fat Período", ascending=False).head(50),
                    money_cols=["Fat 12m","Fat Período","Δ Período vs 12m (R$)"]
                )

                st.markdown("### 7) Pedidos em aberto do representante no período selecionado")

                if is_aberto.any():
                    df_aberto = df_rep_period[is_aberto].copy()
                    if "Pedido" in df_aberto.columns:
                        resumo_aberto = df_aberto.groupby("Pedido", as_index=False).agg({
                            "Valor Pedido R$":"sum",
                            "Nome Cliente": "first",
                            date_col: "max"
                        }).rename(columns={"Valor Pedido R$":"Valor Pedido","Nome Cliente":"Cliente","Pedido":"Nº Pedido",date_col:"Data Última Atualização"})
                        display_table(resumo_aberto.sort_values("Valor Pedido", ascending=False), money_cols=["Valor Pedido"])
                    else:
                        display_table(df_aberto, money_cols=["Valor Pedido R$"])
                else:
                    st.caption("Nenhum pedido em aberto para esse representante no período selecionado (considerando a heurística de status).")

# ---------------- Simulador de Vendas – (mantido) ----------------
with tab_sim:
    st.subheader("Simulador de Vendas – multi-SKU com preço manual, elasticidade, impostos, despesas variáveis, alvo de MC e export (CSV/PDF)")

    if qty_col is None or "ITEM" not in df.columns or "Valor Pedido R$" not in df.columns:
        st.warning("Simulador requer colunas 'ITEM', 'Valor Pedido R$' e uma coluna de quantidade (Qtde/QTD/etc.).")
    else:
        itens_disp = sorted(df["ITEM"].dropna().unique())
        default_list = [itens_disp[0]] if itens_disp else []
        skus_sel = st.multiselect("Selecione os SKUs para simulação", itens_disp, default=default_list)

        if not skus_sel:
            st.info("Selecione pelo menos um SKU para simular.")
        else:
            st.markdown("**Base histórica por SKU (toda a base, ignorando filtros):**")

            rows_hist = []
            for sku in skus_sel:
                sku_base = df[df["ITEM"] == sku].copy()
                sku_base[qty_col] = sku_base[qty_col].apply(to_num)

                total_qty_hist = sku_base[qty_col].sum()
                total_fat_hist = sku_base["Valor Pedido R$"].sum()
                total_custo_hist = sku_base["Custo Total"].sum()

                if total_qty_hist and total_qty_hist > 0:
                    preco_medio_hist = total_fat_hist / total_qty_hist
                    custo_medio_hist = total_custo_hist / total_qty_hist
                else:
                    preco_medio_hist = np.nan
                    custo_medio_hist = np.nan

                rows_hist.append({
                    "SKU": sku,
                    "Qtde Histórica": total_qty_hist,
                    "Faturamento Histórico": total_fat_hist,
                    "Preço Médio Histórico": preco_medio_hist,
                    "Custo Médio Histórico": custo_medio_hist,
                })

            hist_df = pd.DataFrame(rows_hist)
            display_table(
                hist_df,
                money_cols=["Faturamento Histórico","Preço Médio Histórico","Custo Médio Histórico"],
                int_cols=["Qtde Histórica"]
            )
            st.caption("Os valores médios são calculados em cima de **todas as vendas históricas** de cada SKU, não do filtro lateral.")

            st.markdown("---")
            st.markdown("### Parametrização da venda simulada por SKU")

            sim_rows = []

            for sku in skus_sel:
                sku_row = hist_df[hist_df["SKU"] == sku].iloc[0]
                total_qty_hist = sku_row["Qtde Histórica"]
                preco_medio_hist = sku_row["Preço Médio Histórico"]
                custo_medio_hist = sku_row["Custo Médio Histórico"]

                with st.expander(f"Configuração do SKU: {sku}"):
                    col_q, col_p, col_c = st.columns(3)
                    with col_q:
                        default_q = int(total_qty_hist) if (not pd.isna(total_qty_hist) and total_qty_hist>0) else 100
                        qtd_sim = st.number_input(
                            "Quantidade simulada",
                            min_value=1,
                            value=default_q,
                            step=10,
                            key=f"q_{sku}"
                        )
                    with col_p:
                        ajuste_preco = st.slider(
                            "Ajuste % no preço unitário vs. histórico",
                            -50, 100, 0, step=5, key=f"ap_{sku}"
                        )
                    with col_c:
                        ajuste_custo = st.slider(
                            "Ajuste % no custo unitário vs. histórico",
                            -50, 100, 0, step=5, key=f"ac_{sku}"
                        )

                    preco_unit_base = preco_medio_hist if pd.notna(preco_medio_hist) else 0.0
                    custo_unit_base = custo_medio_hist if pd.notna(custo_medio_hist) else 0.0

                    preco_unit_sim_hist = preco_unit_base * (1 + ajuste_preco/100.0)
                    custo_unit_sim = custo_unit_base * (1 + ajuste_custo/100.0)

                    preco_unit_manual = st.number_input(
                        "Preço unitário da simulação (R$) – se 0, usa preço histórico ajustado",
                        min_value=0.0,
                        value=float(round(preco_unit_sim_hist, 2)) if preco_unit_sim_hist > 0 else 0.0,
                        step=0.1,
                        key=f"pm_{sku}"
                    )
                    preco_unit_final = preco_unit_manual if preco_unit_manual > 0 else preco_unit_sim_hist

                    c1, c2 = st.columns(2)
                    c1.metric("Preço unitário usado na simulação", fmt_money(preco_unit_final))
                    c2.metric("Custo unitário simulado", fmt_money(custo_unit_sim))

                    faturamento_sim_sku = qtd_sim * preco_unit_final
                    custo_total_sim_sku = qtd_sim * custo_unit_sim
                    lucro_bruto_sim_sku = faturamento_sim_sku - custo_total_sim_sku
                    margem_bruta_sim_sku = 100*lucro_bruto_sim_sku/faturamento_sim_sku if faturamento_sim_sku>0 else np.nan

                    st.caption(
                        f"Faturamento simulado do SKU {sku}: {fmt_money(faturamento_sim_sku)} | "
                        f"Margem bruta: {fmt_pct_safe(margem_bruta_sim_sku) if not pd.isna(margem_bruta_sim_sku) else '-'}"
                    )

                    sim_rows.append({
                        "SKU": sku,
                        "Qtd Simulada": qtd_sim,
                        "Preço Unitário Simulado": preco_unit_final,
                        "Custo Unitário Simulado": custo_unit_sim,
                        "Faturamento Simulado": faturamento_sim_sku,
                        "Custo Total Simulado": custo_total_sim_sku,
                        "Lucro Bruto Simulado": lucro_bruto_sim_sku,
                        "Margem Bruta %": margem_bruta_sim_sku
                    })

            if not sim_rows:
                st.warning("Nenhum SKU configurado para simulação.")
            else:
                sim_df = pd.DataFrame(sim_rows)

                st.markdown("### Impostos, despesas variáveis e alvo de Margem de Contribuição")

                col_i1, col_i2, col_i3, col_i4 = st.columns(4)
                with col_i1:
                    icms_pct = st.number_input("ICMS (%)", min_value=0.0, max_value=50.0, value=18.0, step=0.5)
                with col_i2:
                    pis_pct = st.number_input("PIS (%)", min_value=0.0, max_value=10.0, value=1.65, step=0.05)
                with col_i3:
                    cofins_pct = st.number_input("COFINS (%)", min_value=0.0, max_value=10.0, value=7.60, step=0.05)
                with col_i4:
                    outros_pct = st.number_input("Outros impostos (%)", min_value=0.0, max_value=30.0, value=0.0, step=0.5)

                col_d1, col_d2, col_m = st.columns(3)
                with col_d1:
                    frete_pct = st.number_input("Fretes (% do faturamento)", min_value=0.0, max_value=30.0, value=0.0, step=0.5)
                with col_d2:
                    comissao_pct = st.number_input("Comissões (% do faturamento)", min_value=0.0, max_value=30.0, value=0.0, step=0.5)
                with col_m:
                    margem_target_pct = st.number_input("MC mínima desejada (%)", min_value=0.0, max_value=80.0, value=15.0, step=0.5)

                faturamento_sim = sim_df["Faturamento Simulado"].sum()
                custo_total_sim = sim_df["Custo Total Simulado"].sum()
                lucro_bruto_sim = faturamento_sim - custo_total_sim
                margem_bruta_sim = 100*lucro_bruto_sim/faturamento_sim if faturamento_sim>0 else np.nan

                imposto_icms = faturamento_sim * icms_pct/100.0
                imposto_pis = faturamento_sim * pis_pct/100.0
                imposto_cofins = faturamento_sim * cofins_pct/100.0
                imposto_outros = faturamento_sim * outros_pct/100.0
                imposto_total = imposto_icms + imposto_pis + imposto_cofins + imposto_outros

                frete_val = faturamento_sim * frete_pct/100.0
                comissao_val = faturamento_sim * comissao_pct/100.0
                desp_var_total = frete_val + comissao_val

                receita_liq = faturamento_sim - imposto_total
                margem_contrib = receita_liq - custo_total_sim - desp_var_total
                margem_contrib_pct = 100*margem_contrib/receita_liq if receita_liq>0 else np.nan

                T_imp = (icms_pct + pis_pct + cofins_pct + outros_pct) / 100.0
                T_d   = (frete_pct + comissao_pct) / 100.0
                M     = margem_target_pct / 100.0

                A = (1 - M) * (1 - T_imp) - T_d

                if faturamento_sim > 0 and A <= 0:
                    st.warning(
                        "Com os impostos, fretes, comissões e a MC mínima desejada informados, "
                        "não existe preço viável matematicamente (A ≤ 0). Reduza a MC alvo ou revise percentuais."
                    )
                    sim_df["Preço Unit. Mínimo (MC alvo)"] = np.nan
                    sim_df["Desc Máx vs Preço Sim (%)"] = np.nan
                else:
                    sim_df["Preço Unit. Mínimo (MC alvo)"] = sim_df["Custo Unitário Simulado"] / A
                    sim_df["Desc Máx vs Preço Sim (%)"] = np.where(
                        sim_df["Preço Unitário Simulado"] > 0,
                        100 * (1 - sim_df["Preço Unit. Mínimo (MC alvo)"] / sim_df["Preço Unitário Simulado"]),
                        np.nan
                    )

                st.markdown("### Resumo por SKU – venda simulada")
                display_table(
                    sim_df,
                    money_cols=[
                        "Preço Unitário Simulado",
                        "Custo Unitário Simulado",
                        "Preço Unit. Mínimo (MC alvo)",
                        "Faturamento Simulado",
                        "Custo Total Simulado",
                        "Lucro Bruto Simulado"
                    ],
                    pct_cols=["Margem Bruta %","Desc Máx vs Preço Sim (%)"],
                    int_cols=["Qtd Simulada"]
                )

                st.markdown("### KPIs consolidados da venda simulada (todos os SKUs)")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Faturamento simulado (total)", fmt_money(faturamento_sim))
                c2.metric("Custo total simulado", fmt_money(custo_total_sim))
                c3.metric("Lucro bruto (antes impostos)", fmt_money(lucro_bruto_sim))
                c4.metric("Margem bruta (antes impostos)", fmt_pct_safe(margem_bruta_sim) if pd.notna(margem_bruta_sim) else "-")

                st.markdown("### Margem de Contribuição consolidada vs alvo")
                c5, c6 = st.columns(2)
                if pd.notna(margem_contrib_pct):
                    delta_pp = margem_contrib_pct - margem_target_pct
                    delta_txt = f"{delta_pp:.1f} p.p.".replace(".", ",")
                    c5.metric("MC atual do cenário", fmt_pct_safe(margem_contrib_pct, 1), delta=delta_txt)
                else:
                    c5.metric("MC atual do cenário", "-")
                c6.metric("MC mínima desejada", fmt_pct_safe(margem_target_pct, 1))

                st.markdown("### Mini DRE – Venda simulada (todos os SKUs, até Margem de Contribuição)")
                dre = pd.DataFrame({
                    "Linha": [
                        "Faturamento bruto",
                        f"(-) ICMS ({fmt_pct_safe(icms_pct,1)})",
                        f"(-) PIS ({fmt_pct_safe(pis_pct,2)})",
                        f"(-) COFINS ({fmt_pct_safe(cofins_pct,2)})",
                        f"(-) Outros impostos ({fmt_pct_safe(outros_pct,1)})",
                        f"(-) Fretes ({fmt_pct_safe(frete_pct,1)})",
                        f"(-) Comissões ({fmt_pct_safe(comissao_pct,1)})",
                        "Receita Líquida",
                        "(-) Custo dos Produtos",
                        "Margem de Contribuição (R$)",
                        "Margem de Contribuição (%)"
                    ],
                    "Valor": [
                        faturamento_sim,
                        -imposto_icms,
                        -imposto_pis,
                        -imposto_cofins,
                        -imposto_outros,
                        -frete_val,
                        -comissao_val,
                        receita_liq - desp_var_total,
                        -custo_total_sim,
                        margem_contrib,
                        margem_contrib_pct
                    ]
                })

                dre_view = dre.copy()
                for idx, row in dre_view.iterrows():
                    if row["Linha"] == "Margem de Contribuição (%)":
                        dre_view.loc[idx, "Valor"] = fmt_pct_safe(row["Valor"], 1)
                    else:
                        dre_view.loc[idx, "Valor"] = fmt_money(row["Valor"])

                st.table(dre_view)

                st.markdown("---")
                st.markdown("### Análise de elasticidade preço × volume (cenários)")

                col_e1, col_e2, col_e3 = st.columns(3)
                with col_e1:
                    ativar_elast = st.checkbox("Ativar análise de elasticidade", value=False)
                with col_e2:
                    var_min = st.number_input("Variação mínima de preço (%)", min_value=-50.0, max_value=0.0, value=-10.0, step=1.0)
                with col_e3:
                    var_max = st.number_input("Variação máxima de preço (%)", min_value=0.0, max_value=50.0, value=10.0, step=1.0)

                elast_df = None

                if ativar_elast:
                    if var_max <= var_min:
                        st.warning("A variação máxima de preço deve ser maior que a mínima.")
                    else:
                        col_e4, col_e5 = st.columns(2)
                        with col_e4:
                            n_cenarios = st.slider("Quantidade de cenários", min_value=5, max_value=21, value=9, step=2)
                        with col_e5:
                            elasticidade_manual = st.number_input(
                                "Elasticidade de volume padrão (tipicamente negativa, ex: -1.5)",
                                min_value=-5.0, max_value=1.0, value=-1.5, step=0.1
                            )

                        elastic_all = estimate_elasticities(df, qty_col) if qty_col is not None else None
                        elast_map = {}
                        modo_elast = "Usar valor manual único"

                        if elastic_all is not None and not elastic_all.empty:
                            sub_elast = elastic_all[elastic_all["SKU"].isin(skus_sel)].copy()
                            if len(sub_elast) > 0:
                                st.markdown("#### Elasticidade histórica estimada por SKU (base real)")
                                display_table(
                                    sub_elast.rename(columns={
                                        "Elasticidade": "Elasticidade Estimada",
                                        "N_Obs": "N Observações"
                                    }),
                                    int_cols=["N Observações"]
                                )
                                modo_elast = st.radio(
                                    "Como aplicar elasticidade nos cenários?",
                                    ["Usar valor manual único", "Usar elasticidade histórica por SKU (fallback para valor manual)"],
                                    index=1,
                                )
                                elast_map = dict(zip(sub_elast["SKU"], sub_elast["Elasticidade"]))

                        deltas = np.linspace(var_min, var_max, n_cenarios)
                        rows_elast = []

                        base_qtd = sim_df["Qtd Simulada"].astype(float)
                        base_preco = sim_df["Preço Unitário Simulado"].astype(float)
                        base_custo_unit = sim_df["Custo Unitário Simulado"].astype(float)
                        base_skus = sim_df["SKU"].astype(str).tolist()

                        for d in deltas:
                            fator_preco = 1 + d/100.0

                            if modo_elast.startswith("Usar elasticidade histórica") and elast_map:
                                fatores_volume = []
                                for sku in base_skus:
                                    e_sku = elast_map.get(sku, elasticidade_manual)
                                    fv = max(1 + e_sku * (d/100.0), 0.0)
                                    fatores_volume.append(fv)
                                fatores_volume = np.array(fatores_volume)
                                qtd_cenario = base_qtd.values * fatores_volume
                            else:
                                fator_volume = max(1 + elasticidade_manual * (d/100.0), 0.0)
                                qtd_cenario = base_qtd * fator_volume

                            preco_cenario = base_preco * fator_preco

                            fat_cenario = (qtd_cenario * preco_cenario).sum()
                            custo_cenario = (qtd_cenario * base_custo_unit).sum()
                            lucro_bruto_cenario = fat_cenario - custo_cenario

                            imposto_icms_c = fat_cenario * icms_pct/100.0
                            imposto_pis_c = fat_cenario * pis_pct/100.0
                            imposto_cofins_c = fat_cenario * cofins_pct/100.0
                            imposto_outros_c = fat_cenario * outros_pct/100.0
                            imposto_total_c = imposto_icms_c + imposto_pis_c + imposto_cofins_c + imposto_outros_c

                            frete_c = fat_cenario * frete_pct/100.0
                            comissao_c = fat_cenario * comissao_pct/100.0
                            desp_var_c = frete_c + comissao_c

                            receita_liq_c = fat_cenario - imposto_total_c
                            margem_contrib_c = receita_liq_c - custo_cenario - desp_var_c
                            margem_contrib_pct_c = 100*margem_contrib_c/receita_liq_c if receita_liq_c>0 else np.nan

                            rows_elast.append({
                                "Δ Preço (%)": d,
                                "Faturamento": fat_cenario,
                                "Lucro Bruto": lucro_bruto_cenario,
                                "Margem Contribuição (R$)": margem_contrib_c,
                                "Margem Contribuição (%)": margem_contrib_pct_c
                            })

                        elast_df = pd.DataFrame(rows_elast)

                        st.markdown("#### Tabela de cenários – preço × volume × margem")
                        display_table(
                            elast_df,
                            money_cols=["Faturamento","Lucro Bruto","Margem Contribuição (R$)"],
                            pct_cols=["Margem Contribuição (%)"]
                        )

                        st.markdown("#### Curva de Margem de Contribuição (R$) por variação de preço")
                        try:
                            chart_mc = alt.Chart(elast_df).mark_line(point=True).encode(
                                x=alt.X("Δ Preço (%):Q", title="Variação de preço (%)"),
                                y=alt.Y("Margem Contribuição (R$):Q", title="MC (R$)"),
                                tooltip=[
                                    alt.Tooltip("Δ Preço (%):Q", format=".1f"),
                                    alt.Tooltip("Faturamento:Q", format=",.0f"),
                                    alt.Tooltip("Margem Contribuição (R$):Q", format=",.0f"),
                                    alt.Tooltip("Margem Contribuição (%):Q", format=".1f")
                                ]
                            ).properties(height=380)
                            st.altair_chart(chart_mc, use_container_width=True)
                        except Exception:
                            pass

                        st.caption(
                            "A elasticidade define como o **volume reage à variação de preço**. "
                            "No modo histórico, cada SKU usa sua curva real sempre que há dados suficientes; "
                            "quando não há, caímos no valor manual padrão."
                        )

                st.markdown("---")
                st.markdown("### Exportar simulação de vendas")

                export_df = sim_df.copy()
                export_df["ICMS (%)"] = icms_pct
                export_df["PIS (%)"] = pis_pct
                export_df["COFINS (%)"] = cofins_pct
                export_df["Outros Impostos (%)"] = outros_pct
                export_df["Frete (%)"] = frete_pct
                export_df["Comissão (%)"] = comissao_pct
                export_df["MC alvo (%)"] = margem_target_pct
                export_df["Faturamento Total Simulado"] = faturamento_sim
                export_df["MC Total (R$)"] = margem_contrib
                export_df["MC Total (%)"] = margem_contrib_pct
                export_df["Tem Análise de Elasticidade"] = elast_df is not None

                csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")

                col_exp1, col_exp2 = st.columns(2)
                with col_exp1:
                    st.download_button(
                        "Baixar simulação (CSV)",
                        data=csv_bytes,
                        file_name="simulacao_venda_brasforma.csv",
                        mime="text/csv"
                    )

                with col_exp2:
                    try:
                        pdf_bytes = build_simulation_pdf(
                            sim_df=sim_df,
                            faturamento_sim=faturamento_sim,
                            margem_contrib=margem_contrib,
                            margem_contrib_pct=margem_contrib_pct,
                            icms_pct=icms_pct,
                            pis_pct=pis_pct,
                            cofins_pct=cofins_pct,
                            outros_pct=outros_pct,
                            frete_pct=frete_pct,
                            comissao_pct=comissao_pct,
                            margem_target_pct=margem_target_pct
                        )
                        st.download_button(
                            "Baixar simulação (PDF)",
                            data=pdf_bytes,
                            file_name="simulacao_venda_brasforma.pdf",
                            mime="application/pdf"
                        )
                    except Exception:
                        st.warning("Falha ao gerar PDF. Verifique se o pacote 'fpdf2' está instalado no ambiente.")

# ---------------- Export ----------------
with tab_export:
    st.subheader("Exportar")
    st.download_button("Baixar CSV filtrado", data=flt.to_csv(index=False).encode("utf-8-sig"), file_name="brasforma_filtrado.csv", mime="text/csv")
    with st.expander("Prévia dos dados filtrados"):
        st.dataframe(flt)

if qty_col:
    st.caption(f"✓ Custo calculado como **unitário × quantidade**. Coluna de quantidade detectada: **{qty_col}**.")
else:
    st.caption("! Atenção: coluna de quantidade não identificada — usando Custo como total.")
