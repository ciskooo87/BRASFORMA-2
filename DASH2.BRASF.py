# streamlit_app_brasforma_v22.py
# Brasforma – Dashboard Comercial v22
# Atualização: filtro "Transação" + branding com logo (page icon, header e sidebar)
# Mantém todas as abas e funcionalidades já construídas.

import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path
from fpdf import FPDF
from PIL import Image

# ===================== CONFIG & BRANDING =====================
LOGO_PATH = "/mnt/data/images.png"     # ajuste se necessário
LOGO_FALLBACK = "logo_brasforma.png"   # fallback local opcional

st.set_page_config(
    page_title="Brasforma – Dashboard Comercial v22",
    layout="wide",
    page_icon=LOGO_PATH if Path(LOGO_PATH).exists() else "🧭",
)

def load_logo():
    for p in [LOGO_PATH, LOGO_FALLBACK]:
        try:
            return Image.open(p)
        except Exception:
            continue
    return None

APP_LOGO = load_logo()

# ===================== HELPERS =====================
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
    try:
        return f"{int(v):,}".replace(",", ".")
    except Exception:
        return "-"

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

# ===================== LOADERS =====================
@st.cache_data(show_spinner=False)
def load_data(path: str, sheet_name="Carteira de Vendas"):
    xls = pd.ExcelFile(path, engine="openpyxl")
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = [c.strip() for c in df.columns]

    # normaliza "Transação" caso venha sem acento
    if "Transacao" in df.columns and "Transação" not in df.columns:
        df.rename(columns={"Transacao": "Transação"}, inplace=True)

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

    # detectar coluna de quantidade
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
        if qty_col is not None and qty_col in df.columns:
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

@st.cache_data(show_spinner=False)
def load_goals(path="Metas_Brasforma.xlsx", sheet_name="Metas"):
    p = Path(path)
    if not p.exists():
        return None
    xls = pd.ExcelFile(p)
    target_sheet = None
    for sn in xls.sheet_names:
        if sn.strip().lower() == sheet_name.lower():
            target_sheet = sn
            break
    if target_sheet is None:
        return None
    metas = pd.read_excel(xls, sheet_name=target_sheet)
    metas.columns = [c.strip() for c in metas.columns]
    if "Meta_Faturamento" not in metas.columns:
        for c in metas.columns:
            if "meta" in c.lower() and ("fat" in c.lower() or "fatur" in c.lower()):
                metas = metas.rename(columns={c: "Meta_Faturamento"})
                break
    required = {"Ano","Mes","Representante","Meta_Faturamento"}
    if not required.issubset(metas.columns):
        return None
    metas["Ano"] = pd.to_numeric(metas["Ano"], errors="coerce").astype("Int64")
    metas["Mes"] = pd.to_numeric(metas["Mes"], errors="coerce").astype("Int64")
    metas["Meta_Faturamento"] = pd.to_numeric(metas["Meta_Faturamento"], errors="coerce")
    metas = metas.dropna(subset=["Ano","Mes","Representante","Meta_Faturamento"])
    metas["Representante"] = metas["Representante"].astype(str).str.strip()
    if "Meta_Margem_Bruta" in metas.columns:
        metas["Meta_Margem_Bruta"] = pd.to_numeric(metas["Meta_Margem_Bruta"], errors="coerce")
    return metas

# ===================== HEADER COM LOGO =====================
with st.container():
    cL, cT = st.columns([1, 6])
    with cL:
        if APP_LOGO is not None:
            st.image(APP_LOGO, use_container_width=True)
    with cT:
        st.markdown("### **Dashboard Comercial – Brasforma**")
        st.caption("Inteligência comercial executiva • Vendas • Metas & Forecast • Rentabilidade • Operação")

# ===================== FONTE DE DADOS =====================
DEFAULT_DATA = "Dashboard - Comite Semanal - Brasforma IA (1).xlsx"
ALT_DATA = "Dashboard - Comite Semanal - Brasforma (1).xlsx"

st.sidebar.image(APP_LOGO, use_column_width=True)
st.sidebar.title("Fonte de dados")
uploaded = st.sidebar.file_uploader("Base (.xlsx)", type=["xlsx"], accept_multiple_files=False)
data_path = uploaded if uploaded is not None else (DEFAULT_DATA if Path(DEFAULT_DATA).exists() else ALT_DATA)
st.sidebar.caption(f"Arquivo em uso: **{data_path}**")

df, qty_col = load_data(data_path)
goals_df = load_goals()

# ===================== FILTROS GLOBAIS =====================
st.sidebar.title("Filtros")

if "Data do Pedido" in df.columns and df["Data do Pedido"].notna().any():
    date_col_main = "Data do Pedido"
elif "Data / Mês" in df.columns:
    date_col_main = "Data / Mês"
else:
    date_col_main = None

if date_col_main is not None:
    min_date = pd.to_datetime(df[date_col_main]).min()
    max_date = pd.to_datetime(df[date_col_main]).max()
    d_ini, d_fim = st.sidebar.date_input("Período (data principal)", value=(min_date, max_date))
else:
    d_ini = d_fim = None

reg = st.sidebar.multiselect("Regional", sorted(df["Regional"].dropna().unique()) if "Regional" in df.columns else [])
rep = st.sidebar.multiselect("Representante", sorted(df["Representante"].dropna().unique()) if "Representante" in df.columns else [])
uf  = st.sidebar.multiselect("UF", sorted(df["UF"].dropna().unique()) if "UF" in df.columns else [])
stat = st.sidebar.multiselect("Status Prod./Fat.", sorted(df["Status de Produção / Faturamento"].dropna().unique()) if "Status de Produção / Faturamento" in df.columns else [])
trans = st.sidebar.multiselect("Transação", sorted(df["Transação"].dropna().unique()) if "Transação" in df.columns else [])
cliente = st.sidebar.text_input("Cliente (contém)")
item = st.sidebar.text_input("SKU/Item (contém)")
show_neg = st.sidebar.checkbox("Mostrar apenas linhas com margem negativa", value=False)

# opcional: sinal de devoluções (subtrai do faturamento)
invert_devol = st.sidebar.checkbox("Subtrair devoluções do faturamento", value=False)

def apply_filters(_df):
    flt = _df.copy()
    # aplica sinal de devolução antes de agregar
    if invert_devol and "Transação" in flt.columns and "Valor Pedido R$" in flt.columns:
        devol_mask = flt["Transação"].astype(str).str.contains("devol", case=False, na=False)
        flt.loc[devol_mask, "Valor Pedido R$"] = -flt.loc[devol_mask, "Valor Pedido R$"].abs()
        if "Lucro Bruto" in flt.columns:
            # recomputa lucro e margem com sinal invertido
            flt["Lucro Bruto"] = flt["Valor Pedido R$"] - flt["Custo Total"]
            flt["Margem %"] = np.where(flt["Valor Pedido R$"] != 0,
                                       100 * flt["Lucro Bruto"] / flt["Valor Pedido R$"], np.nan)

    if date_col_main is not None and d_ini is not None:
        flt = flt[(flt[date_col_main] >= pd.to_datetime(d_ini)) & (flt[date_col_main] <= pd.to_datetime(d_fim))]
    if reg:
        flt = flt[flt["Regional"].isin(reg)]
    if rep:
        flt = flt[flt["Representante"].isin(rep)]
    if uf:
        flt = flt[flt["UF"].isin(uf)]
    if stat:
        flt = flt[flt["Status de Produção / Faturamento"].isin(stat)]
    if trans:
        flt = flt[flt["Transação"].isin(trans)]
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

# ===================== TABS =====================
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
(tab_dir, tab_exec, tab_rfm, tab_profit, tab_cli, tab_sku,
 tab_rep, tab_geo, tab_ops, tab_pareto, tab_seb, tab_sim, tab_export) = tabs

# ===================== DIRETORIA – METAS & FORECAST =====================
with tab_dir:
    st.subheader("Diretoria – Metas & Forecast (mensal)")
    if goals_df is None:
        st.info("Metas_Brasforma.xlsx não encontrado ou com colunas inválidas (Ano, Mes, Representante, Meta_Faturamento).")
    else:
        if date_col_main is None:
            st.warning("Não encontrei coluna de data para forecast.")
        else:
            ref_date = pd.to_datetime(d_fim) if d_fim is not None else pd.to_datetime(df[date_col_main].dropna().max())
            ano_ref, mes_ref = ref_date.year, ref_date.month
            st.caption(f"Mês de referência: **{mes_ref:02d}/{ano_ref}**")

            metas_ref = goals_df[(goals_df["Ano"] == ano_ref) & (goals_df["Mes"] == mes_ref)].copy()
            if metas_ref.empty:
                st.info("Sem metas para o mês de referência.")
            else:
                df_month = df[df[date_col_main].dt.to_period("M") == pd.Period(ref_date, "M")].copy()
                df_month = df_month[df_month[date_col_main] <= ref_date]

                dias_mes = pd.Period(ref_date, "M").days_in_month
                dias_passados = ref_date.day

                if "Valor Pedido R$" in df_month.columns:
                    fat_real_total = df_month["Valor Pedido R$"].sum()
                else:
                    fat_real_total = np.nan
                fat_fore_total = fat_real_total / dias_passados * dias_mes if dias_passados > 0 else np.nan
                meta_total = metas_ref["Meta_Faturamento"].sum()
                ating_fore = 100 * fat_fore_total / meta_total if meta_total > 0 else np.nan

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Meta do mês (total)", fmt_money(meta_total))
                c2.metric(f"Realizado até {ref_date.strftime('%d/%m')}", fmt_money(fat_real_total))
                c3.metric("Forecast do mês", fmt_money(fat_fore_total))
                c4.metric("Atingimento projetado", fmt_pct_safe(ating_fore) if pd.notna(ating_fore) else "-")

                st.markdown("### Metas por hierarquia")
                nivel = st.radio("Nível", ["Regional", "Representante"], index=0, horizontal=True)

                # mapa rep->regional
                if {"Representante", "Regional"}.issubset(df.columns):
                    map_rep_reg = df[["Representante","Regional"]].dropna(subset=["Representante"]).drop_duplicates("Representante")
                else:
                    map_rep_reg = pd.DataFrame(columns=["Representante","Regional"])

                # realizado/forecast por rep
                if "Representante" in df_month.columns and "Valor Pedido R$" in df_month.columns:
                    real_rep = df_month.groupby("Representante", as_index=False)["Valor Pedido R$"].sum().rename(columns={"Valor Pedido R$":"Realizado"})
                else:
                    real_rep = pd.DataFrame(columns=["Representante","Realizado"])
                real_rep["Realizado"] = pd.to_numeric(real_rep["Realizado"], errors="coerce").fillna(0.0)
                real_rep["Forecast"] = real_rep["Realizado"] / dias_passados * dias_mes if dias_passados>0 else real_rep["Realizado"]

                metas_ref_rep = metas_ref.groupby("Representante", as_index=False)["Meta_Faturamento"].sum()
                metas_ref_rep["Representante"] = metas_ref_rep["Representante"].astype(str).str.strip()

                painel_rep = real_rep.merge(metas_ref_rep, on="Representante", how="outer")
                painel_rep["Realizado"] = painel_rep["Realizado"].fillna(0.0)
                painel_rep["Forecast"] = painel_rep["Forecast"].fillna(painel_rep["Realizado"])
                painel_rep["Meta_Faturamento"] = painel_rep["Meta_Faturamento"].fillna(0.0)
                painel_rep["Atingimento Atual (%)"] = np.where(painel_rep["Meta_Faturamento"]>0, 100*painel_rep["Realizado"]/painel_rep["Meta_Faturamento"], np.nan)
                painel_rep["Atingimento Forecast (%)"] = np.where(painel_rep["Meta_Faturamento"]>0, 100*painel_rep["Forecast"]/painel_rep["Meta_Faturamento"], np.nan)
                painel_rep["GAP (R$)"] = painel_rep["Meta_Faturamento"] - painel_rep["Forecast"]

                if not map_rep_reg.empty:
                    painel_rep = painel_rep.merge(map_rep_reg, on="Representante", how="left")
                else:
                    painel_rep["Regional"] = "SEM REGIONAL"

                if nivel == "Regional":
                    painel_reg = painel_rep.groupby("Regional", as_index=False).agg({
                        "Meta_Faturamento":"sum","Realizado":"sum","Forecast":"sum"
                    })
                    painel_reg["Atingimento Atual (%)"] = np.where(painel_reg["Meta_Faturamento"]>0, 100*painel_reg["Realizado"]/painel_reg["Meta_Faturamento"], np.nan)
                    painel_reg["Atingimento Forecast (%)"] = np.where(painel_reg["Meta_Faturamento"]>0, 100*painel_reg["Forecast"]/painel_reg["Meta_Faturamento"], np.nan)
                    painel_reg["GAP (R$)"] = painel_reg["Meta_Faturamento"] - painel_reg["Forecast"]

                    st.markdown("#### Regional (mês de referência)")
                    display_table(painel_reg.sort_values("Meta_Faturamento", ascending=False),
                                  money_cols=["Meta_Faturamento","Realizado","Forecast","GAP (R$)"],
                                  pct_cols=["Atingimento Atual (%)","Atingimento Forecast (%)"])
                    try:
                        chart_long = painel_reg.melt(id_vars=["Regional"], value_vars=["Meta_Faturamento","Forecast","Realizado"], var_name="Tipo", value_name="Valor")
                        st.altair_chart(
                            alt.Chart(chart_long).mark_bar().encode(
                                x=alt.X("Valor:Q", title="R$"), y=alt.Y("Regional:N", sort="-x"),
                                color=alt.Color("Tipo:N", title=None),
                                tooltip=["Regional","Tipo", alt.Tooltip("Valor:Q", format=",.0f")]
                            ).properties(height=420),
                            use_container_width=True
                        )
                    except Exception:
                        pass

                    regionais = sorted(painel_reg["Regional"].dropna().unique())
                    if regionais:
                        reg_sel = st.selectbox("Detalhar representantes de qual regional?", options=["(todas)"]+regionais)
                        det = painel_rep if reg_sel=="(todas)" else painel_rep[painel_rep["Regional"]==reg_sel]
                        with st.expander("Detalhamento por representante"):
                            display_table(det.sort_values("Meta_Faturamento", ascending=False),
                                          money_cols=["Meta_Faturamento","Realizado","Forecast","GAP (R$)"],
                                          pct_cols=["Atingimento Atual (%)","Atingimento Forecast (%)"])
                else:
                    st.markdown("#### Representantes (mês de referência)")
                    regionais = sorted(painel_rep["Regional"].dropna().unique())
                    reg_sel_multi = st.multiselect("Filtrar por regional (opcional)", options=regionais)
                    painel_rep_view = painel_rep if not reg_sel_multi else painel_rep[painel_rep["Regional"].isin(reg_sel_multi)]
                    display_table(painel_rep_view.sort_values("Meta_Faturamento", ascending=False),
                                  money_cols=["Meta_Faturamento","Realizado","Forecast","GAP (R$)"],
                                  pct_cols=["Atingimento Atual (%)","Atingimento Forecast (%)"])
                    st.markdown("##### Gráfico – Top N representantes")
                    col_n, col_metric = st.columns([1,1])
                    with col_n:
                        top_n = st.slider("Qtd de reps no gráfico", 5, 50, 20, step=5)
                    with col_metric:
                        metric_opt = st.selectbox("Ordenar por", options=["Meta_Faturamento","Realizado","Forecast","GAP (R$)"], index=1)
                    chart_rep = painel_rep_view.sort_values(metric_opt, ascending=False).head(top_n)
                    try:
                        chart_long = chart_rep.melt(id_vars=["Representante"], value_vars=["Meta_Faturamento","Forecast","Realizado"], var_name="Tipo", value_name="Valor")
                        st.altair_chart(
                            alt.Chart(chart_long).mark_bar().encode(
                                x=alt.X("Valor:Q", title="R$"), y=alt.Y("Representante:N", sort="-x"),
                                color=alt.Color("Tipo:N", title=None),
                                tooltip=["Representante","Tipo", alt.Tooltip("Valor:Q", format=",.0f")]
                            ).properties(height=480),
                            use_container_width=True
                        )
                    except Exception:
                        pass

                dias_restantes = pd.Period(ref_date, "M").days_in_month - ref_date.day
                if dias_restantes > 0:
                    painel_rep2 = painel_rep.copy()
                    painel_rep2["Necessidade por dia útil (R$)"] = np.where(
                        painel_rep2["Meta_Faturamento"]>painel_rep2["Forecast"],
                        (painel_rep2["Meta_Faturamento"]-painel_rep2["Forecast"])/dias_restantes,
                        0.0,
                    )
                    with st.expander("Necessidade diária por representante (para bater a meta)"):
                        display_table(
                            painel_rep2.sort_values("Necessidade por dia útil (R$)", ascending=False),
                            money_cols=["Necessidade por dia útil (R$)","GAP (R$)","Meta_Faturamento","Forecast","Realizado"],
                            pct_cols=["Atingimento Atual (%)","Atingimento Forecast (%)"]
                        )

# ===================== VISÃO EXECUTIVA =====================
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

    if {"Ano-Mes","Valor Pedido R$","Lucro Bruto"}.issubset(flt.columns):
        serie = flt.groupby("Ano-Mes", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum"}).sort_values("Ano-Mes")
        mg = flt.groupby("Ano-Mes", as_index=False).apply(
            lambda d: pd.Series({"Margem %": (100*d["Lucro Bruto"].sum()/d["Valor Pedido R$"].sum()) if d["Valor Pedido R$"].sum()>0 else np.nan})
        ).reset_index(drop=True)
        serie = serie.merge(mg, on="Ano-Mes", how="left")
        if len(serie) > 12: serie = serie.tail(12)

        k1, k2, k3 = st.columns(3)
        with k1:
            st.caption("Faturamento – últimos 12 meses")
            st.altair_chart(alt.Chart(serie).mark_area(opacity=0.4).encode(
                x=alt.X("Ano-Mes:N", sort=None, title=None),
                y=alt.Y("Valor Pedido R$:Q", title=None),
                tooltip=[alt.Tooltip("Ano-Mes:N"), alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
            ), use_container_width=True)
        with k2:
            st.caption("Lucro Bruto – últimos 12 meses")
            st.altair_chart(alt.Chart(serie).mark_area(opacity=0.4).encode(
                x=alt.X("Ano-Mes:N", sort=None, title=None),
                y=alt.Y("Lucro Bruto:Q", title=None),
                tooltip=[alt.Tooltip("Ano-Mes:N"), alt.Tooltip("Lucro Bruto:Q", format=",.0f")]
            ), use_container_width=True)
        with k3:
            st.caption("Margem Bruta (%) – últimos 12 meses")
            st.altair_chart(alt.Chart(serie).mark_line(point=True).encode(
                x=alt.X("Ano-Mes:N", sort=None, title=None),
                y=alt.Y("Margem %:Q", title=None),
                tooltip=[alt.Tooltip("Ano-Mes:N"), alt.Tooltip("Margem %:Q", format=",.1f")]
            ), use_container_width=True)

    if "Lucro Bruto" in flt.columns and len(flt) > 0:
        pos = int((flt["Lucro Bruto"] > 0).sum())
        neg = int((flt["Lucro Bruto"] < 0).sum())
        donut_df = pd.DataFrame({"Categoria": ["Rentáveis","Negativos"], "Qtd": [pos, neg]})
        cdon1, cdon2 = st.columns([2,1])
        with cdon1:
            st.caption("Composição de linhas – rentáveis vs negativas")
            st.altair_chart(alt.Chart(donut_df).mark_arc(innerRadius=60).encode(
                theta="Qtd:Q", color="Categoria:N", tooltip=["Categoria","Qtd"]
            ).properties(height=300), use_container_width=True)
        with cdon2:
            tot = pos + neg
            st.metric("% Linhas Rentáveis", fmt_pct_safe(100*pos/tot) if tot>0 else "-")

# ===================== RFM =====================
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
        try: return pd.qcut(s.rank(method="first"), q=len(labels), labels=labels)
        except Exception: return pd.Series([labels[len(labels)//2]]*len(s), index=s.index)
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
    st.subheader("Clientes – RFM")
    ref_date = pd.to_datetime(d_fim) if d_fim is not None else None
    rfm = compute_rfm(flt, ref_date=ref_date)
    segs = sorted(rfm["Segmento"].unique())
    pick = st.multiselect("Segmentos", segs, default=segs)
    view = rfm[rfm["Segmento"].isin(pick)]
    c1, c2, c3 = st.columns(3)
    c1.metric("Clientes avaliados", fmt_int(len(view)))
    c2.metric("Mediana Recência (dias)", fmt_int(np.nanmedian(view["RecenciaDias"])) if len(view)>0 else "-")
    c3.metric("Mediana Valor (R$)", fmt_money(np.nanmedian(view["Valor"])) if len(view)>0 else "-")
    cols = ["Nome Cliente","RecenciaDias","Frequencia","Valor","R_Score","F_Score","M_Score","Score","Segmento"]
    display_table(view[cols], money_cols=["Valor"], int_cols=["RecenciaDias","Frequencia","Score"])
    try:
        st.altair_chart(alt.Chart(view.reset_index(drop=True)).mark_circle(size=70).encode(
            x=alt.X("Frequencia:Q", title="Frequência"),
            y=alt.Y("Valor:Q", title="Valor (R$)"),
            color=alt.Color("Segmento:N"),
            tooltip=["Nome Cliente","Frequencia", alt.Tooltip("Valor:Q", format=",.0f"), "RecenciaDias","Segmento"]
        ).properties(height=420), use_container_width=True)
    except Exception: pass

# ===================== RENTABILIDADE =====================
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
        st.markdown("#### Top 20 – Clientes por Lucro Bruto")
        top_cli = flt.groupby("Nome Cliente", as_index=False)["Lucro Bruto"].sum().sort_values("Lucro Bruto", ascending=False).head(20)
        display_table(top_cli, money_cols=["Lucro Bruto"])
    if {"ITEM","Lucro Bruto"}.issubset(flt.columns):
        st.markdown("#### Top 20 – SKUs por Lucro Bruto")
        top_sku = flt.groupby("ITEM", as_index=False)["Lucro Bruto"].sum().sort_values("Lucro Bruto", ascending=False).head(20)
        display_table(top_sku, money_cols=["Lucro Bruto"])
    if {"Nome Cliente","Valor Pedido R$","Lucro Bruto"}.issubset(flt.columns):
        st.markdown("#### Dispersão – Valor x Margem (%) por Cliente")
        disp = flt.groupby("Nome Cliente", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum"})
        disp["Margem %"] = np.where(disp["Valor Pedido R$"]>0, 100.0*disp["Lucro Bruto"]/disp["Valor Pedido R$"], np.nan)
        st.altair_chart(alt.Chart(disp).mark_circle(size=70).encode(
            x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
            y=alt.Y("Margem %:Q", title="Margem (%)"),
            tooltip=["Nome Cliente", alt.Tooltip("Valor Pedido R$:Q", format=",.0f"), alt.Tooltip("Margem %:Q", format=",.1f")]
        ).properties(height=420), use_container_width=True)
    if {"UF","Lucro Bruto","Valor Pedido R$"}.issubset(flt.columns):
        st.markdown("#### Margem por UF")
        por_uf = flt.groupby("UF", as_index=False).agg({"Lucro Bruto":"sum","Valor Pedido R$":"sum"})
        por_uf["Margem %"] = np.where(por_uf["Valor Pedido R$"]>0, 100.0*por_uf["Lucro Bruto"]/por_uf["Valor Pedido R$"], np.nan)
        display_table(por_uf.sort_values("Margem %", ascending=False), money_cols=["Lucro Bruto","Valor Pedido R$"], pct_cols=["Margem %"])
    if "Lucro Bruto" in flt.columns:
        st.markdown("#### Auditoria – Linhas com Margem Negativa")
        neg = flt[flt["Lucro Bruto"] < 0].copy()
        cols_show = [c for c in ["Transação","Nome Cliente","Pedido","ITEM","Representante","UF","Valor Pedido R$","Custo","Custo Total","Lucro Bruto","Margem %","Data do Pedido","Data / Mês"] if c in neg.columns]
        display_table(neg[cols_show], money_cols=["Valor Pedido R$","Custo","Custo Total","Lucro Bruto"], pct_cols=["Margem %"])

# ===================== CLIENTES (UPGRADE) =====================
with tab_cli:
    st.subheader("Clientes – Base ativa, expansão e risco")
    if "Nome Cliente" not in df.columns or date_col_main is None or "Valor Pedido R$" not in df.columns:
        st.info("Requer colunas 'Nome Cliente', data e 'Valor Pedido R$'.")
    else:
        d_ini_ts = pd.to_datetime(d_ini) if d_ini is not None else df[date_col_main].min()
        d_fim_ts = pd.to_datetime(d_fim) if d_fim is not None else df[date_col_main].max()
        df_cli_periodo = flt.dropna(subset=["Nome Cliente"]).copy()
        df_cli_periodo = df_cli_periodo[df_cli_periodo[date_col_main].between(d_ini_ts, d_fim_ts)]

        clientes_ativos = df_cli_periodo["Nome Cliente"].nunique()
        grp_dates = df.groupby("Nome Cliente")[date_col_main]
        first_buy = grp_dates.min()
        last_buy = grp_dates.max()
        clientes_periodo = set(df_cli_periodo["Nome Cliente"].unique())
        clientes_novos = [c for c in clientes_periodo if first_buy[c] >= d_ini_ts]
        janela_previa_ini = d_ini_ts - pd.DateOffset(months=12)
        prev_mask = (last_buy >= janela_previa_ini) & (last_buy < d_ini_ts)
        clientes_prev = set(last_buy[prev_mask].index)
        clientes_perdidos = sorted(clientes_prev - clientes_periodo)

        valor_total = df_cli_periodo["Valor Pedido R$"].sum()
        n_ped_cli = df_cli_periodo["Pedido"].nunique() if "Pedido" in df_cli_periodo.columns else len(df_cli_periodo)
        ticket_med_cli = valor_total / n_ped_cli if n_ped_cli else np.nan

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Clientes ativos", fmt_int(clientes_ativos))
        c2.metric("Clientes novos", fmt_int(len(clientes_novos)))
        c3.metric("Clientes perdidos (vs 12m)", fmt_int(len(clientes_perdidos)))
        c4.metric("Ticket médio", fmt_money(ticket_med_cli) if pd.notna(ticket_med_cli) else "-")

        st.markdown("### Ranking de clientes (faturamento do período)")
        rank_cli = df_cli_periodo.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        if not rank_cli.empty:
            rank_cli["% do total"] = 100 * rank_cli["Valor Pedido R$"] / rank_cli["Valor Pedido R$"].sum()
            rank_cli["% acumulado"] = rank_cli["% do total"].cumsum()
            display_table(rank_cli.head(100), money_cols=["Valor Pedido R$"], pct_cols=["% do total","% acumulado"])
            st.altair_chart(alt.Chart(rank_cli.head(40)).mark_bar().encode(
                x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
                y=alt.Y("Nome Cliente:N", sort="-x"),
                tooltip=["Nome Cliente", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
            ).properties(height=500), use_container_width=True)

        if {"Lucro Bruto","Valor Pedido R$"}.issubset(df_cli_periodo.columns):
            st.markdown("### Dispersão – Faturamento x Margem por cliente")
            cli_disp = df_cli_periodo.groupby("Nome Cliente", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum"})
            cli_disp["Margem %"] = np.where(cli_disp["Valor Pedido R$"]>0, 100*cli_disp["Lucro Bruto"]/cli_disp["Valor Pedido R$"], np.nan)
            st.altair_chart(alt.Chart(cli_disp).mark_circle(size=70).encode(
                x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
                y=alt.Y("Margem %:Q", title="Margem (%)"),
                tooltip=["Nome Cliente", alt.Tooltip("Valor Pedido R$:Q", format=",.0f"), alt.Tooltip("Margem %:Q", format=",.1f")]
            ).properties(height=420), use_container_width=True)

        with st.expander("Clientes novos e perdidos – listas"):
            col_n, col_p = st.columns(2)
            if clientes_novos:
                col_n.markdown("**Novos (primeira compra no período)**"); col_n.write(", ".join(sorted(clientes_novos[:80])) + (" ..." if len(clientes_novos) > 80 else ""))
            else:
                col_n.caption("Sem novos.")
            if clientes_perdidos:
                col_p.markdown("**Perdidos (compravam até 12m antes e não compraram no período)**"); col_p.write(", ".join(clientes_perdidos[:80]) + (" ..." if len(clientes_perdidos) > 80 else ""))
            else:
                col_p.caption("Sem perdidos.")

# ===================== PRODUTOS (UPGRADE) =====================
with tab_sku:
    st.subheader("Produtos – Mix, margem e curva ABC")
    if "ITEM" not in df.columns or "Valor Pedido R$" not in df.columns:
        st.info("Requer colunas 'ITEM' e 'Valor Pedido R$'.")
    else:
        df_prod_periodo = flt.copy()
        if qty_col is not None and qty_col in df_prod_periodo.columns:
            df_prod_periodo[qty_col] = df_prod_periodo[qty_col].apply(to_num)
        else:
            df_prod_periodo[qty_col] = np.nan
        grup = df_prod_periodo.groupby("ITEM", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum", qty_col:"sum"}).rename(columns={qty_col:"Quantidade"})
        grup["Margem %"] = np.where(grup["Valor Pedido R$"]>0, 100*grup["Lucro Bruto"]/grup["Valor Pedido R$"], np.nan)
        skus_ativos = grup["ITEM"].nunique()
        skus_totais = df["ITEM"].nunique()
        c1, c2, c3 = st.columns(3)
        c1.metric("SKUs ativos", fmt_int(skus_ativos))
        c2.metric("SKUs totais", fmt_int(skus_totais))
        c3.metric("Produtos com margem negativa", fmt_int((grup["Margem %"]<0).sum()))
        st.markdown("### Ranking de produtos (faturamento)")
        display_table(grup.sort_values("Valor Pedido R$", ascending=False).head(100),
                      money_cols=["Valor Pedido R$","Lucro Bruto"], pct_cols=["Margem %"], int_cols=["Quantidade"])
        if not grup.empty:
            abc = grup.sort_values("Valor Pedido R$", ascending=False).copy()
            abc["%Acum"] = 100 * abc["Valor Pedido R$"].cumsum() / abc["Valor Pedido R$"].sum()
            abc["Classe"] = abc["%Acum"].apply(lambda p: "A" if p<=80 else ("B" if p<=95 else "C"))
            st.markdown("### Curva ABC – faturamento por SKU")
            display_table(abc.head(150), money_cols=["Valor Pedido R$"], pct_cols=["%Acum"])
            st.altair_chart(alt.Chart(abc.head(80)).mark_bar().encode(
                x=alt.X("ITEM:N", sort=None, title="SKU"),
                y=alt.Y("Valor Pedido R$:Q", title="Faturamento (R$)"),
                color=alt.Color("Classe:N"),
                tooltip=["ITEM", alt.Tooltip("Valor Pedido R$:Q", format=",.0f"), "Classe","%Acum"]
            ).properties(height=420), use_container_width=True)
        st.markdown("### Dispersão – Volume x Margem (%)")
        st.altair_chart(alt.Chart(grup.dropna(subset=["Quantidade"])).mark_circle(size=70).encode(
            x=alt.X("Quantidade:Q", title="Quantidade vendida"),
            y=alt.Y("Margem %:Q", title="Margem (%)"),
            tooltip=["ITEM","Quantidade", alt.Tooltip("Valor Pedido R$:Q", format=",.0f"), alt.Tooltip("Margem %:Q", format=",.1f")]
        ).properties(height=420), use_container_width=True)

# ===================== REPRESENTANTES (UPGRADE) =====================
with tab_rep:
    st.subheader("Representantes – Comparativo de performance")
    if "Representante" not in df.columns or "Valor Pedido R$" not in df.columns:
        st.info("Requer colunas 'Representante' e 'Valor Pedido R$'.")
    else:
        df_rep_per = flt.copy()
        if "Nome Cliente" in df_rep_per.columns:
            fat_rep = df_rep_per.groupby("Representante", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum","Nome Cliente":"nunique"}).rename(columns={"Nome Cliente":"Clientes Ativos"})
        else:
            fat_rep = df_rep_per.groupby("Representante", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum"}); fat_rep["Clientes Ativos"]=np.nan
        if "Pedido" in df_rep_per.columns:
            ped_rep = df_rep_per.groupby("Representante", as_index=False)["Pedido"].nunique().rename(columns={"Pedido":"Pedidos"})
            fat_rep = fat_rep.merge(ped_rep, on="Representante", how="left")
        else:
            fat_rep["Pedidos"]=np.nan
        fat_rep["Ticket Médio"] = np.where(fat_rep["Pedidos"]>0, fat_rep["Valor Pedido R$"]/fat_rep["Pedidos"], np.nan)
        fat_rep["Margem %"] = np.where(fat_rep["Valor Pedido R$"]>0, 100*fat_rep["Lucro Bruto"]/fat_rep["Valor Pedido R$"], np.nan)
        reps_ativos = fat_rep["Representante"].nunique()
        c1, c2, c3 = st.columns(3)
        c1.metric("Representantes ativos", fmt_int(reps_ativos))
        c2.metric("Ticket médio global", fmt_money(fat_rep["Ticket Médio"].mean()) if len(fat_rep)>0 else "-")
        c3.metric("Clientes médios/rep", fmt_int(fat_rep["Clientes Ativos"].mean()) if "Clientes Ativos" in fat_rep.columns else "-")
        st.markdown("### Ranking de representantes")
        display_table(fat_rep.sort_values("Valor Pedido R$", ascending=False),
                      money_cols=["Valor Pedido R$","Lucro Bruto","Ticket Médio"], pct_cols=["Margem %"], int_cols=["Clientes Ativos","Pedidos"])
        top_n_rep = st.slider("Top N representantes no gráfico", 5, 40, 20, step=5)
        chart_rep = fat_rep.sort_values("Valor Pedido R$", ascending=False).head(top_n_rep)
        st.altair_chart(alt.Chart(chart_rep).mark_bar().encode(
            x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
            y=alt.Y("Representante:N", sort="-x"),
            tooltip=["Representante", alt.Tooltip("Valor Pedido R$:Q", format=",.0f"), alt.Tooltip("Margem %:Q", format=",.1f")]
        ).properties(height=480), use_container_width=True)
        st.markdown("### Dispersão – Produtividade x Margem")
        fat_rep["Produtividade (R$ / cliente)"] = np.where(fat_rep["Clientes Ativos"]>0, fat_rep["Valor Pedido R$"]/fat_rep["Clientes Ativos"], np.nan)
        st.altair_chart(alt.Chart(fat_rep).mark_circle(size=70).encode(
            x=alt.X("Produtividade (R$ / cliente):Q", title="R$ por cliente"),
            y=alt.Y("Margem %:Q", title="Margem (%)"),
            tooltip=["Representante", alt.Tooltip("Valor Pedido R$:Q", format=",.0f"), alt.Tooltip("Clientes Ativos:Q"), alt.Tooltip("Produtividade (R$ / cliente):Q", format=",.0f"), alt.Tooltip("Margem %:Q", format=",.1f")]
        ).properties(height=420), use_container_width=True)

# ===================== GEOGRAFIA (UPGRADE) =====================
with tab_geo:
    st.subheader("Geografia – Cobertura e performance por UF")
    if "UF" not in df.columns or "Valor Pedido R$" not in df.columns:
        st.info("Requer colunas 'UF' e 'Valor Pedido R$'.")
    else:
        geo = flt.copy()
        fat_uf = geo.groupby("UF", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum"})
        fat_uf["Pedidos"] = geo.groupby("UF")["Pedido"].nunique().values if "Pedido" in geo.columns else np.nan
        fat_uf["Clientes"] = geo.groupby("UF")["Nome Cliente"].nunique().values if "Nome Cliente" in geo.columns else np.nan
        fat_uf["Margem %"] = np.where(fat_uf["Valor Pedido R$"]>0, 100*fat_uf["Lucro Bruto"]/fat_uf["Valor Pedido R$"], np.nan)
        total_fat = fat_uf["Valor Pedido R$"].sum()
        fat_uf["% Part"] = np.where(total_fat>0, 100*fat_uf["Valor Pedido R$"]/total_fat, np.nan)
        c1, c2, c3 = st.columns(3)
        c1.metric("Faturamento (filtro)", fmt_money(total_fat))
        c2.metric("UFs ativas", fmt_int(fat_uf["UF"].nunique()))
        c3.metric("Maior participação", fmt_pct_safe(fat_uf["% Part"].max()) if not fat_uf.empty else "-")
        st.markdown("### Ranking de UFs")
        display_table(fat_uf.sort_values("Valor Pedido R$", ascending=False),
                      money_cols=["Valor Pedido R$","Lucro Bruto"], pct_cols=["Margem %","% Part"], int_cols=["Pedidos","Clientes"])
        st.markdown("### Faturamento por UF")
        st.altair_chart(alt.Chart(fat_uf.sort_values("Valor Pedido R$", ascending=False)).mark_bar().encode(
            x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
            y=alt.Y("UF:N", sort="-x"),
            tooltip=["UF", alt.Tooltip("Valor Pedido R$:Q", format=",.0f"), alt.Tooltip("% Part:Q", format=".1f")]
        ).properties(height=420), use_container_width=True)
        st.markdown("### Margem por UF")
        st.altair_chart(alt.Chart(fat_uf.sort_values("Margem %", ascending=False)).mark_bar().encode(
            x=alt.X("Margem %:Q", title="Margem (%)"),
            y=alt.Y("UF:N", sort="-x"),
            tooltip=["UF", alt.Tooltip("Margem %:Q", format=".1f")]
        ).properties(height=420), use_container_width=True)

# ===================== OPERACIONAL (UPGRADE) =====================
with tab_ops:
    st.subheader("Operacional – Lead time, atrasos e execução")
    df_ops = flt.copy()
    lt = df_ops["LeadTime (dias)"].dropna() if "LeadTime (dias)" in df_ops.columns else pd.Series([], dtype=float)
    atrasados = df_ops["AtrasadoFlag"].fillna(False) if "AtrasadoFlag" in df_ops.columns else pd.Series([False]*len(df_ops), index=df_ops.index)
    total_pedidos = df_ops["Pedido"].nunique() if "Pedido" in df_ops.columns else len(df_ops)
    pedidos_atrasados = df_ops.loc[atrasados, "Pedido"].nunique() if "Pedido" in df_ops.columns else int(atrasados.sum())
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pedidos (filtro)", fmt_int(total_pedidos))
    c2.metric("Pedidos com atraso", fmt_int(pedidos_atrasados))
    c3.metric("% de pedidos atrasados", fmt_pct_safe(100*pedidos_atrasados/total_pedidos) if total_pedidos else "-")
    c4.metric("Lead time médio (dias)", fmt_int(lt.mean()) if len(lt)>0 else "-")
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        if len(lt)>0:
            st.markdown("### Distribuição de Lead Time (dias)")
            hist_df = pd.DataFrame({"LeadTime": lt})
            st.altair_chart(alt.Chart(hist_df).mark_bar().encode(
                x=alt.X("LeadTime:Q", bin=alt.Bin(maxbins=20), title="Lead time (dias)"),
                y=alt.Y("count()", title="Qtde"),
                tooltip=[alt.Tooltip("count():Q", title="Qtde")]
            ).properties(height=320), use_container_width=True)
    with col_l2:
        if "Atrasado / No prazo" in df_ops.columns and "Pedido" in df_ops.columns:
            atrasos = df_ops.groupby("Atrasado / No prazo", as_index=False)["Pedido"].nunique().rename(columns={"Pedido":"Qtde Pedidos"})
            display_table(atrasos, int_cols=["Qtde Pedidos"])
            st.altair_chart(alt.Chart(atrasos).mark_bar().encode(
                x=alt.X("Qtde Pedidos:Q", title="Pedidos"),
                y=alt.Y("Atrasado / No prazo:N", sort="-x", title="Status"),
                tooltip=["Atrasado / No prazo","Qtde Pedidos"]
            ).properties(height=320), use_container_width=True)
    st.markdown("### Pedidos em aberto e em produção")
    if "Status de Produção / Faturamento" in df_ops.columns:
        status_series = df_ops["Status de Produção / Faturamento"].astype(str)
        is_aberto = status_series.str.contains("abert|pend|prod", case=False, na=False) & ~status_series.str.contains("fatur", case=False, na=False)
        df_abertos = df_ops[is_aberto].copy()
        if not df_abertos.empty:
            if "Pedido" in df_abertos.columns:
                resumo_abertos = df_abertos.groupby("Pedido", as_index=False).agg({"Valor Pedido R$":"sum","Nome Cliente":"first",date_col_main:"max"}).rename(columns={"Valor Pedido R$":"Valor Pedido","Nome Cliente":"Cliente",date_col_main:"Data Última Atualização"})
                display_table(resumo_abertos.sort_values("Valor Pedido", ascending=False), money_cols=["Valor Pedido"])
            else:
                display_table(df_abertos, money_cols=["Valor Pedido R$"])
        else:
            st.caption("Nenhum pedido em aberto / produção no filtro atual.")

# ===================== PARETO / ABC =====================
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

# ===================== SEBASTIAN (consolidado) =====================
with tab_seb:
    st.subheader("SEBASTIAN – Visão integrada tática")
    # 1) KPIs por vendedor / gerente
    if {"Representante","Valor Pedido R$"}.issubset(flt.columns):
        kpis = flt.groupby("Representante", as_index=False).agg({"Valor Pedido R$":"sum","Pedido":"nunique","Nome Cliente":"nunique"})
        kpis = kpis.rename(columns={"Pedido":"Pedidos","Nome Cliente":"Clientes"})
        kpis["Ticket Médio"] = np.where(kpis["Pedidos"]>0, kpis["Valor Pedido R$"]/kpis["Pedidos"], np.nan)
        display_table(kpis.sort_values("Valor Pedido R$", ascending=False),
                      money_cols=["Valor Pedido R$","Ticket Médio"], int_cols=["Pedidos","Clientes"])
    # 2) Histórico 12 meses
    if {"Ano-Mes","Valor Pedido R$"}.issubset(flt.columns):
        hist = flt.groupby("Ano-Mes", as_index=False)["Valor Pedido R$"].sum().sort_values("Ano-Mes")
        st.altair_chart(alt.Chart(hist).mark_line(point=True).encode(
            x=alt.X("Ano-Mes:N", title=None), y=alt.Y("Valor Pedido R$:Q", title="R$"),
            tooltip=["Ano-Mes", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
        ).properties(height=320), use_container_width=True)
    # 3) Família (Observação) – últimos 12m
    if {"Observação","Ano-Mes","Valor Pedido R$"}.issubset(flt.columns):
        fam = flt.groupby(["Ano-Mes","Observação"], as_index=False)["Valor Pedido R$"].sum()
        fam_last = fam[fam["Ano-Mes"].isin(sorted(fam["Ano-Mes"].unique())[-12:])]
        st.altair_chart(alt.Chart(fam_last).mark_area(opacity=0.5).encode(
            x=alt.X("Ano-Mes:N", title=None), y=alt.Y("Valor Pedido R$:Q", title="R$"),
            color="Observação:N",
            tooltip=["Ano-Mes","Observação", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
        ).properties(height=360), use_container_width=True)
    # 4) Pedidos em aberto
    if "Status de Produção / Faturamento" in flt.columns:
        status_series = flt["Status de Produção / Faturamento"].astype(str)
        is_aberto = status_series.str.contains("abert|pend|prod", case=False, na=False) & ~status_series.str.contains("fatur", case=False, na=False)
        df_ab = flt[is_aberto].copy()
        if not df_ab.empty:
            cols = [c for c in ["Transação","Pedido","Nome Cliente","Representante","ITEM","Valor Pedido R$","Data do Pedido","Status de Produção / Faturamento"] if c in df_ab.columns]
            st.markdown("### Pedidos em aberto")
            display_table(df_ab[cols].sort_values("Valor Pedido R$", ascending=False), money_cols=["Valor Pedido R$"])

# ===================== SIMULADOR DE VENDAS =====================
with tab_sim:
    st.subheader("Simulador de Vendas – multi-SKU com MC e impostos")
    if {"ITEM","Valor Pedido R$"}.issubset(df.columns):
        # Seleção de SKUs
        skus = sorted(df["ITEM"].astype(str).dropna().unique().tolist())
        pick_skus = st.multiselect("Selecione os SKUs para simular", skus[:1])

        # históricos
        qty_candidates = [qty_col] if qty_col else []
        qty_used = qty_candidates[0] if qty_candidates else None

        if pick_skus:
            base_hist = df[df["ITEM"].astype(str).isin(pick_skus)].copy()
            if qty_used is not None and qty_used in base_hist.columns:
                base_hist["Qtd"] = base_hist[qty_used].apply(to_num)
            else:
                base_hist["Qtd"] = np.nan
            base_hist["PreçoUnit"] = np.where(base_hist["Qtd"]>0, base_hist["Valor Pedido R$"]/base_hist["Qtd"], np.nan)
            base_hist["CustoUnit"] = np.where(base_hist["Qtd"]>0, base_hist["Custo Total"]/base_hist["Qtd"], np.nan)
            hist = base_hist.groupby("ITEM", as_index=False).agg({"Qtd":"sum","Valor Pedido R$":"sum","Custo Total":"sum","PreçoUnit":"mean","CustoUnit":"mean"})
            hist = hist.rename(columns={"ITEM":"SKU","Valor Pedido R$":"FatHist","Custo Total":"CustoHist","Qtd":"QtdHist","PreçoUnit":"PrecoMed","CustoUnit":"CustoMed"})
            st.markdown("### Histórico por SKU")
            display_table(hist, money_cols=["FatHist","CustoHist","PrecoMed","CustoMed"], int_cols=["QtdHist"])

            st.markdown("### Parâmetros globais")
            col_tax1, col_tax2, col_mc = st.columns([1,1,1])
            with col_tax1:
                icms = st.number_input("ICMS (%)", 0.0, 100.0, 18.0, 0.1)
                pis = st.number_input("PIS (%)", 0.0, 100.0, 1.65, 0.05)
                cof = st.number_input("COFINS (%)", 0.0, 100.0, 7.6, 0.1)
            with col_tax2:
                outros = st.number_input("Outros impostos (%)", 0.0, 100.0, 0.0, 0.1)
                frete = st.number_input("Frete (%)", 0.0, 100.0, 2.0, 0.1)
                comis = st.number_input("Comissão (%)", 0.0, 100.0, 3.0, 0.1)
            with col_mc:
                mc_alvo = st.number_input("MC mínima desejada (%)", 0.0, 99.9, 20.0, 0.1)

            rows = []
            for _, r in hist.iterrows():
                with st.expander(f"Configurar {r['SKU']}"):
                    qtd_sim = st.number_input(f"Quantidade simulada – {r['SKU']}", 0.0, 1e9, float(r["QtdHist"] if pd.notna(r["QtdHist"]) and r["QtdHist"]>0 else 100.0), 1.0, key=f"q_{r['SKU']}")
                    adj_preco = st.number_input(f"Ajuste % preço vs histórico – {r['SKU']}", -50.0, 100.0, 0.0, 0.5, key=f"ap_{r['SKU']}")
                    adj_custo = st.number_input(f"Ajuste % custo vs histórico – {r['SKU']}", -50.0, 100.0, 0.0, 0.5, key=f"ac_{r['SKU']}")
                    preco_manual = st.number_input(f"Preço unitário manual – {r['SKU']} (0 = usar histórico ajustado)", 0.0, 1e9, 0.0, 0.01, key=f"pm_{r['SKU']}")
                    preco_base = float(r["PrecoMed"]) if pd.notna(r["PrecoMed"]) and r["PrecoMed"]>0 else 0.0
                    custo_base = float(r["CustoMed"]) if pd.notna(r["CustoMed"]) and r["CustoMed"]>0 else 0.0
                    preco_sim = preco_manual if preco_manual>0 else preco_base * (1 + adj_preco/100.0)
                    custo_sim = custo_base * (1 + adj_custo/100.0)
                    fat_sim = preco_sim * qtd_sim
                    custo_tot = custo_sim * qtd_sim
                    lucro = fat_sim - custo_tot
                    marg = 100*lucro/fat_sim if fat_sim>0 else np.nan
                    rows.append({"SKU": r["SKU"], "Qtd Simulada": qtd_sim, "Preço Unitário Simulado": preco_sim, "Custo Unitário Simulado": custo_sim, "Faturamento Simulado": fat_sim, "Lucro Bruto Simulado": lucro, "Margem Bruta %": marg})
            sim_tbl = pd.DataFrame(rows)
            st.markdown("### Consolidação")
            display_table(sim_tbl, money_cols=["Preço Unitário Simulado","Custo Unitário Simulado","Faturamento Simulado","Lucro Bruto Simulado"], pct_cols=["Margem Bruta %"], int_cols=["Qtd Simulada"])

            T_imp = (icms + pis + cof + outros)/100.0
            T_d = (frete + comis)/100.0
            M = mc_alvo/100.0
            A = (1 - M) * (1 - T_imp) - T_d
            if A <= 0:
                st.error("Parâmetros globais inviáveis para calcular preço mínimo (A <= 0).")
            else:
                sim_tbl["Preço Unit. Mínimo (MC alvo)"] = sim_tbl["Custo Unitário Simulado"] / A
                sim_tbl["Desc Máx vs Preço Sim (%)"] = 100*(1 - sim_tbl["Preço Unit. Mínimo (MC alvo)"] / sim_tbl["Preço Unitário Simulado"])
                st.markdown("### Engenharia de preço mínimo por SKU")
                display_table(sim_tbl[["SKU","Preço Unitário Simulado","Custo Unitário Simulado","Preço Unit. Mínimo (MC alvo)","Desc Máx vs Preço Sim (%)","Margem Bruta %","Qtd Simulada","Faturamento Simulado","Lucro Bruto Simulado"]],
                              money_cols=["Preço Unitário Simulado","Custo Unitário Simulado","Preço Unit. Mínimo (MC alvo)","Faturamento Simulado","Lucro Bruto Simulado"],
                              pct_cols=["Desc Máx vs Preço Sim (%)","Margem Bruta %"], int_cols=["Qtd Simulada"])

            st.markdown("### Exportar simulação")
            fat_total = sim_tbl["Faturamento Simulado"].sum()
            imp_total = fat_total * T_imp
            desp_var_total = fat_total * T_d
            custo_total = (sim_tbl["Custo Unitário Simulado"] * sim_tbl["Qtd Simulada"]).sum()
            receita_liq = fat_total - imp_total
            mc_val = receita_liq - custo_total - desp_var_total
            mc_pct = 100*mc_val/receita_liq if receita_liq>0 else np.nan

            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                st.download_button("Baixar simulação (CSV)", data=sim_tbl.to_csv(index=False, encoding="utf-8-sig"), file_name="simulacao_venda_brasforma.csv", mime="text/csv")
            with col_exp2:
                try:
                    pdf_bytes = build_simulation_pdf(sim_tbl, fat_total, mc_val, mc_pct, icms, pis, cof, outros, frete, comis, mc_alvo)
                    st.download_button("Baixar simulação (PDF)", data=pdf_bytes, file_name="simulacao_venda_brasforma.pdf", mime="application/pdf")
                except Exception as e:
                    st.exception(e)
    else:
        st.info("Base precisa conter 'ITEM' e 'Valor Pedido R$'.")

# ===================== EXPORTAR =====================
with tab_export:
    st.subheader("Exportar CSV (filtro atual)")
    st.dataframe(flt, use_container_width=True)
    st.download_button("Baixar CSV filtrado", data=flt.to_csv(index=False, encoding="utf-8-sig"), file_name="brasforma_filtro.csv", mime="text/csv")

# ===================== FOOTER =====================
if qty_col:
    st.caption(f"✓ Custo calculado como **unitário × quantidade**. Coluna de quantidade detectada: **{qty_col}**.")
else:
    st.caption("! Atenção: coluna de quantidade não identificada — usando 'Custo' como total.")
