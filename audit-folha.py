
# =============================================================================
# Auditoria de Folha - App Streamlit
# Igarape Digital
# Padrao: Anderson Marinho
# =============================================================================

import io
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Auditoria de Folha - Igarape Digital",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# Paleta Igarape Digital
# -----------------------------------------------------------------------------
IG_AZUL       = "#0083CA"
IG_AZUL_ESC   = "#003C64"
IG_VERDE_ESC  = "#005A64"
IG_AZUL_CLR   = "#6EB4DC"
IG_VINHO      = "#7D0041"
IG_LARANJA    = "#8C321E"
IG_VERMELHO   = "#7D0041"
TEXTO_PRINCIPAL   = "#333333"
TEXTO_SECUNDARIO  = "#646464"
GRID_CINZA        = "#CCCCCC"
FUNDO             = "#FFFFFF"
COR_DESLIGADO     = "#646464"

STATUS_CRITICOS = {"AUSENTE", "NEGATIVO", "ZERO_SUSPEITO", "EXTREMA", "ALTA", "Z_2SIGMA"}

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500&family=Rufina:wght@400;700&display=swap');
    html, body, [class*="css"]  {{
        font-family: 'Roboto', sans-serif;
        font-weight: 300;
        color: {TEXTO_PRINCIPAL};
    }}
    h1, h2, h3, h4 {{
        font-family: 'Rufina', serif;
        color: {IG_AZUL_ESC};
    }}
    .brand-header {{
        background-color: {IG_AZUL};
        color: white;
        padding: 18px 24px;
        border-radius: 6px;
        margin-bottom: 18px;
    }}
    .brand-header h1 {{
        color: white;
        margin: 0;
        font-size: 22px;
    }}
    .brand-header p {{
        color: white;
        margin: 4px 0 0 0;
        font-size: 13px;
        font-weight: 300;
        opacity: 0.92;
    }}
    .metric-card {{
        background-color: white;
        border: 1px solid {GRID_CINZA};
        border-left: 4px solid {IG_AZUL};
        padding: 14px 18px;
        border-radius: 4px;
        height: 100%;
    }}
    .metric-card .titulo {{
        font-size: 11px;
        color: {TEXTO_SECUNDARIO};
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }}
    .metric-card .valor {{
        font-size: 22px;
        font-weight: 500;
        color: {IG_AZUL_ESC};
    }}
    .metric-card .delta {{
        font-size: 12px;
        margin-top: 4px;
    }}
    .metric-card .delta.positivo {{ color: {IG_VERDE_ESC}; }}
    .metric-card .delta.negativo {{ color: {IG_VERMELHO}; }}
    .metric-card.alerta {{ border-left-color: {IG_VERMELHO}; }}
    .metric-card.atencao {{ border-left-color: {IG_VINHO}; }}
    .app-footer {{
        text-align: center;
        font-size: 11px;
        color: {TEXTO_SECUNDARIO};
        padding: 14px 0;
        margin-top: 24px;
        border-top: 1px solid {GRID_CINZA};
    }}
    .stDataFrame {{
        border: 1px solid {GRID_CINZA};
    }}
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="brand-header">
    <h1>Auditoria de Folha - Confronto Folha Calculada vs Implantadas</h1>
    <p>Igarape Digital | Analise de Payroll | Padrao DP</p>
</div>
""", unsafe_allow_html=True)


def fmt_brl(x):
    if pd.isna(x):
        return "-"
    s = f"{abs(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-R$ {s}" if x < 0 else f"R$ {s}"


def fmt_pct(x):
    if pd.isna(x):
        return "-"
    return f"{x:+.1f}%"


@st.cache_data(show_spinner=False)
def carregar_csv(conteudo_bytes):
    df = pd.read_csv(io.BytesIO(conteudo_bytes), sep=";", encoding="latin-1", dtype=str)
    df = df.drop(columns=["Unnamed: 49"], errors="ignore")
    df = df[df["Empresa"] != "Empresa"].copy()

    cols_valor = [c for c in df.columns if " - Valor" in c]
    meses_detectados = [c.replace(" - Valor", "") for c in cols_valor]

    for col in cols_valor:
        df[col] = (df[col].astype(str)
                   .str.replace(".", "", regex=False)
                   .str.replace(",", ".", regex=False))
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df, meses_detectados


_MAP_MES = {"JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
            "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12}


def ultimo_dia_mes(rotulo_mes):
    try:
        mmm, aa = rotulo_mes.split("/")
        mes = _MAP_MES[mmm.upper()]
        ano = 2000 + int(aa) if int(aa) < 80 else 1900 + int(aa)
        if mes == 12:
            return pd.Timestamp(year=ano + 1, month=1, day=1) - pd.Timedelta(days=1)
        return pd.Timestamp(year=ano, month=mes + 1, day=1) - pd.Timedelta(days=1)
    except Exception:
        return None


def detectar_desligados(df, mes_alvo):
    sub = df[["Matrícula", "Nome", "Data Rescisão"]].drop_duplicates(["Matrícula", "Nome"])
    sub = sub[sub["Data Rescisão"].notna() & (sub["Data Rescisão"].astype(str).str.strip() != "")]
    sub = sub.copy()
    sub["DATA_RESC"] = pd.to_datetime(sub["Data Rescisão"], format="%d/%m/%Y", errors="coerce")
    sub = sub[sub["DATA_RESC"].notna()]
    limite = ultimo_dia_mes(mes_alvo)
    if limite is not None:
        sub = sub[sub["DATA_RESC"] <= limite]
    return {(row["Matrícula"], row["Nome"]): row["DATA_RESC"] for _, row in sub.iterrows()}


def resumo_macro(df, meses):
    pgto = df[df["Clas."] == "PGTO"]
    desc = df[df["Clas."] == "DESC"]
    v9950 = df[df["Código"] == "9950"]
    linhas = []
    for m in meses:
        col = f"{m} - Valor"
        total_pgto = pgto[col].sum()
        total_desc = desc[col].sum()
        total_9950 = v9950[col].sum()
        linhas.append({
            "Mes": m,
            "PGTO": total_pgto,
            "DESC": total_desc,
            "Liquido": total_pgto - total_desc,
            "Verba9950": total_9950,
            "FuncPGTO": pgto[pgto[col] != 0]["Matrícula"].nunique(),
        })
    return pd.DataFrame(linhas)


def salario_verba20_por_mes(df, meses, codigo_verba="0020"):
    """HC e salario medio da verba 0020 (Armazena Salario - classe OUTRO) por mes."""
    alvo = str(codigo_verba).strip().lstrip("0") or "0"
    codigos_norm = df["Código"].astype(str).str.strip().str.lstrip("0").replace("", "0")
    v20 = df[codigos_norm == alvo]
    linhas = []
    for m in meses:
        col = f"{m} - Valor"
        ativos = v20[v20[col] > 0]
        hc = ativos["Matrícula"].nunique()
        total = ativos[col].sum()
        media = (total / hc) if hc else 0
        linhas.append({
            "Mes": m,
            "HC": hc,
            "Salario_Total": total,
            "Salario_Medio": media,
        })
    return pd.DataFrame(linhas)


def liquido_por_funcionario(df, meses, mes_alvo, meses_baseline, desligados=None,
                            lim_extrema_pct=98, lim_extrema_abs=500,
                            lim_alta_pct=30, lim_alta_abs=200):
    desligados = desligados or {}
    val_cols = [f"{m} - Valor" for m in meses]
    pgto_func = df[df["Clas."] == "PGTO"].groupby(["Matrícula", "Nome"])[val_cols].sum()
    desc_func = df[df["Clas."] == "DESC"].groupby(["Matrícula", "Nome"])[val_cols].sum()
    liquido = pgto_func.subtract(desc_func, fill_value=0)
    liquido.columns = [c.replace(" - Valor", "") for c in liquido.columns]

    base = liquido[meses_baseline].replace(0, np.nan)
    liquido["MEDIA_BASELINE"] = base.mean(axis=1)
    liquido["DESVIO_BASELINE"] = base.std(axis=1)
    liquido["LIQUIDO_ALVO"] = liquido[mes_alvo]
    liquido["VAR_ABS"] = liquido["LIQUIDO_ALVO"] - liquido["MEDIA_BASELINE"]
    liquido["VAR_PCT"] = ((liquido["LIQUIDO_ALVO"] / liquido["MEDIA_BASELINE"]) - 1) * 100
    liquido["Z_SCORE"] = liquido["VAR_ABS"] / liquido["DESVIO_BASELINE"]
    liquido["DATA_RESCISAO"] = [desligados.get((mat, nome), pd.NaT) for (mat, nome) in liquido.index]

    def status(r):
        if pd.notna(r["DATA_RESCISAO"]):
            return "DESLIGADO"
        if pd.isna(r["MEDIA_BASELINE"]) or r["MEDIA_BASELINE"] == 0:
            return "NOVO_FUNC" if r["LIQUIDO_ALVO"] != 0 else "SEM_DADOS"
        if r["LIQUIDO_ALVO"] == 0 and r["MEDIA_BASELINE"] > 0:
            return "AUSENTE"
        if r["LIQUIDO_ALVO"] < -100:
            return "NEGATIVO"
        if r["LIQUIDO_ALVO"] < 500 and r["MEDIA_BASELINE"] >= 1000:
            return "ZERO_SUSPEITO"
        if abs(r["VAR_PCT"]) >= lim_extrema_pct and abs(r["VAR_ABS"]) > lim_extrema_abs:
            return "EXTREMA"
        if abs(r["VAR_PCT"]) >= lim_alta_pct and abs(r["VAR_ABS"]) > lim_alta_abs:
            return "ALTA"
        if pd.notna(r["Z_SCORE"]) and abs(r["Z_SCORE"]) >= 2:
            return "Z_2SIGMA"
        return "OK"

    liquido["STATUS"] = liquido.apply(status, axis=1)

    def auditoria(r):
        s = r["STATUS"]
        if s == "DESLIGADO":
            data_rsc = r["DATA_RESCISAO"].strftime("%d/%m/%Y") if pd.notna(r["DATA_RESCISAO"]) else "-"
            return f"Desligado em {data_rsc}. Liquido {fmt_brl(r['LIQUIDO_ALVO'])} em {mes_alvo}."
        if s == "AUSENTE":
            return f"Sem movimento; baseline {fmt_brl(r['MEDIA_BASELINE'])}. Verificar afastamento ou processamento."
        if s == "NEGATIVO":
            return f"Liquido negativo: {fmt_brl(r['LIQUIDO_ALVO'])}."
        if s == "ZERO_SUSPEITO":
            return f"Liquido proximo de zero ({fmt_brl(r['LIQUIDO_ALVO'])}) com baseline {fmt_brl(r['MEDIA_BASELINE'])}."
        if s == "EXTREMA":
            return f"Variacao extrema: {fmt_pct(r['VAR_PCT'])} ({fmt_brl(r['VAR_ABS'])})."
        if s == "ALTA":
            return f"Variacao alta: {fmt_pct(r['VAR_PCT'])} ({fmt_brl(r['VAR_ABS'])})."
        if s == "Z_2SIGMA":
            return f"Fora de 2 sigmas (z={r['Z_SCORE']:+.2f})."
        if s == "NOVO_FUNC":
            return f"Sem baseline; liquido {fmt_brl(r['LIQUIDO_ALVO'])}."
        if s == "SEM_DADOS":
            return "Sem movimento no periodo."
        return "Sem desvio relevante."

    liquido["AUDITORIA"] = liquido.apply(auditoria, axis=1)
    return liquido


def verbas_zeradas(df, classificacao, mes_alvo, meses_baseline, n_meses_min=3, valor_min=500):
    sub = df[df["Clas."] == classificacao].copy()
    val_cols_base = [f"{m} - Valor" for m in meses_baseline]
    col_alvo = f"{mes_alvo} - Valor"
    agg = sub.groupby(["Código", "Descrição"])[val_cols_base + [col_alvo]].sum()
    agg["N_MESES_ATIVOS"] = (agg[val_cols_base] >= valor_min).sum(axis=1)
    agg["MEDIA_BASELINE"] = agg[val_cols_base].replace(0, np.nan).mean(axis=1)
    agg["VALOR_ALVO"] = agg[col_alvo]
    zerados = agg[(agg["N_MESES_ATIVOS"] >= n_meses_min) & (agg["VALOR_ALVO"] == 0)]
    return zerados[["N_MESES_ATIVOS", "MEDIA_BASELINE", "VALOR_ALVO"]].sort_values("MEDIA_BASELINE", ascending=False)


def plotly_layout_brand(titulo, height=420):
    return dict(
        title=dict(text=titulo, font=dict(family="Rufina, serif", color=IG_AZUL_ESC, size=15)),
        plot_bgcolor=FUNDO,
        paper_bgcolor=FUNDO,
        font=dict(family="Roboto, sans-serif", color=TEXTO_PRINCIPAL, size=11),
        xaxis=dict(gridcolor=GRID_CINZA, linecolor=GRID_CINZA, tickcolor=GRID_CINZA),
        yaxis=dict(gridcolor=GRID_CINZA, linecolor=GRID_CINZA, tickcolor=GRID_CINZA),
        height=height,
        hoverlabel=dict(bgcolor="white", bordercolor=IG_AZUL, font=dict(family="Roboto", color=TEXTO_PRINCIPAL)),
        legend=dict(bgcolor="rgba(255,255,255,0.85)", bordercolor=GRID_CINZA, borderwidth=1, font=dict(size=10)),
        margin=dict(l=60, r=30, t=60, b=50),
    )


def cor_status(status):
    if status in STATUS_CRITICOS:
        return IG_VERMELHO
    mapa = {
        "OK": IG_AZUL_CLR,
        "NOVO_FUNC": IG_AZUL_ESC,
        "SEM_DADOS": GRID_CINZA,
        "DESLIGADO": COR_DESLIGADO,
    }
    return mapa.get(status, GRID_CINZA)


def limpar_filtros_tabela():
    st.session_state["filtro_status_tabela"] = []
    st.session_state["ordenar_por_tabela"] = "VAR_ABS (impacto)"
    st.session_state["buscar_nome"] = ""


def limpar_drilldown():
    st.session_state["func_escolhido"] = ""
    st.session_state["dd_status"] = []
    st.session_state["dd_cr"] = []
    st.session_state["dd_classe"] = []
    st.session_state["dd_processo"] = []
    st.session_state["dd_busca_verba"] = ""


with st.sidebar:
    st.markdown(f"<h3 style='color:{IG_AZUL_ESC};margin-top:0'>Configuracao</h3>", unsafe_allow_html=True)
    arquivo = st.file_uploader(
        "Arquivo CSV do relatorio (Confere - Codigos por Periodo)",
        type=["csv"],
        help="Exporte do ADP em formato CSV pt-BR (separador ; / decimal virgula)."
    )
    st.markdown("---")
    st.markdown(f"<small style='color:{TEXTO_SECUNDARIO}'>Apos carregar o arquivo, ajuste os parametros de auditoria abaixo.</small>", unsafe_allow_html=True)

if not arquivo:
    st.info("Carregue o arquivo CSV no painel lateral para iniciar a auditoria.")
    st.markdown(f"""
    <div style='margin-top:20px;padding:18px;background:#f7f9fb;border-left:4px solid {IG_AZUL};border-radius:4px'>
    <strong style='color:{IG_AZUL_ESC}'>Como funciona</strong><br>
    <small style='color:{TEXTO_PRINCIPAL}'>
    1. Carregue o relatorio "Confere Valores Codigo de Folha por Periodo" do ADP em CSV.<br>
    2. Selecione o mes-alvo e os meses de baseline.<br>
    3. O app calcula o liquido por funcionario, aplica as regras de auditoria e gera os graficos.<br>
    4. Voce pode filtrar a tabela, comparar meses e fazer drill-down por funcionario e verba.<br>
    5. Exporte o resultado em Excel pronto para analise.
    </small>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

try:
    df, meses_detectados = carregar_csv(arquivo.getvalue())
except Exception as e:
    st.error(f"Falha ao ler o arquivo. Confirme se e o CSV no padrao esperado.\n\n{e}")
    st.stop()

with st.sidebar:
    empresas_disponiveis = sorted([e for e in df["Empresa"].dropna().astype(str).unique() if e.strip()])
    if len(empresas_disponiveis) > 1:
        empresas_selecionadas = st.multiselect(
            "Empresa(s)",
            options=empresas_disponiveis,
            default=empresas_disponiveis,
            help="Filtre por uma ou mais empresas. Se nenhuma for marcada, todas serao consideradas."
        )
        if not empresas_selecionadas:
            empresas_selecionadas = empresas_disponiveis
    else:
        empresas_selecionadas = empresas_disponiveis
        st.caption(f"Empresa unica: {empresas_disponiveis[0] if empresas_disponiveis else '-'}")

    df = df[df["Empresa"].astype(str).isin(empresas_selecionadas)].copy()
    if df.empty:
        st.error("Nenhuma linha apos aplicar o filtro de empresa.")
        st.stop()

    processos_disponiveis = sorted([p for p in df["Processo"].dropna().astype(str).unique() if p.strip()])
    if len(processos_disponiveis) > 1:
        processos_selecionados = st.multiselect(
            "Processo (tipo de folha)",
            options=processos_disponiveis,
            default=processos_disponiveis,
            help="Filtre por tipo de folha (Mensal, 13o, Ferias, Adiantamento, Rescisao, etc.). Util quando ha mais de uma folha no mesmo mes."
        )
        if not processos_selecionados:
            processos_selecionados = processos_disponiveis
    else:
        processos_selecionados = processos_disponiveis
        if processos_disponiveis:
            st.caption(f"Processo unico: {processos_disponiveis[0]}")

    df = df[df["Processo"].astype(str).isin(processos_selecionados)].copy()
    if df.empty:
        st.error("Nenhuma linha apos aplicar o filtro de processo.")
        st.stop()

    mes_alvo = st.selectbox("Mes-alvo (folha calculada)", options=meses_detectados, index=len(meses_detectados)-1)
    meses_disponiveis_baseline = [m for m in meses_detectados if m != mes_alvo]
    default_baseline = meses_disponiveis_baseline[-5:] if len(meses_disponiveis_baseline) >= 5 else meses_disponiveis_baseline
    meses_baseline = st.multiselect("Meses para baseline (folhas implantadas)", options=meses_disponiveis_baseline, default=default_baseline)
    if not meses_baseline:
        st.error("Selecione pelo menos um mes para a baseline.")
        st.stop()
    excluir_desligados = st.toggle("Excluir desligados da analise", value=True)

    with st.expander("Parametros de auditoria (limiares de outliers)", expanded=False):
        st.caption("Ajuste a sensibilidade das classificacoes EXTREMA e ALTA. Padrao DP: 98% / 30%.")
        lim_extrema_pct = st.slider("Limiar EXTREMA - variacao % minima", min_value=50, max_value=200, value=98, step=1, help="Variacao percentual absoluta minima para classificar como EXTREMA.")
        lim_extrema_abs = st.slider("Limiar EXTREMA - impacto R$ minimo", min_value=100, max_value=5000, value=500, step=50, help="Impacto absoluto em R$ minimo para classificar como EXTREMA.")
        lim_alta_pct = st.slider("Limiar ALTA - variacao % minima", min_value=10, max_value=80, value=30, step=1, help="Variacao percentual absoluta minima para classificar como ALTA.")
        lim_alta_abs = st.slider("Limiar ALTA - impacto R$ minimo", min_value=50, max_value=2000, value=200, step=50, help="Impacto absoluto em R$ minimo para classificar como ALTA.")
        if lim_alta_pct >= lim_extrema_pct:
            st.warning("O limiar ALTA esta maior ou igual ao EXTREMA - revise os valores.")

    st.markdown("---")
    st.caption(f"{len(df):,} linhas | {df['Matrícula'].nunique()} funcionarios | {len(empresas_selecionadas)} empresa(s) | {len(processos_selecionados)} processo(s)")
    st.caption(f"Periodo: {meses_detectados[0]} a {meses_detectados[-1]}")

desligados = detectar_desligados(df, mes_alvo)
resumo = resumo_macro(df, meses_detectados)
liquido = liquido_por_funcionario(
    df, meses_detectados, mes_alvo, meses_baseline, desligados,
    lim_extrema_pct=lim_extrema_pct, lim_extrema_abs=lim_extrema_abs,
    lim_alta_pct=lim_alta_pct, lim_alta_abs=lim_alta_abs,
)
salario_v20 = salario_verba20_por_mes(df, meses_detectados)
zerados_pgto = verbas_zeradas(df, "PGTO", mes_alvo, meses_baseline)
zerados_desc = verbas_zeradas(df, "DESC", mes_alvo, meses_baseline)
liquido_analise = liquido[liquido["STATUS"] != "DESLIGADO"] if excluir_desligados else liquido

linha_alvo = resumo[resumo["Mes"] == mes_alvo].iloc[0]
linhas_baseline = resumo[resumo["Mes"].isin(meses_baseline)]
media_pgto_base = linhas_baseline["PGTO"].mean()
media_desc_base = linhas_baseline["DESC"].mean()
media_liq_base = linhas_baseline["Liquido"].mean()

st.markdown("### Visao geral")

n_desligados = (liquido["STATUS"] == "DESLIGADO").sum()
if n_desligados > 0:
    estado_filtro = "excluidos da analise" if excluir_desligados else "incluidos na analise"
    st.caption(f"Desligados detectados: {n_desligados} funcionarios | atualmente {estado_filtro}.")

c1, c2, c3, c4 = st.columns(4)
def card(coluna, titulo, valor, delta_txt=None, delta_classe=None, classe_card=""):
    delta_html = f"<div class='delta {delta_classe}'>{delta_txt}</div>" if delta_txt is not None else ""
    coluna.markdown(f"""
    <div class='metric-card {classe_card}'>
        <div class='titulo'>{titulo}</div>
        <div class='valor'>{valor}</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)

var_pgto = (linha_alvo["PGTO"] / media_pgto_base - 1) * 100 if media_pgto_base else 0
var_desc = (linha_alvo["DESC"] / media_desc_base - 1) * 100 if media_desc_base else 0
var_liq = (linha_alvo["Liquido"] / media_liq_base - 1) * 100 if media_liq_base else 0
diff_9950 = linha_alvo["Liquido"] - linha_alvo["Verba9950"]
pct_9950 = (diff_9950 / linha_alvo["Verba9950"] * 100) if linha_alvo["Verba9950"] else 0
classe = "" if abs(pct_9950) < 5 else ("atencao" if abs(pct_9950) < 15 else "alerta")

card(c1, f"PGTO {mes_alvo}", fmt_brl(linha_alvo["PGTO"]), f"{fmt_pct(var_pgto)} vs media baseline", "negativo" if var_pgto < -10 else ("positivo" if var_pgto > 10 else ""))
card(c2, f"DESC {mes_alvo}", fmt_brl(linha_alvo["DESC"]), f"{fmt_pct(var_desc)} vs media baseline", "negativo" if var_desc < -10 else ("positivo" if var_desc > 10 else ""))
card(c3, f"Liquido (PGTO-DESC) {mes_alvo}", fmt_brl(linha_alvo["Liquido"]), f"{fmt_pct(var_liq)} vs media baseline", "negativo" if var_liq < -10 else ("positivo" if var_liq > 10 else ""))
card(c4, "Verba 9950 (Liquido Mensal)", fmt_brl(linha_alvo["Verba9950"]), f"Diferenca p/ calculado: {fmt_brl(diff_9950)} ({fmt_pct(pct_9950)})", "negativo" if abs(pct_9950) > 5 else "positivo", classe_card=classe)

st.markdown("### Evolucao mensal - PGTO, DESC, Liquido e Verba 9950")
fig1 = go.Figure()
fig1.add_trace(go.Bar(name="PGTO", x=resumo["Mes"], y=resumo["PGTO"], marker_color=IG_AZUL, hovertemplate="<b>%{x}</b><br>PGTO: R$ %{y:,.2f}<extra></extra>"))
fig1.add_trace(go.Bar(name="DESC", x=resumo["Mes"], y=resumo["DESC"], marker_color=IG_VINHO, hovertemplate="<b>%{x}</b><br>DESC: R$ %{y:,.2f}<extra></extra>"))
fig1.add_trace(go.Scatter(name="Liquido (PGTO-DESC)", x=resumo["Mes"], y=resumo["Liquido"], mode="lines+markers", line=dict(color=IG_AZUL_ESC, width=2.5), marker=dict(size=8), yaxis="y2", hovertemplate="<b>%{x}</b><br>Liquido: R$ %{y:,.2f}<extra></extra>"))
v9950_pts = resumo[resumo["Verba9950"] > 0]
if not v9950_pts.empty:
    fig1.add_trace(go.Scatter(name="Verba 9950", x=v9950_pts["Mes"], y=v9950_pts["Verba9950"], mode="markers", marker=dict(color=IG_LARANJA, size=14, symbol="diamond", line=dict(color="white", width=1.5)), yaxis="y2", hovertemplate="<b>%{x}</b><br>Verba 9950: R$ %{y:,.2f}<extra></extra>"))

idx_alvo = list(resumo["Mes"]).index(mes_alvo)
fig1.add_vrect(x0=idx_alvo - 0.5, x1=idx_alvo + 0.5, fillcolor=IG_AZUL_CLR, opacity=0.15, line_width=0)
fig1.add_annotation(x=idx_alvo, y=1.06, xref="x", yref="paper", text="<b>Folha calculada</b>", showarrow=False, font=dict(color=IG_AZUL_ESC, size=11))
layout1 = plotly_layout_brand("", height=460)
layout1["barmode"] = "group"
layout1["yaxis"] = dict(title="PGTO / DESC (R$)", gridcolor=GRID_CINZA, linecolor=GRID_CINZA, tickformat=",.0f")
layout1["yaxis2"] = dict(title="Liquido (R$)", overlaying="y", side="right", gridcolor=GRID_CINZA, linecolor=GRID_CINZA, tickformat=",.0f")
layout1["legend"] = dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0)
fig1.update_layout(**layout1)
st.plotly_chart(fig1, use_container_width=True)

st.markdown("### Headcount (verba 0020 - Armazena Salario) e salario medio mensal")
st.caption("HC = quantidade de matriculas distintas com a verba 0020 (Armazena Salario, classe OUTRO) > 0 no mes. Salario medio = total da verba 0020 dividido pelo HC do mes.")

linha_alvo_v20 = salario_v20[salario_v20["Mes"] == mes_alvo].iloc[0]
v20_baseline = salario_v20[salario_v20["Mes"].isin(meses_baseline)]
hc_medio_base = v20_baseline["HC"].mean() if not v20_baseline.empty else 0
sal_medio_base = v20_baseline["Salario_Medio"].mean() if not v20_baseline.empty else 0

cv1, cv2, cv3 = st.columns(3)
var_hc = (linha_alvo_v20["HC"] / hc_medio_base - 1) * 100 if hc_medio_base else 0
var_sal = (linha_alvo_v20["Salario_Medio"] / sal_medio_base - 1) * 100 if sal_medio_base else 0
classe_hc = "" if abs(var_hc) < 3 else ("atencao" if abs(var_hc) < 8 else "alerta")
classe_sal = "" if abs(var_sal) < 3 else ("atencao" if abs(var_sal) < 8 else "alerta")
card(cv1, f"HC verba 0020 em {mes_alvo}", f"{int(linha_alvo_v20['HC']):,}".replace(",", "."), f"{fmt_pct(var_hc)} vs media baseline", "negativo" if var_hc < 0 else ("positivo" if var_hc > 0 else ""), classe_card=classe_hc)
card(cv2, f"Salario medio em {mes_alvo}", fmt_brl(linha_alvo_v20["Salario_Medio"]), f"{fmt_pct(var_sal)} vs media baseline", "negativo" if var_sal < 0 else ("positivo" if var_sal > 0 else ""), classe_card=classe_sal)
card(cv3, f"Total verba 0020 em {mes_alvo}", fmt_brl(linha_alvo_v20["Salario_Total"]))

fig_v20 = go.Figure()
fig_v20.add_trace(go.Bar(
    name="HC (verba 0020)",
    x=salario_v20["Mes"],
    y=salario_v20["HC"],
    marker_color=IG_AZUL_CLR,
    marker_line=dict(color=IG_AZUL_ESC, width=0.6),
    text=[f"{int(v):,}".replace(",", ".") for v in salario_v20["HC"]],
    textposition="outside",
    textfont=dict(family="Roboto", color=TEXTO_PRINCIPAL, size=10),
    hovertemplate="<b>%{x}</b><br>HC: %{y}<extra></extra>",
))
fig_v20.add_trace(go.Scatter(
    name="Salario medio (R$)",
    x=salario_v20["Mes"],
    y=salario_v20["Salario_Medio"],
    mode="lines+markers",
    line=dict(color=IG_AZUL_ESC, width=2.6),
    marker=dict(size=9, color=IG_AZUL_ESC, line=dict(color="white", width=1)),
    yaxis="y2",
    hovertemplate="<b>%{x}</b><br>Salario medio: R$ %{y:,.2f}<extra></extra>",
))
if hc_medio_base:
    fig_v20.add_hline(y=hc_medio_base, line_dash="dot", line_color=IG_VERDE_ESC, annotation_text=f"HC medio baseline: {int(hc_medio_base)}", annotation_position="top left")

idx_alvo_v20 = list(salario_v20["Mes"]).index(mes_alvo)
fig_v20.add_vrect(x0=idx_alvo_v20 - 0.5, x1=idx_alvo_v20 + 0.5, fillcolor=IG_AZUL_CLR, opacity=0.18, line_width=0)
fig_v20.add_annotation(x=idx_alvo_v20, y=1.06, xref="x", yref="paper", text="<b>Folha calculada</b>", showarrow=False, font=dict(color=IG_AZUL_ESC, size=11))

layout_v20 = plotly_layout_brand("", height=440)
layout_v20["yaxis"] = dict(title="HC (quantidade)", gridcolor=GRID_CINZA, linecolor=GRID_CINZA, tickformat=",d")
layout_v20["yaxis2"] = dict(title="Salario medio (R$)", overlaying="y", side="right", gridcolor=GRID_CINZA, linecolor=GRID_CINZA, tickformat=",.0f")
layout_v20["legend"] = dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0)
fig_v20.update_layout(**layout_v20)
st.plotly_chart(fig_v20, use_container_width=True)

col_conc, col_status = st.columns([1,1])
with col_conc:
    st.markdown(f"### Conciliacao em {mes_alvo}")
    valores = [linha_alvo["PGTO"], -linha_alvo["DESC"], linha_alvo["Liquido"], linha_alvo["Verba9950"]]
    rotulos = ["PGTO", "DESC", "Liquido (PGTO-DESC)", "Verba 9950 (declarado)"]
    cores = [IG_AZUL, IG_VINHO, IG_AZUL_ESC, IG_LARANJA]
    fig_conc = go.Figure(go.Bar(x=rotulos, y=valores, marker_color=cores, text=[fmt_brl(v) for v in valores], textposition="outside", textfont=dict(family="Roboto", color=TEXTO_PRINCIPAL, size=12), hovertemplate="<b>%{x}</b><br>%{text}<extra></extra>"))
    layout_conc = plotly_layout_brand("", height=380)
    layout_conc["yaxis"]["tickformat"] = ",.0f"
    layout_conc["showlegend"] = False
    fig_conc.update_layout(**layout_conc)
    st.plotly_chart(fig_conc, use_container_width=True)

with col_status:
    st.markdown("### Distribuicao de status")
    contagem = liquido_analise["STATUS"].value_counts().reset_index()
    contagem.columns = ["Status", "Funcionarios"]
    contagem["cor"] = contagem["Status"].map(cor_status).fillna(GRID_CINZA)
    fig_st = go.Figure(go.Bar(x=contagem["Status"], y=contagem["Funcionarios"], marker_color=contagem["cor"], text=contagem["Funcionarios"], textposition="outside", textfont=dict(family="Roboto", color=TEXTO_PRINCIPAL, size=12), hovertemplate="<b>%{x}</b><br>%{y} funcionarios<extra></extra>"))
    layout_st = plotly_layout_brand("", height=380)
    layout_st["showlegend"] = False
    fig_st.update_layout(**layout_st)
    st.plotly_chart(fig_st, use_container_width=True)

st.markdown(f"### Comparacao do liquido {mes_alvo} com meses selecionados")
meses_disponiveis_comp = [m for m in meses_detectados if m != mes_alvo]
default_comp = meses_baseline[-3:] if len(meses_baseline) >= 3 else meses_baseline
meses_comp = st.multiselect(f"Selecione os meses para comparar com {mes_alvo}", options=meses_disponiveis_comp, default=default_comp, key="meses_comp")
if meses_comp:
    valores_comp, rotulos_comp, cores_comp = [], [], []
    for m in meses_comp:
        linha = resumo[resumo["Mes"] == m].iloc[0]
        valores_comp.append(linha["Liquido"]); rotulos_comp.append(m); cores_comp.append(IG_AZUL_CLR)
    valores_comp.append(linha_alvo["Liquido"]); rotulos_comp.append(f"{mes_alvo} (alvo)"); cores_comp.append(IG_LARANJA)
    fig_comp = go.Figure(go.Bar(x=rotulos_comp, y=valores_comp, marker_color=cores_comp, text=[fmt_brl(v) for v in valores_comp], textposition="outside", textfont=dict(family="Roboto", color=TEXTO_PRINCIPAL, size=11), hovertemplate="<b>%{x}</b><br>Liquido total: %{text}<extra></extra>"))
    media_comp = sum(valores_comp[:-1]) / len(meses_comp)
    fig_comp.add_hline(y=media_comp, line_dash="dot", line_color=IG_VERDE_ESC, annotation_text=f"Media meses selecionados: {fmt_brl(media_comp)}", annotation_position="top right")
    layout_comp = plotly_layout_brand("", height=380)
    layout_comp["yaxis"]["title"] = "Liquido total (R$)"
    layout_comp["yaxis"]["tickformat"] = ",.0f"
    layout_comp["showlegend"] = False
    fig_comp.update_layout(**layout_comp)
    st.plotly_chart(fig_comp, use_container_width=True)
    diff_total = linha_alvo["Liquido"] - media_comp
    pct_total = (diff_total / media_comp * 100) if media_comp else 0
    st.caption(f"Liquido {mes_alvo}: {fmt_brl(linha_alvo['Liquido'])} | Media selecionada: {fmt_brl(media_comp)} | Diferenca: {fmt_brl(diff_total)} ({fmt_pct(pct_total)}).")
else:
    st.caption("Selecione pelo menos um mes para gerar a comparacao.")

st.markdown("### Dispersao do liquido por funcionario - Baseline vs Folha calculada")
st.caption("Pontos vermelhos representam saida do padrao. Banda azul clara indica faixa normal de +/-30%.")
valido = liquido_analise[liquido_analise["MEDIA_BASELINE"] > 0].copy()
valido["LIQ_PLOT"] = valido["LIQUIDO_ALVO"].clip(lower=100)
valido["MAT_NOME"] = [f"Mat. {idx[0]} - {idx[1]}" for idx in valido.index]
cores_disp = valido["STATUS"].map(cor_status).fillna(GRID_CINZA)
lim_min = 200
lim_max = max(valido["MEDIA_BASELINE"].max(), valido["LIQUIDO_ALVO"].max()) * 1.5 if len(valido) else 1000

fig_disp = go.Figure()
xs = np.logspace(np.log10(lim_min), np.log10(lim_max), 80)
fig_disp.add_trace(go.Scatter(x=xs, y=xs * 1.3, mode="lines", line=dict(width=0), showlegend=False, hoverinfo="skip"))
fig_disp.add_trace(go.Scatter(x=xs, y=xs * 0.7, mode="lines", line=dict(width=0), fill="tonexty", fillcolor="rgba(110,180,220,0.18)", name="Banda +/-30%", hoverinfo="skip"))
fig_disp.add_trace(go.Scatter(x=xs, y=xs, mode="lines", line=dict(color=TEXTO_SECUNDARIO, width=1, dash="dash"), name="Referencia y=x", hoverinfo="skip"))
fig_disp.add_trace(go.Scatter(
    x=valido["MEDIA_BASELINE"], y=valido["LIQ_PLOT"], mode="markers",
    marker=dict(color=cores_disp, size=10, line=dict(color="white", width=0.7)),
    text=valido["MAT_NOME"],
    customdata=np.column_stack((valido["LIQUIDO_ALVO"], valido["STATUS"], valido["VAR_PCT"], valido["AUDITORIA"])),
    hovertemplate="<b>%{text}</b><br>Baseline: R$ %{x:,.2f}<br>Liquido " + mes_alvo + ": R$ %{customdata[0]:,.2f}<br>Variacao: %{customdata[2]:+.1f}%<br>Status: %{customdata[1]}<br><br>%{customdata[3]}<extra></extra>",
    showlegend=False
))
layout_disp = plotly_layout_brand("", height=520)
layout_disp["xaxis"] = dict(title="Media liquido baseline (R$) - log", type="log", gridcolor=GRID_CINZA, linecolor=GRID_CINZA, range=[np.log10(lim_min), np.log10(lim_max)])
layout_disp["yaxis"] = dict(title=f"Liquido {mes_alvo} (R$) - log", type="log", gridcolor=GRID_CINZA, linecolor=GRID_CINZA, range=[np.log10(lim_min / 2), np.log10(lim_max)])
layout_disp["legend"] = dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="left", x=0)
fig_disp.update_layout(**layout_disp)
st.plotly_chart(fig_disp, use_container_width=True)

st.markdown("### Top outliers - Maior impacto em reais")
criticos = liquido_analise[liquido_analise["STATUS"].isin(list(STATUS_CRITICOS))].copy()
if not criticos.empty:
    criticos["IMPACTO_ABS"] = criticos["VAR_ABS"].abs()
    top = criticos.sort_values("IMPACTO_ABS", ascending=False).head(15)
    rotulos = [f"Mat. {idx[0]} - {idx[1][:32]}" for idx in top.index]
    fig_top = go.Figure(go.Bar(
        x=top["VAR_ABS"], y=rotulos, orientation="h",
        marker_color=IG_VERMELHO,
        text=[f"{fmt_brl(v)}  [{s}]" for v, s in zip(top["VAR_ABS"], top["STATUS"])],
        textposition="auto",
        textfont=dict(color="white", family="Roboto", size=11),
        customdata=np.column_stack((top["VAR_PCT"], top["AUDITORIA"], top["STATUS"])),
        hovertemplate="<b>%{y}</b><br>Variacao: %{x:,.2f}<br>Var pct: %{customdata[0]:+.1f}%<br>Status: %{customdata[2]}<br><br>%{customdata[1]}<extra></extra>"
    ))
    layout_top = plotly_layout_brand("", height=520)
    layout_top["xaxis"]["title"] = f"Variacao absoluta ({mes_alvo} - media baseline)"
    layout_top["xaxis"]["tickformat"] = ",.0f"
    layout_top["yaxis"]["autorange"] = "reversed"
    layout_top["showlegend"] = False
    layout_top["margin"]["l"] = 240
    fig_top.update_layout(**layout_top)
    st.plotly_chart(fig_top, use_container_width=True)
else:
    st.success(f"Nenhum outlier critico em {mes_alvo}.")

col_zp, col_zd = st.columns(2)
def grafico_verbas_zeradas(dados, titulo, cor, container):
    container.markdown(f"### {titulo}")
    if dados.empty:
        container.info(f"Nenhuma verba regular zerou em {mes_alvo}.")
        return
    rotulos = [f"{idx[0]} - {idx[1][:30]}" for idx in dados.index]
    fig = go.Figure(go.Bar(x=dados["MEDIA_BASELINE"], y=rotulos, orientation="h", marker_color=cor, text=[fmt_brl(v) for v in dados["MEDIA_BASELINE"]], textposition="auto", textfont=dict(color="white", family="Roboto", size=10), hovertemplate="<b>%{y}</b><br>Media baseline: %{x:,.2f}<extra></extra>"))
    layout = plotly_layout_brand("", height=320)
    layout["xaxis"]["title"] = "Media baseline (R$)"
    layout["xaxis"]["tickformat"] = ",.0f"
    layout["yaxis"]["autorange"] = "reversed"
    layout["showlegend"] = False
    layout["margin"]["l"] = 220
    fig.update_layout(**layout)
    container.plotly_chart(fig, use_container_width=True)

grafico_verbas_zeradas(zerados_pgto, f"PGTO regulares zeradas em {mes_alvo}", IG_AZUL, col_zp)
grafico_verbas_zeradas(zerados_desc, f"DESC regulares zeradas em {mes_alvo}", IG_VINHO, col_zd)

st.markdown("### Tabela detalhada por funcionario")
cf1, cf2, cf3, cf4 = st.columns([1.5, 1.4, 2.2, 1])

opcoes_status = sorted(liquido_analise["STATUS"].unique())
defaults_preferidos = ["AUSENTE", "EXTREMA", "NEGATIVO", "ZERO_SUSPEITO", "ALTA", "Z_2SIGMA"]
defaults_validos = [s for s in defaults_preferidos if s in opcoes_status]

if "filtro_status_tabela" not in st.session_state:
    st.session_state["filtro_status_tabela"] = defaults_validos
if "ordenar_por_tabela" not in st.session_state:
    st.session_state["ordenar_por_tabela"] = "VAR_ABS (impacto)"
if "buscar_nome" not in st.session_state:
    st.session_state["buscar_nome"] = ""

cf1.multiselect("Filtrar por status", options=opcoes_status, key="filtro_status_tabela")
cf2.selectbox("Ordenar por", options=["VAR_ABS (impacto)", "VAR_PCT (variacao %)", "LIQUIDO_ALVO", "MEDIA_BASELINE"], key="ordenar_por_tabela")
cf3.text_input("Buscar nome ou matricula", key="buscar_nome", placeholder="Digite parte do nome ou da matricula")
cf4.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
cf4.button("Limpar escolhas", key="btn_limpar_tabela", on_click=limpar_filtros_tabela, use_container_width=True)

filtro_status = st.session_state.get("filtro_status_tabela", [])
ordenar_por = st.session_state.get("ordenar_por_tabela", "VAR_ABS (impacto)")
buscar_nome = st.session_state.get("buscar_nome", "")

tab = liquido_analise.reset_index()
if filtro_status:
    tab = tab[tab["STATUS"].isin(filtro_status)]
if buscar_nome.strip():
    txt = buscar_nome.strip().upper()
    tab = tab[tab["Nome"].str.upper().str.contains(txt) | tab["Matrícula"].astype(str).str.contains(txt)]

ordem_map = {
    "VAR_ABS (impacto)": ("VAR_ABS", True),
    "VAR_PCT (variacao %)": ("VAR_PCT", True),
    "LIQUIDO_ALVO": ("LIQUIDO_ALVO", False),
    "MEDIA_BASELINE": ("MEDIA_BASELINE", False),
}
col_ord, ascendente = ordem_map[ordenar_por]
if "impacto" in ordenar_por or "variacao" in ordenar_por:
    tab = tab.reindex(tab[col_ord].abs().sort_values(ascending=False).index)
else:
    tab = tab.sort_values(col_ord, ascending=ascendente)

cols_show_base = ["Matrícula", "Nome", "MEDIA_BASELINE", "LIQUIDO_ALVO", "VAR_ABS", "VAR_PCT", "Z_SCORE", "DATA_RESCISAO", "STATUS", "AUDITORIA"]
if meses_comp:
    cols_meses = [m for m in meses_comp if m in tab.columns]
    cols_show = ["Matrícula", "Nome"] + cols_meses + cols_show_base[2:]
else:
    cols_meses = []
    cols_show = cols_show_base

tab_show = tab[cols_show].copy()
novos_nomes = ["Mat", "Nome"] + [f"Liq {m}" for m in cols_meses] + ["Media baseline", f"Liquido {mes_alvo}", "Var abs", "Var %", "Z-score", "Data Rescisao", "Status", "Auditoria"]
tab_show.columns = novos_nomes
col_config = {
    "Media baseline": st.column_config.NumberColumn(format="R$ %.2f"),
    f"Liquido {mes_alvo}": st.column_config.NumberColumn(format="R$ %.2f"),
    "Var abs": st.column_config.NumberColumn(format="R$ %.2f"),
    "Var %": st.column_config.NumberColumn(format="%+.1f%%"),
    "Z-score": st.column_config.NumberColumn(format="%+.2f"),
    "Data Rescisao": st.column_config.DateColumn(format="DD/MM/YYYY"),
}
for m in cols_meses:
    col_config[f"Liq {m}"] = st.column_config.NumberColumn(format="R$ %.2f")

st.dataframe(tab_show, use_container_width=True, height=420, column_config=col_config, hide_index=True)
st.caption(f"Linhas exibidas: {len(tab_show)} de {len(liquido_analise)} (total no CSV: {len(liquido)})")

st.markdown("### Drill-down por funcionario")
if "func_escolhido" not in st.session_state:
    st.session_state["func_escolhido"] = ""
for k, default in [("dd_status", []), ("dd_cr", []), ("dd_classe", []), ("dd_processo", []), ("dd_busca_verba", "")]:
    if k not in st.session_state:
        st.session_state[k] = default

fd0, fd1, fd2, fd3, fd4, fd5 = st.columns([2.4, 1.2, 1.4, 1.4, 2, 1])
opcoes_func = [f"Mat. {idx[0]} - {idx[1]}" for idx in liquido.index]
fd0.selectbox("Selecione um funcionario", options=[""] + opcoes_func, key="func_escolhido")
fd1.multiselect("Status", options=sorted(liquido["STATUS"].unique()), key="dd_status")

func_escolhido = st.session_state.get("func_escolhido", "")
mat_alvo, nome_alvo = None, None
func_df = pd.DataFrame()
cr_options, classe_options, proc_options = [], [], []

if func_escolhido:
    partes = func_escolhido.replace("Mat. ", "").split(" - ", 1)
    mat_alvo = partes[0]
    nome_alvo = partes[1]
    func_df = df[(df["Matrícula"] == mat_alvo) & (df["Nome"] == nome_alvo)].copy()
    cr_col = "C.R." if "C.R." in func_df.columns else None
    if cr_col:
        cr_options = sorted([str(x) for x in func_df[cr_col].dropna().astype(str).unique() if str(x).strip()])
    classe_options = sorted([str(x) for x in func_df["Clas."].dropna().astype(str).unique() if str(x).strip()])
    proc_options = sorted([str(x) for x in func_df["Processo"].dropna().astype(str).unique() if str(x).strip()])

fd2.multiselect("C.R.", options=cr_options, key="dd_cr", disabled=not bool(cr_options))
fd3.multiselect("Classe", options=classe_options, key="dd_classe", disabled=not bool(classe_options))
fd4.text_input("Buscar verba", key="dd_busca_verba", placeholder="Codigo ou descricao")
fd5.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
fd5.button("Limpar escolhas", key="btn_limpar_drilldown", on_click=limpar_drilldown, use_container_width=True)

fd6 = st.columns([2.4, 1.8, 1])[1]
with fd6:
    st.multiselect("Processo", options=proc_options, key="dd_processo", disabled=not bool(proc_options))

func_escolhido = st.session_state.get("func_escolhido", "")
if func_escolhido:
    partes = func_escolhido.replace("Mat. ", "").split(" - ", 1)
    mat_alvo = partes[0]
    nome_alvo = partes[1]
    serie_func = liquido.loc[(mat_alvo, nome_alvo)]

    filtros_status_dd = st.session_state.get("dd_status", [])
    if filtros_status_dd and serie_func["STATUS"] not in filtros_status_dd:
        st.warning("O funcionario selecionado esta fora do filtro de status atual no drill-down.")

    cdd1, cdd2, cdd3 = st.columns(3)
    cdd1.metric("Media baseline", fmt_brl(serie_func["MEDIA_BASELINE"]))
    cdd2.metric(f"Liquido {mes_alvo}", fmt_brl(serie_func["LIQUIDO_ALVO"]), fmt_pct(serie_func["VAR_PCT"]))
    cdd3.metric("Status", str(serie_func["STATUS"]))
    st.caption(serie_func["AUDITORIA"])

    serie_meses = serie_func[meses_detectados].copy()
    df_serie = pd.DataFrame({"Mes": serie_meses.index, "Liquido": serie_meses.values})
    cor_linha = IG_VERMELHO if serie_func["STATUS"] in STATUS_CRITICOS else IG_AZUL
    cor_ponto = IG_VERMELHO if serie_func["STATUS"] in STATUS_CRITICOS else IG_AZUL_ESC
    fig_func = go.Figure()
    fig_func.add_trace(go.Scatter(x=df_serie["Mes"], y=df_serie["Liquido"], mode="lines+markers", line=dict(color=cor_linha, width=2.6), marker=dict(size=9, color=cor_ponto), hovertemplate="<b>%{x}</b><br>Liquido: R$ %{y:,.2f}<extra></extra>", name="Liquido"))
    if pd.notna(serie_func["MEDIA_BASELINE"]):
        fig_func.add_hline(y=serie_func["MEDIA_BASELINE"], line_dash="dash", line_color=IG_VERDE_ESC, annotation_text=f"Media baseline: {fmt_brl(serie_func['MEDIA_BASELINE'])}", annotation_position="top right")
    idx_alvo_f = list(df_serie["Mes"]).index(mes_alvo)
    fig_func.add_vrect(x0=idx_alvo_f - 0.5, x1=idx_alvo_f + 0.5, fillcolor=IG_AZUL_CLR, opacity=0.18, line_width=0)
    layout_f = plotly_layout_brand("", height=350)
    layout_f["yaxis"]["title"] = "Liquido (R$)"
    layout_f["yaxis"]["tickformat"] = ",.0f"
    layout_f["showlegend"] = False
    fig_func.update_layout(**layout_f)
    st.plotly_chart(fig_func, use_container_width=True)

    st.markdown(f"**Verbas do funcionario em {mes_alvo}:**")
    func_df = df[(df["Matrícula"] == mat_alvo) & (df["Nome"] == nome_alvo)].copy()
    col_alvo = f"{mes_alvo} - Valor"
    func_df = func_df[func_df[col_alvo] != 0].copy()

    idx_alvo_drill = meses_detectados.index(mes_alvo)
    meses_3_anteriores = meses_detectados[max(0, idx_alvo_drill - 3):idx_alvo_drill]
    cols_3m_valores = [f"{m} - Valor" for m in meses_3_anteriores]

    cr_col = "C.R." if "C.R." in func_df.columns else None
    filtros_cr = st.session_state.get("dd_cr", [])
    filtros_classe = st.session_state.get("dd_classe", [])
    filtros_proc = st.session_state.get("dd_processo", [])
    busca_verba = st.session_state.get("dd_busca_verba", "").strip().upper()

    if cr_col and filtros_cr:
        func_df = func_df[func_df[cr_col].astype(str).isin(filtros_cr)]
    if filtros_classe:
        func_df = func_df[func_df["Clas."].astype(str).isin(filtros_classe)]
    if filtros_proc:
        func_df = func_df[func_df["Processo"].astype(str).isin(filtros_proc)]
    if busca_verba:
        func_df = func_df[
            func_df["Código"].astype(str).str.upper().str.contains(busca_verba, na=False) |
            func_df["Descrição"].astype(str).str.upper().str.contains(busca_verba, na=False)
        ]

    cols_func = ["Código", "Descrição", "Clas.", "Processo"] + cols_3m_valores + [col_alvo]
    if cr_col:
        cols_func.insert(3, cr_col)
    func_df = func_df[cols_func].copy()

    ordem_clas = ["PGTO", "DESC", "OUTRO"]
    outros_clas = [c for c in func_df["Clas."].astype(str).unique() if c not in ordem_clas]
    ordem_completa = ordem_clas + sorted(outros_clas)
    func_df["Clas."] = pd.Categorical(func_df["Clas."], categories=ordem_completa, ordered=True)
    sort_cols = ["Clas.", "Código"]
    if cr_col:
        sort_cols = ["Clas.", cr_col, "Código"]
    func_df = func_df.sort_values(sort_cols).reset_index(drop=True)

    rename_map = {"Código": "Codigo", "Descrição": "Descricao", "Clas.": "Clas.", "Processo": "Processo", col_alvo: f"Valor {mes_alvo}"}
    if cr_col:
        rename_map[cr_col] = "C.R."
    for m, col_m in zip(meses_3_anteriores, cols_3m_valores):
        rename_map[col_m] = f"Valor {m}"
    func_df = func_df.rename(columns=rename_map)

    col_config_drill = {f"Valor {mes_alvo}": st.column_config.NumberColumn(format="R$ %.2f")}
    for m in meses_3_anteriores:
        col_config_drill[f"Valor {m}"] = st.column_config.NumberColumn(format="R$ %.2f")

    if len(func_df):
        st.dataframe(func_df, use_container_width=True, hide_index=True, column_config=col_config_drill)
    else:
        st.info("Nenhuma verba encontrada para os filtros selecionados.")
else:
    st.caption("Selecione um funcionario para abrir o drill-down.")

st.markdown("### Exportar relatorio")
@st.cache_data
def gerar_excel(_resumo, _liquido, _zerados_pgto, _zerados_desc):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as w:
        _resumo.to_excel(w, sheet_name="Resumo_Mensal", index=False)
        liq = _liquido.reset_index()
        ordem = ["AUSENTE", "NEGATIVO", "ZERO_SUSPEITO", "EXTREMA", "Z_2SIGMA", "ALTA", "NOVO_FUNC", "OK", "SEM_DADOS", "DESLIGADO"]
        liq["STATUS"] = pd.Categorical(liq["STATUS"], categories=ordem, ordered=True)
        liq = liq.sort_values(["STATUS", "VAR_ABS"], na_position="last")
        liq.to_excel(w, sheet_name="Liquido_Funcionario", index=False)
        _zerados_pgto.reset_index().to_excel(w, sheet_name="Verbas_PGTO_Zeradas", index=False)
        _zerados_desc.reset_index().to_excel(w, sheet_name="Verbas_DESC_Zeradas", index=False)
    buffer.seek(0)
    return buffer.getvalue()

excel_bytes = gerar_excel(resumo, liquido, zerados_pgto, zerados_desc)
st.download_button(
    label="Baixar relatorio Excel completo",
    data=excel_bytes,
    file_name=f"auditoria_folha_{mes_alvo.replace('/','-')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.markdown(f"""
<div class='app-footer'>
Igarape Digital | Anderson Marinho<br>
Auditoria de Folha - regras DP padrao
</div>
""", unsafe_allow_html=True)
