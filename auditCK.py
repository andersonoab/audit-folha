# =============================================================================
# Auditoria de Folha - Desktop App (CustomTkinter)
# Igarape Digital
# Padrao: Anderson Marinho
# Port do app Streamlit original mantendo todas as funcionalidades
# Parametros de outliers baseados em literatura estatistica de mercado
# =============================================================================

import io
import os
import sys
import threading
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

import customtkinter as ctk
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.ticker import FuncFormatter

import numpy as np
import pandas as pd

# -----------------------------------------------------------------------------
# Paleta Igarape Digital (mesma do app Streamlit)
# -----------------------------------------------------------------------------
IG_AZUL          = "#0083CA"
IG_AZUL_ESC      = "#003C64"
IG_VERDE_ESC     = "#005A64"
IG_AZUL_CLR      = "#6EB4DC"
IG_VINHO         = "#7D0041"
IG_LARANJA       = "#8C321E"
IG_VERMELHO      = "#7D0041"
TEXTO_PRINCIPAL  = "#333333"
TEXTO_SECUNDARIO = "#646464"
GRID_CINZA       = "#CCCCCC"
FUNDO            = "#FFFFFF"
COR_DESLIGADO    = "#646464"

STATUS_CRITICOS = {"AUSENTE", "NEGATIVO", "ZERO_SUSPEITO", "EXTREMA", "ALTA", "Z_2SIGMA",
                   "EXTREMA_IQR", "ALTA_IQR", "MAD_OUTLIER", "MATERIAL"}

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# -----------------------------------------------------------------------------
# PARAMETROS ESTATISTICOS DE MERCADO
# -----------------------------------------------------------------------------
# Referencias bibliograficas:
#  - Shewhart (1931) / Western Electric Rules (1956): SPC, 3-sigma rule
#  - Tukey, J.W. (1977) Exploratory Data Analysis: IQR 1.5x e 3x
#  - Iglewicz, B. & Hoaglin, D.C. (1993): MAD modified, threshold 3.5
#  - ISA 320 (IFAC) / NBC TA 320 (CFC): materialidade 0.5%-5% da base
# -----------------------------------------------------------------------------

PARAMETROS_MERCADO = {
    "Shewhart 3-sigma + Materialidade ISA 320 (recomendado)": {
        "metodo": "sigma",
        "alta": 2.0,
        "extrema": 3.0,
        "materialidade_pct": 0.01,
        "descricao": "Padrao SPC industrial (Western Electric). |z|>=3 ~ 0.3% esperado. ISA 320: 1% da folha total."
    },
    "Tukey IQR (boxplot classico)": {
        "metodo": "iqr",
        "alta": 1.5,
        "extrema": 3.0,
        "materialidade_pct": 0.01,
        "descricao": "Tukey 1977. Robusto a distribuicoes nao-normais. 1.5x IQR = outlier moderado, 3x = extremo."
    },
    "MAD modificado (Iglewicz-Hoaglin)": {
        "metodo": "mad",
        "alta": 2.5,
        "extrema": 3.5,
        "materialidade_pct": 0.01,
        "descricao": "Mediana absoluta de desvio. Mais robusto que sigma. Threshold 3.5 (paper original 1993)."
    },
    "Anderson Legacy (98%/30%)": {
        "metodo": "legacy",
        "alta_pct": 30, "alta_abs": 200,
        "extrema_pct": 98, "extrema_abs": 500,
        "materialidade_pct": 0.0,
        "descricao": "Heuristica original Igarape Digital. Mantida para compatibilidade."
    },
}

# -----------------------------------------------------------------------------
# Utilitarios
# -----------------------------------------------------------------------------
def fmt_brl(x):
    if pd.isna(x):
        return "-"
    s = f"{abs(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-R$ {s}" if x < 0 else f"R$ {s}"

def fmt_pct(x):
    if pd.isna(x):
        return "-"
    return f"{x:+.1f}%"

def fmt_int(x):
    try:
        return f"{int(x):,}".replace(",", ".")
    except Exception:
        return "-"

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

# -----------------------------------------------------------------------------
# Camada de dados (mesma logica do app original)
# -----------------------------------------------------------------------------
def carregar_csv(path):
    df = pd.read_csv(path, sep=";", encoding="latin-1", dtype=str)
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
        linhas.append({"Mes": m, "HC": hc, "Salario_Total": total, "Salario_Medio": media})
    return pd.DataFrame(linhas)


def hc_e_total_por_verba(df, meses, codigo_verba):
    """
    Para uma verba ou conjunto de verbas: retorna por mes o HC (matriculas com
    valor != 0 em qualquer das verbas), o total agregado e o valor medio.
    codigo_verba: pode ser uma string (codigo unico) ou lista de strings.
    """
    if isinstance(codigo_verba, (list, tuple, set)):
        cods_alvo = {str(c).strip().lstrip("0") or "0" for c in codigo_verba}
    else:
        cods_alvo = {str(codigo_verba).strip().lstrip("0") or "0"}
    codigos_norm = df["Código"].astype(str).str.strip().str.lstrip("0").replace("", "0")
    sub = df[codigos_norm.isin(cods_alvo)]
    linhas = []
    for m in meses:
        col = f"{m} - Valor"
        ativos = sub[sub[col] != 0]
        # HC = matriculas unicas com valor != 0 em QUALQUER uma das verbas
        hc = ativos["Matrícula"].nunique()
        total = ativos[col].sum()
        media = (total / hc) if hc else 0
        linhas.append({"Mes": m, "HC": hc, "Total": total, "Media": media})
    return pd.DataFrame(linhas)


def impacto_por_verba(df, meses, mes_alvo, meses_baseline, top_n=20):
    """
    Calcula impacto/variacao de cada verba: total no mes-alvo vs media baseline.
    Retorna DataFrame ordenado por classe (PGTO -> DESC -> OUTRO) e dentro da
    classe por |VAR_ABS| descendente.
    """
    val_cols = [f"{m} - Valor" for m in meses]
    col_alvo = f"{mes_alvo} - Valor"
    cols_base = [f"{m} - Valor" for m in meses_baseline]
    if not cols_base:
        return pd.DataFrame()
    agg = df.groupby(["Código", "Descrição", "Clas."])[val_cols].sum().reset_index()
    agg["MEDIA_BASELINE"] = agg[cols_base].replace(0, np.nan).mean(axis=1)
    agg["VALOR_ALVO"] = agg[col_alvo]
    agg["VAR_ABS"] = agg["VALOR_ALVO"] - agg["MEDIA_BASELINE"].fillna(0)
    agg["VAR_PCT"] = ((agg["VALOR_ALVO"] / agg["MEDIA_BASELINE"]) - 1) * 100
    agg["IMPACTO_ABS"] = agg["VAR_ABS"].abs()
    agg["COD_NORM"] = agg["Código"].astype(str).str.strip()
    # Ordem por classe (PGTO/DESC/OUTRO) e dentro da classe por impacto
    ordem_clas = {"PGTO": 0, "DESC": 1, "OUTRO": 2}
    agg["_ord_clas"] = agg["Clas."].map(lambda c: ordem_clas.get(str(c), 99))
    # Pega top_n por impacto absoluto, depois reordena por classe
    top = agg.sort_values("IMPACTO_ABS", ascending=False).head(top_n)
    top = top.sort_values(["_ord_clas", "IMPACTO_ABS"],
                            ascending=[True, False]).reset_index(drop=True)
    return top.drop(columns=["_ord_clas"])


def liquido_por_funcionario(df, meses, mes_alvo, meses_baseline, desligados=None,
                             config_param=None):
    """
    Calcula liquido por funcionario e classifica outliers usando o metodo selecionado.
    config_param: dict de PARAMETROS_MERCADO.
    """
    desligados = desligados or {}
    config_param = config_param or PARAMETROS_MERCADO["Shewhart 3-sigma + Materialidade ISA 320 (recomendado)"]

    val_cols = [f"{m} - Valor" for m in meses]
    pgto_func = df[df["Clas."] == "PGTO"].groupby(["Matrícula", "Nome"])[val_cols].sum()
    desc_func = df[df["Clas."] == "DESC"].groupby(["Matrícula", "Nome"])[val_cols].sum()
    liquido = pgto_func.subtract(desc_func, fill_value=0)
    liquido.columns = [c.replace(" - Valor", "") for c in liquido.columns]

    base = liquido[meses_baseline].replace(0, np.nan)
    liquido["MEDIA_BASELINE"] = base.mean(axis=1)
    liquido["DESVIO_BASELINE"] = base.std(axis=1)
    liquido["MEDIANA_BASELINE"] = base.median(axis=1)

    # MAD individual = mediana(|x_mes - mediana_baseline|) para cada funcionario
    def calc_mad_individual(row):
        vals = row[meses_baseline].dropna()
        vals = vals[vals != 0]
        if len(vals) < 2:
            return np.nan
        med = vals.median()
        return (vals - med).abs().median()
    liquido["MAD_INDIVIDUAL"] = liquido.apply(calc_mad_individual, axis=1)

    liquido["LIQUIDO_ALVO"] = liquido[mes_alvo]
    liquido["VAR_ABS"] = liquido["LIQUIDO_ALVO"] - liquido["MEDIA_BASELINE"]
    liquido["VAR_PCT"] = ((liquido["LIQUIDO_ALVO"] / liquido["MEDIA_BASELINE"]) - 1) * 100
    liquido["Z_SCORE"] = liquido["VAR_ABS"] / liquido["DESVIO_BASELINE"]

    # Estatisticas globais para IQR (Tukey) - calculadas sobre VAR_ABS de todos
    var_abs_validos = liquido["VAR_ABS"].dropna()
    if len(var_abs_validos) > 4:
        q1 = var_abs_validos.quantile(0.25)
        q3 = var_abs_validos.quantile(0.75)
        iqr = q3 - q1
    else:
        q1, q3, iqr = np.nan, np.nan, np.nan

    # Materialidade: % da folha liquida total do mes-alvo
    folha_total_alvo = liquido[mes_alvo].abs().sum()
    materialidade = folha_total_alvo * config_param.get("materialidade_pct", 0.0)

    liquido["DATA_RESCISAO"] = [desligados.get((mat, nome), pd.NaT) for (mat, nome) in liquido.index]

    metodo = config_param.get("metodo", "sigma")

    def status(r):
        # Regras universais primeiro
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

        # Regras estatisticas conforme metodo
        if metodo == "sigma":
            if pd.notna(r["Z_SCORE"]):
                if abs(r["Z_SCORE"]) >= config_param["extrema"]:
                    return "EXTREMA"
                if abs(r["Z_SCORE"]) >= config_param["alta"]:
                    return "ALTA"

        elif metodo == "iqr":
            if pd.notna(r["VAR_ABS"]) and pd.notna(iqr) and iqr > 0:
                if r["VAR_ABS"] < q1 - config_param["extrema"] * iqr or r["VAR_ABS"] > q3 + config_param["extrema"] * iqr:
                    return "EXTREMA_IQR"
                if r["VAR_ABS"] < q1 - config_param["alta"] * iqr or r["VAR_ABS"] > q3 + config_param["alta"] * iqr:
                    return "ALTA_IQR"

        elif metodo == "mad":
            if pd.notna(r["MAD_INDIVIDUAL"]) and r["MAD_INDIVIDUAL"] > 0 and pd.notna(r["MEDIANA_BASELINE"]):
                z_mad = 0.6745 * (r["LIQUIDO_ALVO"] - r["MEDIANA_BASELINE"]) / r["MAD_INDIVIDUAL"]
                if abs(z_mad) >= config_param["extrema"]:
                    return "EXTREMA"
                if abs(z_mad) >= config_param["alta"]:
                    return "ALTA"

        elif metodo == "legacy":
            if abs(r["VAR_PCT"]) >= config_param["extrema_pct"] and abs(r["VAR_ABS"]) > config_param["extrema_abs"]:
                return "EXTREMA"
            if abs(r["VAR_PCT"]) >= config_param["alta_pct"] and abs(r["VAR_ABS"]) > config_param["alta_abs"]:
                return "ALTA"

        # Materialidade ISA 320 - sobreposta
        if materialidade > 0 and abs(r["VAR_ABS"]) >= materialidade:
            return "MATERIAL"

        # Z 2-sigma como ultimo recurso (sempre)
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
        if s in ("EXTREMA", "EXTREMA_IQR"):
            return f"Variacao extrema: {fmt_pct(r['VAR_PCT'])} ({fmt_brl(r['VAR_ABS'])}). Z={r['Z_SCORE']:+.2f}."
        if s in ("ALTA", "ALTA_IQR"):
            return f"Variacao alta: {fmt_pct(r['VAR_PCT'])} ({fmt_brl(r['VAR_ABS'])}). Z={r['Z_SCORE']:+.2f}."
        if s == "MATERIAL":
            return f"Impacto material (>={config_param.get('materialidade_pct', 0)*100:.1f}% da folha): {fmt_brl(r['VAR_ABS'])}."
        if s == "Z_2SIGMA":
            return f"Fora de 2 sigmas (z={r['Z_SCORE']:+.2f})."
        if s == "NOVO_FUNC":
            return f"Sem baseline; liquido {fmt_brl(r['LIQUIDO_ALVO'])}."
        if s == "SEM_DADOS":
            return "Sem movimento no periodo."
        return "Sem desvio relevante."

    liquido["AUDITORIA"] = liquido.apply(auditoria, axis=1)

    metadata = {
        "metodo": metodo,
        "q1": q1, "q3": q3, "iqr": iqr,
        "folha_total_alvo": folha_total_alvo,
        "materialidade": materialidade,
    }
    return liquido, metadata


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


# -----------------------------------------------------------------------------
# Componentes de UI reutilizaveis
# -----------------------------------------------------------------------------
class MetricCard(ctk.CTkFrame):
    def __init__(self, master, titulo="", valor="-", delta="", classe=""):
        super().__init__(master, fg_color="white", corner_radius=4,
                         border_width=1, border_color=GRID_CINZA, height=92)
        self.grid_propagate(False)

        cor_borda = IG_VERMELHO if classe == "alerta" else (IG_VINHO if classe == "atencao" else IG_AZUL)
        # Faixa lateral colorida
        self.faixa = ctk.CTkFrame(self, fg_color=cor_borda, width=4, corner_radius=0)
        self.faixa.place(x=0, y=0, relheight=1)

        self.lbl_titulo = ctk.CTkLabel(self, text=titulo.upper(), font=("Roboto", 10),
                                        text_color=TEXTO_SECUNDARIO, anchor="w")
        self.lbl_titulo.place(x=14, y=10, relwidth=0.95)

        # Valor: agora SelectableEntry para permitir selecao parcial e copia
        self.lbl_valor = SelectableEntry(self, text=valor,
                                           font=("Rufina", 20, "bold"),
                                           fg=IG_AZUL_ESC, bg="white")
        self.lbl_valor.place(x=14, y=30, relwidth=0.95)

        cor_delta = IG_VERMELHO if "negativo" in classe else IG_VERDE_ESC
        # Delta tambem selecionavel (1 linha)
        self.lbl_delta = SelectableEntry(self, text=delta,
                                           font=("Roboto", 11),
                                           fg=cor_delta, bg="white")
        self.lbl_delta.place(x=14, y=62, relwidth=0.95)

        # Right-click no card todo: copia "Titulo: Valor (delta)"
        def _texto_completo():
            t = self.lbl_titulo.cget("text")
            v = self.lbl_valor.cget("text")
            d = self.lbl_delta.cget("text")
            return f"{t}: {v}  ({d})" if d else f"{t}: {v}"
        for w in (self.lbl_titulo, self):
            make_label_copiable(w, _texto_completo)

    def atualizar(self, titulo=None, valor=None, delta=None, classe=""):
        if titulo is not None: self.lbl_titulo.configure(text=titulo.upper())
        if valor is not None: self.lbl_valor.set_text(valor)
        if delta is not None:
            cor_delta = IG_VERMELHO if "negativo" in classe else IG_VERDE_ESC
            self.lbl_delta.set_text(delta)
            self.lbl_delta.configure(fg=cor_delta)
        cor_borda = IG_VERMELHO if classe == "alerta" else (IG_VINHO if classe == "atencao" else IG_AZUL)
        self.faixa.configure(fg_color=cor_borda)


class ChartContainer(ctk.CTkFrame):
    """Frame que hospeda um Figure matplotlib com toolbar opcional."""
    def __init__(self, master, height=420, com_toolbar=False, **kw):
        super().__init__(master, fg_color="white", corner_radius=4,
                         border_width=1, border_color=GRID_CINZA, **kw)
        self.fig = Figure(figsize=(8, height/72), dpi=96, facecolor="white")
        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=4, pady=4)
        if com_toolbar:
            tb_frame = ctk.CTkFrame(self, fg_color="white", height=28)
            tb_frame.pack(fill="x", side="bottom")
            self.toolbar = NavigationToolbar2Tk(self.canvas, tb_frame, pack_toolbar=False)
            self.toolbar.update()
            self.toolbar.pack(side="left")

    def limpar(self):
        self.fig.clear()

    def render(self):
        self.canvas.draw_idle()


def estilo_eixos(ax, titulo=""):
    ax.set_facecolor("white")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#888888")
    ax.spines["bottom"].set_color("#888888")
    ax.spines["left"].set_linewidth(1.0)
    ax.spines["bottom"].set_linewidth(1.0)
    # Fontes maiores e mais escuras para contraste
    ax.tick_params(colors="#1A1A1A", labelsize=10, width=1.0, length=4)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontweight("medium")
    ax.grid(True, color=GRID_CINZA, linestyle="-", linewidth=0.6, alpha=0.7)
    ax.set_axisbelow(True)
    if titulo:
        ax.set_title(titulo, fontsize=12, color=IG_AZUL_ESC, fontweight="bold", loc="left", pad=12)


def texto_label(ax, x, y, texto, **kw):
    """Helper para rotulos de dados com contraste consistente."""
    defaults = dict(fontsize=9, color="#1A1A1A", fontweight="bold",
                    ha="center", va="center", clip_on=False)
    defaults.update(kw)
    return ax.text(x, y, texto, **defaults)


def fmt_eixo_brl(x, pos):
    if abs(x) >= 1_000_000:
        return f"{x/1_000_000:.1f}M"
    if abs(x) >= 1_000:
        return f"{x/1_000:.0f}k"
    return f"{x:.0f}"


# -----------------------------------------------------------------------------
# Utilitarios de copia para clipboard
# -----------------------------------------------------------------------------
def _copy_to_clipboard(widget, texto):
    """Copia texto para o clipboard usando o root do widget."""
    try:
        root = widget.winfo_toplevel()
        root.clipboard_clear()
        root.clipboard_append(texto)
        root.update()
    except Exception:
        pass


def make_label_copiable(label, text_getter=None):
    """
    Adiciona menu de contexto (botao direito) ao CTkLabel para copiar o texto.
    text_getter: callable opcional que retorna o texto a copiar (default: cget('text')).
    """
    def _show_menu(event):
        try:
            txt = text_getter() if text_getter else label.cget("text")
        except Exception:
            txt = ""
        menu = tk.Menu(label, tearoff=0, bg="white", fg=TEXTO_PRINCIPAL,
                       activebackground=IG_AZUL, activeforeground="white",
                       font=("Roboto", 10), bd=1, relief="solid")
        menu.add_command(label="  Copiar  ",
                         command=lambda: _copy_to_clipboard(label, txt))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    label.bind("<Button-3>", _show_menu)  # Windows/Linux
    label.bind("<Button-2>", _show_menu)  # MacOS / fallback
    # Cursor sutil para indicar interatividade
    try:
        label.configure(cursor="hand2")
    except Exception:
        pass


def bind_treeview_copy(tree):
    """
    Habilita Ctrl+C no Treeview para copiar a(s) linha(s) selecionada(s) como TSV
    (cola direto no Excel). Tambem adiciona menu de contexto botao direito.
    """
    def _copiar_selecionado(_event=None):
        sel = tree.selection()
        if not sel:
            return
        linhas = []
        for item in sel:
            valores = tree.item(item, "values")
            linhas.append("\t".join(str(v) for v in valores))
        _copy_to_clipboard(tree, "\n".join(linhas))

    def _copiar_tudo_visivel(_event=None):
        linhas = []
        # Cabecalho
        cabec = [tree.heading(c, "text") for c in tree["columns"]]
        linhas.append("\t".join(cabec))
        for item in tree.get_children():
            valores = tree.item(item, "values")
            linhas.append("\t".join(str(v) for v in valores))
        _copy_to_clipboard(tree, "\n".join(linhas))

    def _show_menu(event):
        item = tree.identify_row(event.y)
        if item and item not in tree.selection():
            tree.selection_set(item)
        menu = tk.Menu(tree, tearoff=0, bg="white", fg=TEXTO_PRINCIPAL,
                       activebackground=IG_AZUL, activeforeground="white",
                       font=("Roboto", 10), bd=1, relief="solid")
        menu.add_command(label="  Copiar linha selecionada (Ctrl+C)  ",
                         command=_copiar_selecionado)
        menu.add_command(label="  Copiar tabela inteira (visivel)  ",
                         command=_copiar_tudo_visivel)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    tree.bind("<Control-c>", _copiar_selecionado)
    tree.bind("<Control-C>", _copiar_selecionado)
    tree.bind("<Button-3>", _show_menu)
    tree.bind("<Button-2>", _show_menu)


# -----------------------------------------------------------------------------
# Widgets de texto SELECIONAVEL (permitem arrastar para selecionar trecho e Ctrl+C)
# Usados nos pontos onde o usuario costuma copiar trechos para outra ferramenta (ADP)
# -----------------------------------------------------------------------------
class SelectableEntry(tk.Entry):
    """
    Substituto de Label para textos curtos (1 linha) que permite selecao e copia.
    Visualmente parece um label, mas o usuario pode arrastar para selecionar.
    """
    def __init__(self, master, text="", font=("Roboto", 11),
                 fg=TEXTO_PRINCIPAL, bg="white", **kw):
        kw.setdefault("borderwidth", 0)
        kw.setdefault("highlightthickness", 0)
        kw.setdefault("relief", "flat")
        super().__init__(master,
                          font=font, fg=fg,
                          background=bg, readonlybackground=bg,
                          disabledbackground=bg,
                          **kw)
        self._var = tk.StringVar(value=text)
        self.configure(textvariable=self._var, state="readonly", cursor="xterm")

    def cget(self, key):
        if key == "text":
            return self._var.get()
        return super().cget(key)

    def configure(self, **kw):
        if "text" in kw:
            self.set_text(kw.pop("text"))
        if "text_color" in kw:
            kw["fg"] = kw.pop("text_color")
        if "fg_color" in kw:
            cor = kw.pop("fg_color")
            kw["background"] = cor
            kw["readonlybackground"] = cor
        if kw:
            super().configure(**kw)

    def set_text(self, text):
        super().configure(state="normal")
        self._var.set(text)
        super().configure(state="readonly")


class SelectableText(tk.Text):
    """
    Substituto de Label para textos longos (multi-linha) que permite selecao parcial
    e Ctrl+C. Bloqueia edicao via bindings (state="disabled" desabilitaria tambem
    a selecao em algumas plataformas).
    """
    def __init__(self, master, text="", height=1, wrap="word",
                 font=("Roboto", 10), fg=TEXTO_PRINCIPAL, bg="white", **kw):
        kw.setdefault("borderwidth", 0)
        kw.setdefault("highlightthickness", 0)
        kw.setdefault("relief", "flat")
        super().__init__(master,
                          height=height, wrap=wrap,
                          font=font, fg=fg, background=bg,
                          cursor="xterm",
                          **kw)
        self.insert("1.0", text)
        # Bloqueia edicao mas mantem selecao
        self.bind("<Key>", self._on_key)
        # Menu de contexto
        self.bind("<Button-3>", self._show_menu)
        self.bind("<Button-2>", self._show_menu)

    def _on_key(self, event):
        # Permite combinacoes com Control (Ctrl+C, Ctrl+A)
        if event.state & 0x4:
            return
        # Permite navegacao
        if event.keysym in ("Left", "Right", "Up", "Down", "Home", "End",
                             "Prior", "Next", "Shift_L", "Shift_R",
                             "Control_L", "Control_R", "Alt_L", "Alt_R"):
            return
        return "break"

    def _show_menu(self, event):
        try:
            sel = self.get("sel.first", "sel.last")
        except tk.TclError:
            sel = self.get("1.0", "end-1c")
        menu = tk.Menu(self, tearoff=0, bg="white", fg=TEXTO_PRINCIPAL,
                       activebackground=IG_AZUL, activeforeground="white",
                       font=("Roboto", 10), bd=1, relief="solid")
        menu.add_command(label="  Copiar selecao  ",
                         command=lambda: _copy_to_clipboard(self, sel))
        menu.add_command(label="  Copiar tudo  ",
                         command=lambda: _copy_to_clipboard(self, self.get("1.0", "end-1c")))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def cget(self, key):
        if key == "text":
            return self.get("1.0", "end-1c")
        return super().cget(key)

    def configure(self, **kw):
        if "text" in kw:
            self.set_text(kw.pop("text"))
        if "text_color" in kw:
            kw["fg"] = kw.pop("text_color")
        if "fg_color" in kw:
            kw["background"] = kw.pop("fg_color")
        if "wraplength" in kw:
            kw.pop("wraplength")  # tk.Text usa wrap mode, nao wraplength
        if "justify" in kw:
            kw.pop("justify")  # tk.Text nao usa justify igual label
        if "font" in kw and isinstance(kw["font"], tuple):
            pass  # ja compativel
        if kw:
            super().configure(**kw)

    def set_text(self, text):
        self.delete("1.0", "end")
        self.insert("1.0", text)


# -----------------------------------------------------------------------------
# Dialogo modal para filtrar verbas
# -----------------------------------------------------------------------------
class FiltroVerbasDialog(ctk.CTkToplevel):
    """
    Janela modal que lista todas as verbas com checkbox para incluir/excluir do calculo.
    A verba 0020 (Armazena Salario) e sempre preservada, pois e usada para HC.
    """
    def __init__(self, master, df, meses, verbas_excluidas_atual):
        super().__init__(master)
        self.title("Gerenciar verbas - Auditoria de Folha")
        self.geometry("960x640")
        self.minsize(800, 520)
        self.configure(fg_color=FUNDO)
        self.transient(master)
        self.grab_set()

        self.df = df
        self.meses = meses
        self.verbas_excluidas_inicial = set(verbas_excluidas_atual)
        # resultado=None significa que o usuario cancelou
        self.resultado = None

        # Pre-calcula os totais por verba no periodo todo
        val_cols = [f"{m} - Valor" for m in meses]
        agg = df.groupby(["Código", "Descrição", "Clas."])[val_cols].sum()
        agg["TOTAL_PERIODO"] = agg.sum(axis=1)
        agg["N_MESES_ATIVOS"] = (agg[val_cols].abs() > 0.01).sum(axis=1)
        agg = agg.reset_index()
        # Normaliza codigo
        agg["COD_NORM"] = agg["Código"].astype(str).str.strip().str.lstrip("0").replace("", "0")
        # Ordem fixa por classe: PGTO, DESC, OUTRO, depois resto
        ordem_clas = {"PGTO": 0, "DESC": 1, "OUTRO": 2}
        agg["_ord_clas"] = agg["Clas."].map(lambda c: ordem_clas.get(str(c), 99))
        # Ordena por classe (PGTO/DESC/OUTRO) e dentro da classe por valor absoluto descendente
        agg = agg.sort_values(["_ord_clas", "TOTAL_PERIODO"],
                                ascending=[True, False]).reset_index(drop=True)
        agg = agg.drop(columns=["_ord_clas"])
        self.dados = agg

        # Estado: codigos atualmente excluidos (sera modificado ao clicar)
        self.excluidos = set(self.verbas_excluidas_inicial)

        self._build_ui()
        self._popular_tabela()

    def _build_ui(self):
        # Header
        header = ctk.CTkFrame(self, fg_color=IG_AZUL, corner_radius=0, height=54)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(header, text="Gerenciar verbas",
                     font=("Rufina", 15, "bold"), text_color="white").pack(side="left", padx=18, pady=8)
        ctk.CTkLabel(header, text="Marque a coluna 'Inc?' para incluir; desmarque para excluir do calculo",
                     font=("Roboto", 10), text_color="white").pack(side="left", padx=8, pady=8)

        # Toolbar de filtros
        toolbar = ctk.CTkFrame(self, fg_color="white", height=64,
                                border_width=1, border_color=GRID_CINZA)
        toolbar.pack(fill="x", padx=8, pady=8)

        ctk.CTkLabel(toolbar, text="Classe:", font=("Roboto", 10)).grid(row=0, column=0, padx=(10, 4), pady=10, sticky="e")
        classes = ["(todas)"] + sorted([c for c in self.dados["Clas."].dropna().astype(str).unique()])
        self.combo_classe = ctk.CTkOptionMenu(toolbar, values=classes,
                                                fg_color="white", text_color=TEXTO_PRINCIPAL,
                                                button_color=IG_AZUL, button_hover_color=IG_AZUL_ESC,
                                                command=lambda _: self._popular_tabela())
        self.combo_classe.grid(row=0, column=1, padx=4, pady=10)

        ctk.CTkLabel(toolbar, text="Buscar:", font=("Roboto", 10)).grid(row=0, column=2, padx=(12, 4), pady=10, sticky="e")
        self.entry_busca = ctk.CTkEntry(toolbar, placeholder_text="codigo ou descricao", width=220)
        self.entry_busca.grid(row=0, column=3, padx=4, pady=10)
        self.entry_busca.bind("<Return>", lambda _: self._popular_tabela())

        ctk.CTkButton(toolbar, text="Aplicar", command=self._popular_tabela,
                      fg_color=IG_AZUL, hover_color=IG_AZUL_ESC, width=80).grid(row=0, column=4, padx=4, pady=10)

        # Acoes em massa
        ctk.CTkLabel(toolbar, text="  |  Acoes:", font=("Roboto", 10),
                     text_color=TEXTO_SECUNDARIO).grid(row=0, column=5, padx=4, pady=10, sticky="e")
        ctk.CTkButton(toolbar, text="Incluir todas (visiveis)", command=self._incluir_todas_visiveis,
                      fg_color="white", text_color=IG_AZUL_ESC, hover_color="#E8F0F7",
                      border_width=1, border_color=IG_AZUL, width=140).grid(row=0, column=6, padx=4, pady=10)
        ctk.CTkButton(toolbar, text="Excluir todas (visiveis)", command=self._excluir_todas_visiveis,
                      fg_color="white", text_color=IG_VINHO, hover_color="#FBEBF1",
                      border_width=1, border_color=IG_VINHO, width=140).grid(row=0, column=7, padx=4, pady=10)
        ctk.CTkButton(toolbar, text="Inverter (visiveis)", command=self._inverter_visiveis,
                      fg_color="white", text_color=TEXTO_PRINCIPAL, hover_color="#F0F0F0",
                      border_width=1, border_color=GRID_CINZA, width=120).grid(row=0, column=8, padx=4, pady=10)

        # Treeview
        tree_frame = ctk.CTkFrame(self, fg_color="white",
                                   border_width=1, border_color=GRID_CINZA)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=4)

        cols = ("Inc", "Codigo", "Descricao", "Clas", "Total", "N_Meses")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=18)
        self.tree.heading("Inc", text="Inc?")
        self.tree.heading("Codigo", text="Codigo")
        self.tree.heading("Descricao", text="Descricao")
        self.tree.heading("Clas", text="Clas.")
        self.tree.heading("Total", text="Total no periodo")
        self.tree.heading("N_Meses", text="N° meses ativos")

        self.tree.column("Inc", width=50, anchor="center")
        self.tree.column("Codigo", width=80, anchor="center")
        self.tree.column("Descricao", width=380, anchor="w")
        self.tree.column("Clas", width=80, anchor="center")
        self.tree.column("Total", width=140, anchor="e")
        self.tree.column("N_Meses", width=110, anchor="center")

        sy = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure("excluida", background="#FBEBF1", foreground=IG_VINHO)
        self.tree.tag_configure("incluida", background="white")
        self.tree.tag_configure("protegida", background="#E8F0F7", foreground=IG_AZUL_ESC)

        # Toggle ao clicar na linha
        self.tree.bind("<Button-1>", self._on_click)
        self.tree.bind("<space>", self._on_space)
        bind_treeview_copy(self.tree)

        # Footer
        footer = ctk.CTkFrame(self, fg_color="white", height=58,
                               border_width=1, border_color=GRID_CINZA)
        footer.pack(fill="x", padx=8, pady=(0, 8))
        footer.pack_propagate(False)

        self.lbl_status = ctk.CTkLabel(footer, text="", font=("Roboto", 10),
                                         text_color=TEXTO_SECUNDARIO)
        self.lbl_status.pack(side="left", padx=14, pady=12)

        ctk.CTkButton(footer, text="Cancelar", command=self._cancelar,
                      fg_color="white", text_color=TEXTO_PRINCIPAL,
                      hover_color="#F0F0F0", border_width=1, border_color=GRID_CINZA,
                      width=110).pack(side="right", padx=8, pady=10)
        ctk.CTkButton(footer, text="Aplicar e fechar", command=self._aplicar,
                      fg_color=IG_AZUL, hover_color=IG_AZUL_ESC, width=140).pack(side="right", padx=4, pady=10)

    def _filtrar_dados(self):
        df = self.dados.copy()
        cl = self.combo_classe.get()
        if cl and cl != "(todas)":
            df = df[df["Clas."].astype(str) == cl]
        busca = self.entry_busca.get().strip().upper()
        if busca:
            df = df[
                df["Código"].astype(str).str.upper().str.contains(busca, na=False) |
                df["Descrição"].astype(str).str.upper().str.contains(busca, na=False)
            ]
        return df

    def _popular_tabela(self):
        for it in self.tree.get_children():
            self.tree.delete(it)
        df = self._filtrar_dados()
        for _, row in df.iterrows():
            cod = str(row["Código"])
            cod_norm = str(row["COD_NORM"])
            eh_protegida = (cod_norm == "20")  # 0020 sempre incluida
            esta_excluida = (cod in self.excluidos) and not eh_protegida
            inc_marker = "—" if eh_protegida else ("Nao" if esta_excluida else "Sim")
            tag = "protegida" if eh_protegida else ("excluida" if esta_excluida else "incluida")
            self.tree.insert("", "end", iid=cod, values=(
                inc_marker,
                cod,
                str(row["Descrição"])[:80],
                str(row["Clas."]),
                fmt_brl(row["TOTAL_PERIODO"]),
                int(row["N_MESES_ATIVOS"]),
            ), tags=(tag,))
        self._atualizar_status()

    def _atualizar_status(self):
        total = len(self.dados)
        excl = len(self.excluidos)
        # Recalcula impacto: total das verbas excluidas no periodo
        if excl:
            df_excl = self.dados[self.dados["Código"].astype(str).isin(self.excluidos)]
            df_excl = df_excl[df_excl["COD_NORM"] != "20"]  # ignora 0020
            total_excl = df_excl["TOTAL_PERIODO"].sum()
            self.lbl_status.configure(
                text=f"{excl} de {total} verbas excluidas  |  Impacto no periodo: {fmt_brl(total_excl)}",
                text_color=IG_VINHO)
        else:
            self.lbl_status.configure(
                text=f"Todas as {total} verbas estao incluidas no calculo",
                text_color=IG_VERDE_ESC)

    def _toggle_codigo(self, cod):
        cod_norm = str(cod).strip().lstrip("0") or "0"
        if cod_norm == "20":
            return  # nao permite alterar 0020
        if cod in self.excluidos:
            self.excluidos.discard(cod)
        else:
            self.excluidos.add(cod)

    def _on_click(self, event):
        # Toggle apenas quando clica na coluna Inc (#1)
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#1":
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self._toggle_codigo(item)
        self._popular_tabela()

    def _on_space(self, event):
        # Espaco alterna a linha selecionada
        sel = self.tree.selection()
        for item in sel:
            self._toggle_codigo(item)
        self._popular_tabela()

    def _incluir_todas_visiveis(self):
        df = self._filtrar_dados()
        for _, row in df.iterrows():
            cod = str(row["Código"])
            self.excluidos.discard(cod)
        self._popular_tabela()

    def _excluir_todas_visiveis(self):
        df = self._filtrar_dados()
        for _, row in df.iterrows():
            cod = str(row["Código"])
            cod_norm = str(row["COD_NORM"])
            if cod_norm != "20":
                self.excluidos.add(cod)
        self._popular_tabela()

    def _inverter_visiveis(self):
        df = self._filtrar_dados()
        for _, row in df.iterrows():
            cod = str(row["Código"])
            cod_norm = str(row["COD_NORM"])
            if cod_norm == "20":
                continue
            if cod in self.excluidos:
                self.excluidos.discard(cod)
            else:
                self.excluidos.add(cod)
        self._popular_tabela()

    def _aplicar(self):
        self.resultado = set(self.excluidos)
        self.destroy()

    def _cancelar(self):
        self.resultado = None
        self.destroy()


# -----------------------------------------------------------------------------
# Aplicacao principal
# -----------------------------------------------------------------------------
class AuditoriaFolhaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auditoria de Folha - Igarape Digital")
        self.geometry("1480x900")
        self.minsize(1280, 800)
        self.configure(fg_color=FUNDO)

        # Estado
        self.df = None
        self.meses_detectados = []
        self.empresas_disponiveis = []
        self.processos_disponiveis = []
        self.empresas_selecionadas = []
        self.processos_selecionados = []
        self.mes_alvo = None
        self.meses_baseline = []
        self.excluir_desligados = True
        self.config_param_nome = "Shewhart 3-sigma + Materialidade ISA 320 (recomendado)"
        self.materialidade_pct_user = 1.0  # %
        self.verbas_excluidas = set()  # Codigos de verba a excluir do calculo

        # Resultados
        self.resumo = None
        self.liquido = None
        self.liquido_analise = None
        self.salario_v20 = None
        self.zerados_pgto = None
        self.zerados_desc = None
        self.metadata_audit = {}

        # Estado especifico da aba Analise de Verbas
        self.av_codigos_sel = []
        self.av_mes_clicado = None
        self.av_dados_plot = None
        self.av_colab_df = pd.DataFrame()
        self.av_colab_selecionado = None
        self.av_top_verbas_bars = []
        self.av_top_verbas_cid = None
        self.av_hc_cid = None

        # Foco visual quando navegar da Analise de Verbas para Outliers
        self.outlier_foco = None

        self._build_layout()

    # ---------- Layout principal ----------
    def _build_layout(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Header
        header = ctk.CTkFrame(self, fg_color=IG_AZUL, corner_radius=0, height=70)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        header.grid_propagate(False)
        ctk.CTkLabel(header, text="Auditoria de Folha - Confronto Folha Calculada vs Implantadas",
                     font=("Rufina", 18, "bold"), text_color="white").place(x=22, y=12)
        ctk.CTkLabel(header, text="Igarape Digital | Analise de Payroll | Padrao DP",
                     font=("Roboto", 11), text_color="white").place(x=22, y=42)

        # Sidebar
        self.sidebar = ctk.CTkScrollableFrame(self, fg_color="#F7F9FB", width=320,
                                              corner_radius=0)
        self.sidebar.grid(row=1, column=0, sticky="nsw")
        self._build_sidebar()

        # Main area com tabview
        self.main = ctk.CTkFrame(self, fg_color=FUNDO, corner_radius=0)
        self.main.grid(row=1, column=1, sticky="nsew")
        self.main.grid_rowconfigure(0, weight=1)
        self.main.grid_columnconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(self.main, fg_color="white", corner_radius=4,
                                       segmented_button_selected_color=IG_AZUL,
                                       segmented_button_selected_hover_color=IG_AZUL_ESC)
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        for nome in ["Visao Geral", "Headcount", "Outliers", "Tabela & Drill-down",
                      "Verbas Zeradas", "Analise de Verbas", "Exportar"]:
            self.tabview.add(nome)

        self._build_tab_visao_geral()
        self._build_tab_headcount()
        self._build_tab_outliers()
        self._build_tab_tabela_drill()
        self._build_tab_verbas_zeradas()
        self._build_tab_analise_verbas()
        self._build_tab_exportar()

        # Footer
        footer = ctk.CTkFrame(self, fg_color="white", height=28, border_width=1,
                              border_color=GRID_CINZA, corner_radius=0)
        footer.grid(row=2, column=0, columnspan=2, sticky="ew")
        footer.grid_propagate(False)
        ctk.CTkLabel(footer, text="Igarape Digital | Anderson Marinho  |  Auditoria de Folha - regras DP padrao",
                     font=("Roboto", 10), text_color=TEXTO_SECUNDARIO).pack(pady=6)

        # Estado inicial
        self._mostrar_estado_inicial()

    # ---------- Sidebar ----------
    def _build_sidebar(self):
        sb = self.sidebar
        ctk.CTkLabel(sb, text="Configuracao", font=("Rufina", 16, "bold"),
                     text_color=IG_AZUL_ESC).pack(anchor="w", padx=14, pady=(14, 6))

        # Upload
        self.btn_upload = ctk.CTkButton(sb, text="Carregar arquivo CSV...",
                                         command=self.acao_carregar_csv,
                                         fg_color=IG_AZUL, hover_color=IG_AZUL_ESC)
        self.btn_upload.pack(fill="x", padx=14, pady=4)
        self.lbl_arquivo = ctk.CTkLabel(sb, text="(nenhum arquivo)", font=("Roboto", 10),
                                         text_color=TEXTO_SECUNDARIO, wraplength=290, justify="left")
        self.lbl_arquivo.pack(anchor="w", padx=14, pady=(0, 8))

        ctk.CTkLabel(sb, text="Exporte do ADP em CSV pt-BR (separador ; / decimal virgula).",
                     font=("Roboto", 9), text_color=TEXTO_SECUNDARIO,
                     wraplength=290, justify="left").pack(anchor="w", padx=14, pady=(0, 10))

        # Separador
        ctk.CTkFrame(sb, height=1, fg_color=GRID_CINZA).pack(fill="x", padx=14, pady=8)

        # Filtros (criados como containers vazios; populados apos carga)
        ctk.CTkLabel(sb, text="Filtros", font=("Rufina", 13, "bold"),
                     text_color=IG_AZUL_ESC).pack(anchor="w", padx=14, pady=(4, 4))

        self.frame_empresa = ctk.CTkFrame(sb, fg_color="transparent")
        self.frame_empresa.pack(fill="x", padx=14, pady=4)

        self.frame_processo = ctk.CTkFrame(sb, fg_color="transparent")
        self.frame_processo.pack(fill="x", padx=14, pady=4)

        ctk.CTkLabel(sb, text="Mes-alvo (folha calculada)",
                     font=("Roboto", 11), text_color=TEXTO_PRINCIPAL).pack(anchor="w", padx=14, pady=(8, 2))
        self.combo_mes_alvo = ctk.CTkOptionMenu(sb, values=["-"], command=lambda _: self._on_mes_alvo_changed(),
                                                 fg_color="white", text_color=TEXTO_PRINCIPAL,
                                                 button_color=IG_AZUL, button_hover_color=IG_AZUL_ESC)
        self.combo_mes_alvo.pack(fill="x", padx=14, pady=2)

        ctk.CTkLabel(sb, text="Meses de baseline (folhas implantadas)",
                     font=("Roboto", 11), text_color=TEXTO_PRINCIPAL).pack(anchor="w", padx=14, pady=(8, 2))
        self.frame_baseline = ctk.CTkScrollableFrame(sb, fg_color="white", height=140,
                                                      border_width=1, border_color=GRID_CINZA)
        self.frame_baseline.pack(fill="x", padx=14, pady=2)
        self.checks_baseline = {}

        self.switch_excluir = ctk.CTkSwitch(sb, text="Excluir desligados da analise",
                                             onvalue=True, offvalue=False,
                                             font=("Roboto", 11),
                                             progress_color=IG_AZUL,
                                             command=self._on_excluir_toggle)
        self.switch_excluir.select()
        self.switch_excluir.pack(anchor="w", padx=14, pady=(10, 4))

        # Filtro de verbas
        self.btn_filtro_verbas = ctk.CTkButton(sb, text="Filtrar verbas...",
                                                command=self.acao_filtrar_verbas,
                                                fg_color="white", text_color=IG_AZUL_ESC,
                                                hover_color="#E8F0F7",
                                                border_width=1, border_color=IG_AZUL,
                                                state="disabled")
        self.btn_filtro_verbas.pack(fill="x", padx=14, pady=(8, 2))
        self.lbl_verbas_status = ctk.CTkLabel(sb, text="(todas as verbas incluidas)",
                                                font=("Roboto", 9), text_color=TEXTO_SECUNDARIO,
                                                wraplength=290, justify="left")
        self.lbl_verbas_status.pack(anchor="w", padx=14, pady=(0, 4))

        # Parametros estatisticos
        ctk.CTkFrame(sb, height=1, fg_color=GRID_CINZA).pack(fill="x", padx=14, pady=8)
        ctk.CTkLabel(sb, text="Parametros de outlier", font=("Rufina", 13, "bold"),
                     text_color=IG_AZUL_ESC).pack(anchor="w", padx=14, pady=(2, 2))

        self.combo_metodo = ctk.CTkOptionMenu(sb, values=list(PARAMETROS_MERCADO.keys()),
                                               command=self._on_metodo_changed,
                                               fg_color="white", text_color=TEXTO_PRINCIPAL,
                                               button_color=IG_AZUL, button_hover_color=IG_AZUL_ESC,
                                               width=290, dynamic_resizing=False)
        self.combo_metodo.set(self.config_param_nome)
        self.combo_metodo.pack(fill="x", padx=14, pady=4)

        self.lbl_metodo_desc = ctk.CTkLabel(sb, text=PARAMETROS_MERCADO[self.config_param_nome]["descricao"],
                                             font=("Roboto", 9), text_color=TEXTO_SECUNDARIO,
                                             wraplength=290, justify="left")
        self.lbl_metodo_desc.pack(anchor="w", padx=14, pady=2)

        # Materialidade slider
        self.frame_materialidade = ctk.CTkFrame(sb, fg_color="transparent")
        self.frame_materialidade.pack(fill="x", padx=14, pady=(8, 4))
        ctk.CTkLabel(self.frame_materialidade, text="Materialidade ISA 320 (% folha total)",
                     font=("Roboto", 10), text_color=TEXTO_PRINCIPAL).pack(anchor="w")
        self.lbl_mat_valor = ctk.CTkLabel(self.frame_materialidade, text="1.0%",
                                           font=("Roboto", 10, "bold"), text_color=IG_AZUL_ESC)
        self.lbl_mat_valor.pack(anchor="e")
        self.slider_mat = ctk.CTkSlider(self.frame_materialidade, from_=0.0, to=5.0, number_of_steps=50,
                                         progress_color=IG_AZUL, button_color=IG_AZUL_ESC,
                                         command=self._on_mat_changed)
        self.slider_mat.set(1.0)
        self.slider_mat.pack(fill="x")

        # Botao reprocessar
        ctk.CTkFrame(sb, height=1, fg_color=GRID_CINZA).pack(fill="x", padx=14, pady=10)
        self.btn_reprocessar = ctk.CTkButton(sb, text="Reprocessar auditoria",
                                              command=self.acao_reprocessar,
                                              fg_color=IG_AZUL_ESC, hover_color=IG_AZUL,
                                              state="disabled")
        self.btn_reprocessar.pack(fill="x", padx=14, pady=6)

        # Stats
        self.lbl_stats = SelectableText(sb, text="", height=4,
                                          font=("Roboto", 9),
                                          fg=TEXTO_SECUNDARIO, bg="#F7F9FB")
        self.lbl_stats.pack(anchor="w", padx=14, pady=(8, 14), fill="x")

    # ---------- Tabs ----------
    def _build_tab_visao_geral(self):
        tab = self.tabview.tab("Visao Geral")
        self.tab_vg = ctk.CTkScrollableFrame(tab, fg_color=FUNDO)
        self.tab_vg.pack(fill="both", expand=True)

        # Cards
        cards_frame = ctk.CTkFrame(self.tab_vg, fg_color="transparent")
        cards_frame.pack(fill="x", pady=(8, 12), padx=8)
        for i in range(4):
            cards_frame.grid_columnconfigure(i, weight=1, uniform="cards")
        self.card_pgto = MetricCard(cards_frame); self.card_pgto.grid(row=0, column=0, sticky="nsew", padx=4)
        self.card_desc = MetricCard(cards_frame); self.card_desc.grid(row=0, column=1, sticky="nsew", padx=4)
        self.card_liq = MetricCard(cards_frame); self.card_liq.grid(row=0, column=2, sticky="nsew", padx=4)
        self.card_9950 = MetricCard(cards_frame); self.card_9950.grid(row=0, column=3, sticky="nsew", padx=4)

        # Aviso desligados
        self.lbl_aviso_desligados = ctk.CTkLabel(self.tab_vg, text="",
                                                  font=("Roboto", 10), text_color=TEXTO_SECUNDARIO)
        self.lbl_aviso_desligados.pack(anchor="w", padx=12)

        # Grafico evolucao
        ctk.CTkLabel(self.tab_vg, text="Evolucao mensal - PGTO, DESC, Liquido e Verba 9950",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w", padx=12, pady=(12, 4))
        self.chart_evolucao = ChartContainer(self.tab_vg, height=420)
        self.chart_evolucao.pack(fill="x", padx=8, pady=4)

        # Conciliacao + distribuicao status
        cc_frame = ctk.CTkFrame(self.tab_vg, fg_color="transparent")
        cc_frame.pack(fill="x", padx=8, pady=8)
        cc_frame.grid_columnconfigure(0, weight=1)
        cc_frame.grid_columnconfigure(1, weight=1)

        cc_left = ctk.CTkFrame(cc_frame, fg_color="transparent")
        cc_left.grid(row=0, column=0, sticky="nsew", padx=4)
        ctk.CTkLabel(cc_left, text="Conciliacao no mes-alvo",
                     font=("Rufina", 13, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w")
        self.chart_conciliacao = ChartContainer(cc_left, height=360)
        self.chart_conciliacao.pack(fill="x")

        cc_right = ctk.CTkFrame(cc_frame, fg_color="transparent")
        cc_right.grid(row=0, column=1, sticky="nsew", padx=4)
        ctk.CTkLabel(cc_right, text="Distribuicao de status",
                     font=("Rufina", 13, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w")
        self.chart_status = ChartContainer(cc_right, height=360)
        self.chart_status.pack(fill="x")

    def _build_tab_headcount(self):
        tab = self.tabview.tab("Headcount")
        self.tab_hc = ctk.CTkScrollableFrame(tab, fg_color=FUNDO)
        self.tab_hc.pack(fill="both", expand=True)

        ctk.CTkLabel(self.tab_hc, text="Headcount (verba 0020) e salario medio mensal",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w", padx=12, pady=(8, 4))
        ctk.CTkLabel(self.tab_hc,
                     text="HC = matriculas distintas com a verba 0020 (Armazena Salario) > 0 no mes. "
                          "Salario medio = total da verba 0020 dividido pelo HC do mes.",
                     font=("Roboto", 10), text_color=TEXTO_SECUNDARIO,
                     wraplength=1100, justify="left").pack(anchor="w", padx=12)

        cards_hc = ctk.CTkFrame(self.tab_hc, fg_color="transparent")
        cards_hc.pack(fill="x", padx=8, pady=10)
        for i in range(3):
            cards_hc.grid_columnconfigure(i, weight=1, uniform="hc")
        self.card_hc = MetricCard(cards_hc); self.card_hc.grid(row=0, column=0, sticky="nsew", padx=4)
        self.card_sal = MetricCard(cards_hc); self.card_sal.grid(row=0, column=1, sticky="nsew", padx=4)
        self.card_total20 = MetricCard(cards_hc); self.card_total20.grid(row=0, column=2, sticky="nsew", padx=4)

        self.chart_hc = ChartContainer(self.tab_hc, height=440)
        self.chart_hc.pack(fill="x", padx=8, pady=8)

    def _build_tab_outliers(self):
        tab = self.tabview.tab("Outliers")
        self.tab_out = ctk.CTkScrollableFrame(tab, fg_color=FUNDO)
        self.tab_out.pack(fill="both", expand=True)

        ctk.CTkLabel(self.tab_out, text="Dispersao do liquido por funcionario - Baseline vs Folha calculada",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w", padx=12, pady=(8, 4))
        ctk.CTkLabel(self.tab_out,
                     text="Pontos vermelhos = saida do padrao. Banda azul clara = +/-30%.",
                     font=("Roboto", 10), text_color=TEXTO_SECUNDARIO).pack(anchor="w", padx=12)
        self.chart_dispersao = ChartContainer(self.tab_out, height=520, com_toolbar=True)
        self.chart_dispersao.pack(fill="x", padx=8, pady=8)

        ctk.CTkLabel(self.tab_out, text="Top outliers - Maior impacto em reais",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w", padx=12, pady=(8, 4))
        self.chart_top = ChartContainer(self.tab_out, height=520)
        self.chart_top.pack(fill="x", padx=8, pady=8)

    def _build_tab_tabela_drill(self):
        tab = self.tabview.tab("Tabela & Drill-down")
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        # ---- Topo: tabela com filtros ----
        top = ctk.CTkFrame(tab, fg_color=FUNDO)
        top.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        ctk.CTkLabel(top, text="Tabela detalhada por funcionario",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w")

        filtros = ctk.CTkFrame(top, fg_color="transparent")
        filtros.pack(fill="x", pady=6)

        ctk.CTkLabel(filtros, text="Status:", font=("Roboto", 10)).grid(row=0, column=0, padx=(0, 4), sticky="e")
        self.combo_status_tab = ctk.CTkOptionMenu(filtros, values=["(todos)"],
                                                   fg_color="white", text_color=TEXTO_PRINCIPAL,
                                                   button_color=IG_AZUL, button_hover_color=IG_AZUL_ESC,
                                                   command=lambda _: self._renderizar_tabela())
        self.combo_status_tab.grid(row=0, column=1, padx=4)

        ctk.CTkLabel(filtros, text="Ordenar:", font=("Roboto", 10)).grid(row=0, column=2, padx=(12, 4), sticky="e")
        self.combo_ordem_tab = ctk.CTkOptionMenu(filtros,
                                                  values=["VAR_ABS (impacto)", "VAR_PCT (variacao %)",
                                                          "LIQUIDO_ALVO", "MEDIA_BASELINE"],
                                                  fg_color="white", text_color=TEXTO_PRINCIPAL,
                                                  button_color=IG_AZUL, button_hover_color=IG_AZUL_ESC,
                                                  command=lambda _: self._renderizar_tabela())
        self.combo_ordem_tab.set("VAR_ABS (impacto)")
        self.combo_ordem_tab.grid(row=0, column=3, padx=4)

        ctk.CTkLabel(filtros, text="Buscar:", font=("Roboto", 10)).grid(row=0, column=4, padx=(12, 4), sticky="e")
        self.entry_busca = ctk.CTkEntry(filtros, placeholder_text="nome ou matricula", width=200)
        self.entry_busca.grid(row=0, column=5, padx=4)
        self.entry_busca.bind("<Return>", lambda _: self._renderizar_tabela())

        ctk.CTkButton(filtros, text="Aplicar", command=self._renderizar_tabela,
                      fg_color=IG_AZUL, hover_color=IG_AZUL_ESC, width=80).grid(row=0, column=6, padx=4)
        ctk.CTkButton(filtros, text="Limpar", command=self._limpar_filtros_tabela,
                      fg_color=TEXTO_SECUNDARIO, hover_color=TEXTO_PRINCIPAL, width=80).grid(row=0, column=7, padx=4)

        # Treeview
        tree_frame = ctk.CTkFrame(top, fg_color="white", border_width=1, border_color=GRID_CINZA)
        tree_frame.pack(fill="both", expand=True, pady=4)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="white", foreground=TEXTO_PRINCIPAL,
                        fieldbackground="white", rowheight=22, font=("Roboto", 9))
        style.configure("Treeview.Heading", background=IG_AZUL, foreground="white",
                        font=("Roboto", 9, "bold"))
        style.map("Treeview", background=[("selected", IG_AZUL_CLR)])

        self.cols_tab = ("Mat", "Nome", "Baseline", "Liquido", "VarAbs", "VarPct", "ZScore", "Status", "Auditoria")
        self.tree = ttk.Treeview(tree_frame, columns=self.cols_tab, show="headings", height=10)
        self.tree.heading("Mat", text="Mat.")
        self.tree.heading("Nome", text="Nome")
        self.tree.heading("Baseline", text="Media Baseline")
        self.tree.heading("Liquido", text="Liquido Alvo")
        self.tree.heading("VarAbs", text="Var Abs")
        self.tree.heading("VarPct", text="Var %")
        self.tree.heading("ZScore", text="Z-score")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Auditoria", text="Auditoria")

        self.tree.column("Mat", width=70, anchor="center")
        self.tree.column("Nome", width=240, anchor="w")
        self.tree.column("Baseline", width=120, anchor="e")
        self.tree.column("Liquido", width=120, anchor="e")
        self.tree.column("VarAbs", width=110, anchor="e")
        self.tree.column("VarPct", width=80, anchor="e")
        self.tree.column("ZScore", width=80, anchor="e")
        self.tree.column("Status", width=120, anchor="center")
        self.tree.column("Auditoria", width=380, anchor="w")

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure("critico", background="#FFE8EE")
        self.tree.tag_configure("desligado", background="#F0F0F0", foreground=COR_DESLIGADO)

        self.tree.bind("<<TreeviewSelect>>", self._on_func_select)
        bind_treeview_copy(self.tree)

        self.lbl_count_tabela = ctk.CTkLabel(top, text="", font=("Roboto", 9),
                                              text_color=TEXTO_SECUNDARIO)
        self.lbl_count_tabela.pack(anchor="w", pady=(2, 0))

        # ---- Embaixo: drill-down ----
        bot = ctk.CTkFrame(tab, fg_color=FUNDO)
        bot.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)

        ctk.CTkLabel(bot, text="Drill-down do funcionario selecionado",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w")
        self.lbl_drill_func = SelectableText(bot,
                                              text="(selecione um funcionario na tabela acima)",
                                              height=2, font=("Roboto", 10),
                                              fg=TEXTO_SECUNDARIO, bg=FUNDO)
        self.lbl_drill_func.pack(anchor="w", pady=(2, 8), fill="x")

        drill_grid = ctk.CTkFrame(bot, fg_color="transparent")
        drill_grid.pack(fill="both", expand=True)
        drill_grid.grid_columnconfigure(0, weight=2)
        drill_grid.grid_columnconfigure(1, weight=3)
        drill_grid.grid_rowconfigure(0, weight=1)

        # Esquerda: grafico do funcionario
        drill_left = ctk.CTkFrame(drill_grid, fg_color="transparent")
        drill_left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.chart_drill = ChartContainer(drill_left, height=360)
        self.chart_drill.pack(fill="both", expand=True)

        # Direita: filtros + verbas
        drill_right = ctk.CTkFrame(drill_grid, fg_color="transparent")
        drill_right.grid(row=0, column=1, sticky="nsew")

        filtros_drill = ctk.CTkFrame(drill_right, fg_color="transparent")
        filtros_drill.pack(fill="x", pady=(0, 4))

        ctk.CTkLabel(filtros_drill, text="C.R.:", font=("Roboto", 10)).grid(row=0, column=0, sticky="e")
        self.combo_dd_cr = ctk.CTkOptionMenu(filtros_drill, values=["(todos)"],
                                              fg_color="white", text_color=TEXTO_PRINCIPAL,
                                              button_color=IG_AZUL, button_hover_color=IG_AZUL_ESC,
                                              command=lambda _: self._renderizar_drill_verbas(),
                                              width=140)
        self.combo_dd_cr.grid(row=0, column=1, padx=4)

        ctk.CTkLabel(filtros_drill, text="Classe:", font=("Roboto", 10)).grid(row=0, column=2, padx=(8, 0), sticky="e")
        self.combo_dd_classe = ctk.CTkOptionMenu(filtros_drill, values=["(todos)"],
                                                   fg_color="white", text_color=TEXTO_PRINCIPAL,
                                                   button_color=IG_AZUL, button_hover_color=IG_AZUL_ESC,
                                                   command=lambda _: self._renderizar_drill_verbas(),
                                                   width=120)
        self.combo_dd_classe.grid(row=0, column=3, padx=4)

        ctk.CTkLabel(filtros_drill, text="Processo:", font=("Roboto", 10)).grid(row=0, column=4, padx=(8, 0), sticky="e")
        self.combo_dd_proc = ctk.CTkOptionMenu(filtros_drill, values=["(todos)"],
                                                fg_color="white", text_color=TEXTO_PRINCIPAL,
                                                button_color=IG_AZUL, button_hover_color=IG_AZUL_ESC,
                                                command=lambda _: self._renderizar_drill_verbas(),
                                                width=120)
        self.combo_dd_proc.grid(row=0, column=5, padx=4)

        ctk.CTkLabel(filtros_drill, text="Verba:", font=("Roboto", 10)).grid(row=0, column=6, padx=(8, 0), sticky="e")
        self.entry_dd_verba = ctk.CTkEntry(filtros_drill, placeholder_text="codigo ou descricao", width=180)
        self.entry_dd_verba.grid(row=0, column=7, padx=4)
        self.entry_dd_verba.bind("<Return>", lambda _: self._renderizar_drill_verbas())

        ctk.CTkButton(filtros_drill, text="Limpar", command=self._limpar_filtros_drill,
                      fg_color=TEXTO_SECUNDARIO, hover_color=TEXTO_PRINCIPAL, width=70).grid(row=0, column=8, padx=4)

        # Treeview de verbas
        verbas_frame = ctk.CTkFrame(drill_right, fg_color="white", border_width=1, border_color=GRID_CINZA)
        verbas_frame.pack(fill="both", expand=True, pady=4)

        self.cols_drill = ("Codigo", "Descricao", "Clas", "CR", "Processo", "M3", "M2", "M1", "Alvo")
        self.tree_drill = ttk.Treeview(verbas_frame, columns=self.cols_drill, show="headings", height=12)
        for c, w, a in [("Codigo", 70, "center"), ("Descricao", 220, "w"), ("Clas", 60, "center"),
                         ("CR", 70, "center"), ("Processo", 110, "w"),
                         ("M3", 100, "e"), ("M2", 100, "e"), ("M1", 100, "e"), ("Alvo", 110, "e")]:
            self.tree_drill.heading(c, text=c)
            self.tree_drill.column(c, width=w, anchor=a)
        scroll_d = ttk.Scrollbar(verbas_frame, orient="vertical", command=self.tree_drill.yview)
        scroll_dh = ttk.Scrollbar(verbas_frame, orient="horizontal", command=self.tree_drill.xview)
        self.tree_drill.configure(yscrollcommand=scroll_d.set, xscrollcommand=scroll_dh.set)
        self.tree_drill.grid(row=0, column=0, sticky="nsew")
        scroll_d.grid(row=0, column=1, sticky="ns")
        scroll_dh.grid(row=1, column=0, sticky="ew")
        verbas_frame.grid_rowconfigure(0, weight=1)
        verbas_frame.grid_columnconfigure(0, weight=1)
        bind_treeview_copy(self.tree_drill)

        self.func_selecionado = None

    def _build_tab_verbas_zeradas(self):
        tab = self.tabview.tab("Verbas Zeradas")
        self.tab_vz = ctk.CTkScrollableFrame(tab, fg_color=FUNDO)
        self.tab_vz.pack(fill="both", expand=True)

        ctk.CTkLabel(self.tab_vz, text="Verbas regulares zeradas no mes-alvo",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w", padx=12, pady=(8, 4))
        ctk.CTkLabel(self.tab_vz,
                     text="Verbas que apareceram em ao menos 3 meses da baseline (>=R$ 500) "
                          "mas zeraram completamente no mes-alvo.",
                     font=("Roboto", 10), text_color=TEXTO_SECUNDARIO,
                     wraplength=1100, justify="left").pack(anchor="w", padx=12)

        cont = ctk.CTkFrame(self.tab_vz, fg_color="transparent")
        cont.pack(fill="x", padx=8, pady=8)
        cont.grid_columnconfigure(0, weight=1)
        cont.grid_columnconfigure(1, weight=1)

        left = ctk.CTkFrame(cont, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=4)
        ctk.CTkLabel(left, text="PGTO regulares zeradas", font=("Rufina", 12, "bold"),
                     text_color=IG_AZUL_ESC).pack(anchor="w")
        self.chart_zerados_pgto = ChartContainer(left, height=380)
        self.chart_zerados_pgto.pack(fill="x")

        right = ctk.CTkFrame(cont, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=4)
        ctk.CTkLabel(right, text="DESC regulares zeradas", font=("Rufina", 12, "bold"),
                     text_color=IG_AZUL_ESC).pack(anchor="w")
        self.chart_zerados_desc = ChartContainer(right, height=380)
        self.chart_zerados_desc.pack(fill="x")

    def _build_tab_analise_verbas(self):
        tab = self.tabview.tab("Analise de Verbas")
        self.tab_av = ctk.CTkScrollableFrame(tab, fg_color=FUNDO)
        self.tab_av.pack(fill="both", expand=True)

        # ---- Bloco 1: Top verbas por impacto (estilo top outliers) ----
        ctk.CTkLabel(self.tab_av, text="Top verbas por impacto - Variacao no mes-alvo vs baseline",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w", padx=12, pady=(8, 4))
        ctk.CTkLabel(self.tab_av,
                     text="Ranking de verbas por impacto absoluto. Verbas excluidas do calculo aparecem em cinza listrado.",
                     font=("Roboto", 10), text_color=TEXTO_SECUNDARIO,
                     wraplength=1100, justify="left").pack(anchor="w", padx=12)
        self.chart_top_verbas = ChartContainer(self.tab_av, height=560)
        self.chart_top_verbas.pack(fill="x", padx=8, pady=8)

        # ---- Bloco 2: Drill-down por verba (estilo HC verba 0020) ----
        ctk.CTkLabel(self.tab_av, text="Detalhe por verba - HC e valor medio mensal",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w", padx=12, pady=(16, 4))
        ctk.CTkLabel(self.tab_av,
                     text="Selecione 1 ou mais verbas (Ctrl+Click ou Shift+Click). "
                          "Os totais sao agregados (soma) e o HC sao matriculas unicas. "
                          "Verbas em ordem PGTO -> DESC -> OUTRO.",
                     font=("Roboto", 10), text_color=TEXTO_SECUNDARIO,
                     wraplength=1100, justify="left").pack(anchor="w", padx=12)

        sel_frame = ctk.CTkFrame(self.tab_av, fg_color="transparent")
        sel_frame.pack(fill="x", padx=12, pady=4)

        # Linha 1: busca + acoes
        linha1 = ctk.CTkFrame(sel_frame, fg_color="transparent")
        linha1.pack(fill="x", pady=2)
        ctk.CTkLabel(linha1, text="Buscar:", font=("Roboto", 11)).pack(side="left", padx=(0, 6))
        self.entry_busca_verba_av = ctk.CTkEntry(linha1, placeholder_text="codigo, descricao ou classe",
                                                   width=320)
        self.entry_busca_verba_av.pack(side="left", padx=4)
        self.entry_busca_verba_av.bind("<KeyRelease>", lambda _: self._filtrar_listbox_verbas())

        ctk.CTkButton(linha1, text="Limpar busca",
                       command=self._limpar_busca_verba_av,
                       fg_color="white", text_color=TEXTO_PRINCIPAL,
                       hover_color="#F0F0F0", border_width=1, border_color=GRID_CINZA,
                       width=110).pack(side="left", padx=4)

        ctk.CTkLabel(linha1, text="  |  ", font=("Roboto", 11),
                     text_color=GRID_CINZA).pack(side="left")

        ctk.CTkButton(linha1, text="Selecionar PGTO",
                       command=lambda: self._selecionar_classe_verbas("PGTO"),
                       fg_color="white", text_color=IG_AZUL_ESC,
                       hover_color="#E8F0F7", border_width=1, border_color=IG_AZUL,
                       width=120).pack(side="left", padx=2)
        ctk.CTkButton(linha1, text="Selecionar DESC",
                       command=lambda: self._selecionar_classe_verbas("DESC"),
                       fg_color="white", text_color=IG_VINHO,
                       hover_color="#FBEBF1", border_width=1, border_color=IG_VINHO,
                       width=120).pack(side="left", padx=2)
        ctk.CTkButton(linha1, text="Selecionar OUTRO",
                       command=lambda: self._selecionar_classe_verbas("OUTRO"),
                       fg_color="white", text_color=TEXTO_PRINCIPAL,
                       hover_color="#F0F0F0", border_width=1, border_color=TEXTO_SECUNDARIO,
                       width=120).pack(side="left", padx=2)
        ctk.CTkButton(linha1, text="Limpar selecao",
                       command=self._limpar_selecao_verbas_av,
                       fg_color="white", text_color=TEXTO_SECUNDARIO,
                       hover_color="#F0F0F0", border_width=1, border_color=GRID_CINZA,
                       width=120).pack(side="left", padx=8)

        # Linha 2: listbox + status
        linha2 = ctk.CTkFrame(sel_frame, fg_color="transparent")
        linha2.pack(fill="x", pady=4)

        list_frame = ctk.CTkFrame(linha2, fg_color="white",
                                    border_width=1, border_color=GRID_CINZA)
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 8))

        self.listbox_verbas_av = tk.Listbox(list_frame, selectmode="extended",
                                              height=8, exportselection=False,
                                              font=("Roboto", 10),
                                              bg="white", fg=TEXTO_PRINCIPAL,
                                              selectbackground=IG_AZUL,
                                              selectforeground="white",
                                              borderwidth=0, highlightthickness=0,
                                              activestyle="none")
        sb_lb = ttk.Scrollbar(list_frame, orient="vertical",
                                command=self.listbox_verbas_av.yview)
        self.listbox_verbas_av.configure(yscrollcommand=sb_lb.set)
        self.listbox_verbas_av.grid(row=0, column=0, sticky="nsew")
        sb_lb.grid(row=0, column=1, sticky="ns")
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.listbox_verbas_av.bind("<<ListboxSelect>>",
                                      lambda _: self._render_drill_verba())

        # Painel lateral: contador
        info_frame = ctk.CTkFrame(linha2, fg_color="#F7F9FB",
                                    border_width=1, border_color=GRID_CINZA, width=220)
        info_frame.pack(side="left", fill="y")
        info_frame.pack_propagate(False)
        ctk.CTkLabel(info_frame, text="Selecao atual",
                     font=("Rufina", 12, "bold"),
                     text_color=IG_AZUL_ESC).pack(anchor="w", padx=10, pady=(8, 4))
        self.lbl_contador_verbas_av = SelectableText(info_frame, text="(nenhuma)",
                                                       height=8, font=("Roboto", 9),
                                                       fg=TEXTO_PRINCIPAL, bg="#F7F9FB")
        self.lbl_contador_verbas_av.pack(fill="both", expand=True, padx=10, pady=4)

        # Cards de metricas da verba selecionada
        cards_av = ctk.CTkFrame(self.tab_av, fg_color="transparent")
        cards_av.pack(fill="x", padx=8, pady=8)
        for i in range(4):
            cards_av.grid_columnconfigure(i, weight=1, uniform="av")
        self.card_av_hc = MetricCard(cards_av); self.card_av_hc.grid(row=0, column=0, sticky="nsew", padx=4)
        self.card_av_total = MetricCard(cards_av); self.card_av_total.grid(row=0, column=1, sticky="nsew", padx=4)
        self.card_av_media = MetricCard(cards_av); self.card_av_media.grid(row=0, column=2, sticky="nsew", padx=4)
        self.card_av_status = MetricCard(cards_av); self.card_av_status.grid(row=0, column=3, sticky="nsew", padx=4)

        # Grafico HC + valor medio mensal da verba
        self.chart_av_hc = ChartContainer(self.tab_av, height=440)
        self.chart_av_hc.pack(fill="x", padx=8, pady=8)

        # ---- Bloco 3: colaborador por clique no grafico ----
        ctk.CTkLabel(self.tab_av, text="Colaboradores da verba selecionada",
                     font=("Rufina", 14, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w", padx=12, pady=(8, 4))
        ctk.CTkLabel(self.tab_av,
                     text="Clique em uma barra ou ponto do grafico acima para listar as matriculas do mes. "
                          "Depois selecione uma linha para abrir o detalhe ou navegar para Outliers.",
                     font=("Roboto", 10), text_color=TEXTO_SECUNDARIO,
                     wraplength=1100, justify="left").pack(anchor="w", padx=12)

        painel_colab = ctk.CTkFrame(self.tab_av, fg_color="transparent")
        painel_colab.pack(fill="x", padx=8, pady=6)

        topo_colab = ctk.CTkFrame(painel_colab, fg_color="transparent")
        topo_colab.pack(fill="x", pady=(0, 4))

        self.lbl_av_colab_status = SelectableText(topo_colab, text="(selecione uma verba para carregar os colaboradores)",
                                                   height=2, font=("Roboto", 10),
                                                   fg=TEXTO_SECUNDARIO, bg=FUNDO)
        self.lbl_av_colab_status.pack(side="left", fill="x", expand=True, padx=(4, 8))

        self.btn_av_tabela = ctk.CTkButton(topo_colab, text="Abrir na Tabela & Drill-down",
                                            command=self._ir_para_tabela_drill_colaborador_av,
                                            fg_color="white", text_color=IG_AZUL_ESC,
                                            hover_color="#E8F0F7", border_width=1, border_color=IG_AZUL,
                                            width=190)
        self.btn_av_tabela.pack(side="right", padx=4)

        self.btn_av_outliers = ctk.CTkButton(topo_colab, text="Ir para Outliers",
                                              command=self._ir_para_outliers_colaborador_av,
                                              fg_color=IG_AZUL, hover_color=IG_AZUL_ESC,
                                              width=130)
        self.btn_av_outliers.pack(side="right", padx=4)

        frame_tree_av = ctk.CTkFrame(painel_colab, fg_color="white", border_width=1, border_color=GRID_CINZA)
        frame_tree_av.pack(fill="both", expand=True)

        self.cols_av_colab = ("Mat", "Nome", "Status", "ValorMes", "MediaBaseVerba",
                               "VarAbsVerba", "VarPctVerba", "LiquidoAlvo", "Auditoria")
        self.tree_av_colab = ttk.Treeview(frame_tree_av, columns=self.cols_av_colab, show="headings", height=12)
        config_cols_av = [
            ("Mat", "Mat.", 75, "center"),
            ("Nome", "Nome", 260, "w"),
            ("Status", "Status liquido", 110, "center"),
            ("ValorMes", "Valor mes", 115, "e"),
            ("MediaBaseVerba", "Media verba", 115, "e"),
            ("VarAbsVerba", "Var. R$", 115, "e"),
            ("VarPctVerba", "Var. %", 80, "e"),
            ("LiquidoAlvo", "Liquido alvo", 115, "e"),
            ("Auditoria", "Auditoria", 360, "w"),
        ]
        for col, titulo, largura, anchor in config_cols_av:
            self.tree_av_colab.heading(col, text=titulo)
            self.tree_av_colab.column(col, width=largura, anchor=anchor)

        sy_av = ttk.Scrollbar(frame_tree_av, orient="vertical", command=self.tree_av_colab.yview)
        sx_av = ttk.Scrollbar(frame_tree_av, orient="horizontal", command=self.tree_av_colab.xview)
        self.tree_av_colab.configure(yscrollcommand=sy_av.set, xscrollcommand=sx_av.set)
        self.tree_av_colab.grid(row=0, column=0, sticky="nsew")
        sy_av.grid(row=0, column=1, sticky="ns")
        sx_av.grid(row=1, column=0, sticky="ew")
        frame_tree_av.grid_rowconfigure(0, weight=1)
        frame_tree_av.grid_columnconfigure(0, weight=1)
        self.tree_av_colab.tag_configure("critico", background="#FFE8EE")
        self.tree_av_colab.tag_configure("desligado", background="#F0F0F0", foreground=COR_DESLIGADO)
        self.tree_av_colab.bind("<<TreeviewSelect>>", self._on_av_colab_select)
        self.tree_av_colab.bind("<Double-1>", lambda _e: self._ir_para_tabela_drill_colaborador_av())
        bind_treeview_copy(self.tree_av_colab)

        self.lbl_av_aviso = SelectableText(self.tab_av, text="",
                                             height=2, font=("Roboto", 10),
                                             fg=TEXTO_SECUNDARIO, bg=FUNDO)
        self.lbl_av_aviso.pack(anchor="w", padx=12, pady=(0, 8), fill="x")

    def _build_tab_exportar(self):
        tab = self.tabview.tab("Exportar")
        cont = ctk.CTkFrame(tab, fg_color=FUNDO)
        cont.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(cont, text="Exportar relatorio Excel completo",
                     font=("Rufina", 16, "bold"), text_color=IG_AZUL_ESC).pack(anchor="w", pady=(8, 12))
        ctk.CTkLabel(cont,
                     text="Gera um arquivo .xlsx com 4 abas:\n"
                          "  - Resumo_Mensal: PGTO, DESC, Liquido e Verba 9950 por mes\n"
                          "  - Liquido_Funcionario: cada matricula com baseline, liquido alvo, var, status\n"
                          "  - Verbas_PGTO_Zeradas: verbas pagas que zeraram\n"
                          "  - Verbas_DESC_Zeradas: descontos que zeraram",
                     font=("Roboto", 11), text_color=TEXTO_PRINCIPAL,
                     justify="left").pack(anchor="w", pady=8)

        self.btn_exportar = ctk.CTkButton(cont, text="Salvar relatorio Excel...",
                                           command=self.acao_exportar_excel,
                                           fg_color=IG_AZUL, hover_color=IG_AZUL_ESC,
                                           font=("Roboto", 12, "bold"), height=44, width=280,
                                           state="disabled")
        self.btn_exportar.pack(anchor="w", pady=12)

        self.lbl_exportar_status = ctk.CTkLabel(cont, text="", font=("Roboto", 10),
                                                 text_color=TEXTO_SECUNDARIO)
        self.lbl_exportar_status.pack(anchor="w", pady=4)

    # ---------- Estado inicial ----------
    def _mostrar_estado_inicial(self):
        msg = ctk.CTkLabel(self.tab_vg,
                            text="Carregue o arquivo CSV no painel lateral para iniciar a auditoria.",
                            font=("Roboto", 12), text_color=TEXTO_SECUNDARIO)
        msg.pack(pady=40)
        self._estado_inicial_label = msg

    # ---------- Acoes ----------
    def acao_carregar_csv(self):
        path = filedialog.askopenfilename(
            title="Selecione o CSV (Confere - Codigos por Periodo)",
            filetypes=[("CSV files", "*.csv"), ("Todos", "*.*")]
        )
        if not path:
            return
        self.lbl_arquivo.configure(text=os.path.basename(path))
        try:
            df, meses = carregar_csv(path)
        except Exception as e:
            messagebox.showerror("Erro ao ler CSV",
                                  f"Falha ao ler o arquivo. Confirme se e o CSV no padrao esperado.\n\n{e}")
            return

        self.df = df
        self.meses_detectados = meses

        # Empresas
        self.empresas_disponiveis = sorted([e for e in df["Empresa"].dropna().astype(str).unique() if e.strip()])
        self.empresas_selecionadas = list(self.empresas_disponiveis)
        self._render_filtro_empresa()

        # Processos
        self.processos_disponiveis = sorted([p for p in df["Processo"].dropna().astype(str).unique() if p.strip()])
        self.processos_selecionados = list(self.processos_disponiveis)
        self._render_filtro_processo()

        # Mes alvo
        self.combo_mes_alvo.configure(values=meses)
        self.combo_mes_alvo.set(meses[-1])
        self.mes_alvo = meses[-1]

        # Baseline
        self._render_baseline_checkboxes()

        self.btn_reprocessar.configure(state="normal")
        self.btn_exportar.configure(state="normal")
        self.btn_filtro_verbas.configure(state="normal")

        if hasattr(self, "_estado_inicial_label"):
            try: self._estado_inicial_label.destroy()
            except: pass

        self.acao_reprocessar()

    def _render_filtro_empresa(self):
        for w in self.frame_empresa.winfo_children():
            w.destroy()
        if len(self.empresas_disponiveis) <= 1:
            ctk.CTkLabel(self.frame_empresa,
                         text=f"Empresa unica: {self.empresas_disponiveis[0] if self.empresas_disponiveis else '-'}",
                         font=("Roboto", 10), text_color=TEXTO_SECUNDARIO).pack(anchor="w")
            return
        ctk.CTkLabel(self.frame_empresa, text="Empresa(s)",
                     font=("Roboto", 11), text_color=TEXTO_PRINCIPAL).pack(anchor="w")
        sub = ctk.CTkScrollableFrame(self.frame_empresa, fg_color="white", height=80,
                                      border_width=1, border_color=GRID_CINZA)
        sub.pack(fill="x", pady=2)
        self.checks_empresa = {}
        for emp in self.empresas_disponiveis:
            var = tk.BooleanVar(value=True)
            cb = ctk.CTkCheckBox(sub, text=emp, variable=var, font=("Roboto", 10),
                                  fg_color=IG_AZUL, hover_color=IG_AZUL_ESC,
                                  command=self._on_empresa_changed)
            cb.pack(anchor="w", padx=4)
            self.checks_empresa[emp] = var

    def _render_filtro_processo(self):
        for w in self.frame_processo.winfo_children():
            w.destroy()
        if len(self.processos_disponiveis) <= 1:
            if self.processos_disponiveis:
                ctk.CTkLabel(self.frame_processo,
                             text=f"Processo unico: {self.processos_disponiveis[0]}",
                             font=("Roboto", 10), text_color=TEXTO_SECUNDARIO).pack(anchor="w")
            return
        ctk.CTkLabel(self.frame_processo, text="Processo (tipo de folha)",
                     font=("Roboto", 11), text_color=TEXTO_PRINCIPAL).pack(anchor="w")
        sub = ctk.CTkScrollableFrame(self.frame_processo, fg_color="white", height=80,
                                      border_width=1, border_color=GRID_CINZA)
        sub.pack(fill="x", pady=2)
        self.checks_processo = {}
        for p in self.processos_disponiveis:
            var = tk.BooleanVar(value=True)
            cb = ctk.CTkCheckBox(sub, text=p, variable=var, font=("Roboto", 10),
                                  fg_color=IG_AZUL, hover_color=IG_AZUL_ESC,
                                  command=self._on_processo_changed)
            cb.pack(anchor="w", padx=4)
            self.checks_processo[p] = var

    def _render_baseline_checkboxes(self):
        for w in self.frame_baseline.winfo_children():
            w.destroy()
        self.checks_baseline = {}
        meses_disp = [m for m in self.meses_detectados if m != self.mes_alvo]
        # Default: TODOS os meses disponiveis marcados (auditoria mais robusta)
        defaults_set = set(meses_disp)
        for m in self.meses_detectados:
            if m == self.mes_alvo:
                continue
            var = tk.BooleanVar(value=(m in defaults_set))
            cb = ctk.CTkCheckBox(self.frame_baseline, text=m, variable=var, font=("Roboto", 10),
                                  fg_color=IG_AZUL, hover_color=IG_AZUL_ESC,
                                  command=self._on_baseline_changed)
            cb.pack(anchor="w", padx=4)
            self.checks_baseline[m] = var
        self.meses_baseline = sorted([m for m, v in self.checks_baseline.items() if v.get()],
                                      key=lambda x: self.meses_detectados.index(x))

    def _on_empresa_changed(self):
        self.empresas_selecionadas = [e for e, v in self.checks_empresa.items() if v.get()]

    def _on_processo_changed(self):
        self.processos_selecionados = [p for p, v in self.checks_processo.items() if v.get()]

    def _on_baseline_changed(self):
        self.meses_baseline = sorted([m for m, v in self.checks_baseline.items() if v.get()],
                                      key=lambda x: self.meses_detectados.index(x))

    def _on_mes_alvo_changed(self):
        self.mes_alvo = self.combo_mes_alvo.get()
        if self.df is not None:
            self._render_baseline_checkboxes()

    def _on_excluir_toggle(self):
        self.excluir_desligados = bool(self.switch_excluir.get())

    def _on_metodo_changed(self, valor):
        self.config_param_nome = valor
        self.lbl_metodo_desc.configure(text=PARAMETROS_MERCADO[valor]["descricao"])

    def _on_mat_changed(self, valor):
        self.materialidade_pct_user = float(valor)
        self.lbl_mat_valor.configure(text=f"{valor:.1f}%")

    def acao_reprocessar(self):
        if self.df is None:
            return
        if not self.empresas_selecionadas:
            messagebox.showwarning("Filtros", "Selecione pelo menos uma empresa.")
            return
        if not self.processos_selecionados and self.processos_disponiveis:
            messagebox.showwarning("Filtros", "Selecione pelo menos um processo.")
            return
        if not self.meses_baseline:
            messagebox.showwarning("Filtros", "Selecione pelo menos um mes para a baseline.")
            return

        df_f = self.df[self.df["Empresa"].astype(str).isin(self.empresas_selecionadas)].copy()
        if self.processos_disponiveis:
            df_f = df_f[df_f["Processo"].astype(str).isin(self.processos_selecionados)].copy()
        if df_f.empty:
            messagebox.showerror("Filtros", "Nenhuma linha apos aplicar os filtros.")
            return

        # Aplica filtro de verbas excluidas (preserva verba 0020 sempre,
        # pois ela e usada para HC e nao entra no calculo do liquido)
        if self.verbas_excluidas:
            # Remove a 0020 do conjunto de exclusoes mesmo se o user marcou
            cods_norm_excl = {str(c).strip().lstrip("0") or "0" for c in self.verbas_excluidas}
            cods_norm_excl.discard("20")  # 0020 sem zeros a esquerda
            df_codigos_norm = df_f["Código"].astype(str).str.strip().str.lstrip("0").replace("", "0")
            mask_excluir = df_codigos_norm.isin(cods_norm_excl)
            df_f = df_f[~mask_excluir].copy()
            if df_f.empty:
                messagebox.showerror("Filtros", "Todas as linhas foram filtradas. Revise as verbas excluidas.")
                return

        # Atualiza label de status das verbas
        if self.verbas_excluidas:
            self.lbl_verbas_status.configure(
                text=f"{len(self.verbas_excluidas)} verba(s) excluida(s) dos calculos",
                text_color=IG_VINHO)
        else:
            self.lbl_verbas_status.configure(
                text="(todas as verbas incluidas)", text_color=TEXTO_SECUNDARIO)

        self.df_filtrado = df_f

        # Configura parametros
        config = dict(PARAMETROS_MERCADO[self.config_param_nome])
        config["materialidade_pct"] = self.materialidade_pct_user / 100.0

        desligados = detectar_desligados(df_f, self.mes_alvo)
        self.resumo = resumo_macro(df_f, self.meses_detectados)
        self.liquido, self.metadata_audit = liquido_por_funcionario(
            df_f, self.meses_detectados, self.mes_alvo, self.meses_baseline, desligados,
            config_param=config
        )
        self.salario_v20 = salario_verba20_por_mes(df_f, self.meses_detectados)
        self.zerados_pgto = verbas_zeradas(df_f, "PGTO", self.mes_alvo, self.meses_baseline)
        self.zerados_desc = verbas_zeradas(df_f, "DESC", self.mes_alvo, self.meses_baseline)
        self.liquido_analise = (self.liquido[self.liquido["STATUS"] != "DESLIGADO"]
                                  if self.excluir_desligados else self.liquido)

        # Stats sidebar
        n_func = df_f["Matrícula"].nunique()
        self.lbl_stats.configure(
            text=f"{len(df_f):,} linhas | {n_func} funcionarios | "
                 f"{len(self.empresas_selecionadas)} empresa(s) | "
                 f"{len(self.processos_selecionados) if self.processos_disponiveis else 1} processo(s)\n"
                 f"Periodo: {self.meses_detectados[0]} a {self.meses_detectados[-1]}\n"
                 f"Folha total {self.mes_alvo}: {fmt_brl(self.metadata_audit['folha_total_alvo'])}\n"
                 f"Materialidade: {fmt_brl(self.metadata_audit['materialidade'])}".replace(",", ".")
        )

        # Renderiza tudo
        self._render_visao_geral()
        self._render_headcount()
        self._render_outliers()
        self._render_verbas_zeradas()
        self._render_analise_verbas()
        self._popular_combos_tabela()
        self._renderizar_tabela()

        # Reset drill
        self.func_selecionado = None
        self.lbl_drill_func.configure(text="(selecione um funcionario na tabela acima)")
        self.chart_drill.limpar(); self.chart_drill.render()
        for it in self.tree_drill.get_children():
            self.tree_drill.delete(it)

    # ---------- Renderizacao - Visao Geral ----------
    def _render_visao_geral(self):
        linha_alvo = self.resumo[self.resumo["Mes"] == self.mes_alvo].iloc[0]
        linhas_baseline = self.resumo[self.resumo["Mes"].isin(self.meses_baseline)]
        media_pgto = linhas_baseline["PGTO"].mean() if len(linhas_baseline) else 0
        media_desc = linhas_baseline["DESC"].mean() if len(linhas_baseline) else 0
        media_liq = linhas_baseline["Liquido"].mean() if len(linhas_baseline) else 0

        var_pgto = (linha_alvo["PGTO"] / media_pgto - 1) * 100 if media_pgto else 0
        var_desc = (linha_alvo["DESC"] / media_desc - 1) * 100 if media_desc else 0
        var_liq = (linha_alvo["Liquido"] / media_liq - 1) * 100 if media_liq else 0
        diff_9950 = linha_alvo["Liquido"] - linha_alvo["Verba9950"]
        pct_9950 = (diff_9950 / linha_alvo["Verba9950"] * 100) if linha_alvo["Verba9950"] else 0
        classe_9950 = "" if abs(pct_9950) < 5 else ("atencao" if abs(pct_9950) < 15 else "alerta")

        self.card_pgto.atualizar(
            titulo=f"PGTO {self.mes_alvo}",
            valor=fmt_brl(linha_alvo["PGTO"]),
            delta=f"{fmt_pct(var_pgto)} vs media baseline",
            classe="negativo" if var_pgto < -10 else ("positivo" if var_pgto > 10 else "")
        )
        self.card_desc.atualizar(
            titulo=f"DESC {self.mes_alvo}",
            valor=fmt_brl(linha_alvo["DESC"]),
            delta=f"{fmt_pct(var_desc)} vs media baseline",
            classe="negativo" if var_desc < -10 else ("positivo" if var_desc > 10 else "")
        )
        self.card_liq.atualizar(
            titulo=f"Liquido {self.mes_alvo}",
            valor=fmt_brl(linha_alvo["Liquido"]),
            delta=f"{fmt_pct(var_liq)} vs media baseline",
            classe="negativo" if var_liq < -10 else ("positivo" if var_liq > 10 else "")
        )
        self.card_9950.atualizar(
            titulo="Verba 9950 (Liquido Mensal)",
            valor=fmt_brl(linha_alvo["Verba9950"]),
            delta=f"Diff: {fmt_brl(diff_9950)} ({fmt_pct(pct_9950)})",
            classe=classe_9950
        )

        # Aviso desligados
        n_desl = (self.liquido["STATUS"] == "DESLIGADO").sum()
        if n_desl > 0:
            estado = "excluidos da analise" if self.excluir_desligados else "incluidos na analise"
            self.lbl_aviso_desligados.configure(
                text=f"Desligados detectados: {n_desl} funcionarios | atualmente {estado}.")
        else:
            self.lbl_aviso_desligados.configure(text="")

        # ---- Grafico evolucao ----
        self.chart_evolucao.limpar()
        ax1 = self.chart_evolucao.fig.add_subplot(111)
        x = np.arange(len(self.resumo["Mes"]))
        w = 0.35
        ax1.bar(x - w/2, self.resumo["PGTO"], w, color=IG_AZUL, label="PGTO",
                edgecolor="white", linewidth=0.5)
        ax1.bar(x + w/2, self.resumo["DESC"], w, color=IG_VINHO, label="DESC",
                edgecolor="white", linewidth=0.5)
        ax1.set_xticks(x)
        ax1.set_xticklabels(self.resumo["Mes"], rotation=45, ha="right", fontsize=9)
        ax1.set_ylabel("PGTO / DESC (R$)", color="#1A1A1A", fontsize=10, fontweight="medium")
        ax1.yaxis.set_major_formatter(FuncFormatter(fmt_eixo_brl))
        estilo_eixos(ax1, "Evolucao mensal")

        ax2 = ax1.twinx()
        ax2.plot(x, self.resumo["Liquido"], color=IG_AZUL_ESC, marker="o", linewidth=2.4,
                  markersize=7, markeredgecolor="white", markeredgewidth=1, label="Liquido")
        # Rotulos sobre cada ponto da linha de liquido
        for xi, yi in zip(x, self.resumo["Liquido"]):
            ax2.text(xi, yi, f" {fmt_eixo_brl(yi, None)}",
                      fontsize=8.5, fontweight="bold", color=IG_AZUL_ESC,
                      ha="left", va="bottom")
        v9950 = self.resumo[self.resumo["Verba9950"] > 0]
        if not v9950.empty:
            v9950_x = [list(self.resumo["Mes"]).index(m) for m in v9950["Mes"]]
            ax2.scatter(v9950_x, v9950["Verba9950"], color=IG_LARANJA, marker="D", s=85,
                        edgecolors="white", linewidths=1.5, label="Verba 9950", zorder=5)
        ax2.set_ylabel("Liquido (R$)", color=IG_AZUL_ESC, fontsize=10, fontweight="medium")
        ax2.yaxis.set_major_formatter(FuncFormatter(fmt_eixo_brl))
        ax2.spines["top"].set_visible(False)
        ax2.tick_params(colors="#1A1A1A", labelsize=9)

        # Highlight mes alvo
        idx_alvo = list(self.resumo["Mes"]).index(self.mes_alvo)
        ax1.axvspan(idx_alvo - 0.5, idx_alvo + 0.5, color=IG_AZUL_CLR, alpha=0.20)
        ax1.text(idx_alvo, ax1.get_ylim()[1] * 0.98, "Folha calculada",
                 ha="center", fontsize=9, color=IG_AZUL_ESC, fontweight="bold")

        # Legenda combinada
        h1, l1 = ax1.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        ax1.legend(h1 + h2, l1 + l2, loc="upper left", fontsize=9, frameon=True,
                   facecolor="white", edgecolor=GRID_CINZA)

        self.chart_evolucao.fig.tight_layout()
        self.chart_evolucao.render()

        # ---- Conciliacao ----
        self.chart_conciliacao.limpar()
        ax = self.chart_conciliacao.fig.add_subplot(111)
        rotulos = ["PGTO", "DESC", "Liquido", "Verba 9950"]
        valores = [linha_alvo["PGTO"], -linha_alvo["DESC"], linha_alvo["Liquido"], linha_alvo["Verba9950"]]
        cores = [IG_AZUL, IG_VINHO, IG_AZUL_ESC, IG_LARANJA]
        bars = ax.bar(rotulos, valores, color=cores, edgecolor="white", linewidth=1.2)
        # Margem para os rotulos nao serem cortados
        ymax = max(valores); ymin = min(valores)
        margem = max(abs(ymax), abs(ymin)) * 0.15
        ax.set_ylim(ymin - margem if ymin < 0 else -margem*0.05, ymax + margem)
        for bar, v in zip(bars, valores):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
                    fmt_brl(v),
                    ha="center",
                    va="bottom" if v >= 0 else "top",
                    fontsize=10, fontweight="bold", color="#1A1A1A")
        ax.yaxis.set_major_formatter(FuncFormatter(fmt_eixo_brl))
        ax.axhline(0, color="#444444", linewidth=0.8)
        estilo_eixos(ax, f"Conciliacao em {self.mes_alvo}")
        self.chart_conciliacao.fig.tight_layout()
        self.chart_conciliacao.render()

        # ---- Distribuicao de status ----
        self.chart_status.limpar()
        ax = self.chart_status.fig.add_subplot(111)
        contagem = self.liquido_analise["STATUS"].value_counts()
        cores_st = [cor_status(s) for s in contagem.index]
        bars = ax.barh(contagem.index, contagem.values, color=cores_st,
                        edgecolor="white", linewidth=0.8)
        max_v = max(contagem.values) if len(contagem) else 1
        ax.set_xlim(0, max_v * 1.15)
        for bar, v in zip(bars, contagem.values):
            ax.text(bar.get_width() + max_v * 0.012, bar.get_y() + bar.get_height()/2,
                    fmt_int(v), va="center", ha="left",
                    fontsize=10, fontweight="bold", color="#1A1A1A")
        ax.invert_yaxis()
        estilo_eixos(ax, "Distribuicao de status")
        ax.set_xlabel("Funcionarios", fontsize=10, color="#1A1A1A", fontweight="medium")
        self.chart_status.fig.tight_layout()
        self.chart_status.render()

    # ---------- Renderizacao - Headcount ----------
    def _render_headcount(self):
        linha = self.salario_v20[self.salario_v20["Mes"] == self.mes_alvo].iloc[0]
        base = self.salario_v20[self.salario_v20["Mes"].isin(self.meses_baseline)]
        hc_med = base["HC"].mean() if not base.empty else 0
        sal_med = base["Salario_Medio"].mean() if not base.empty else 0

        var_hc = (linha["HC"] / hc_med - 1) * 100 if hc_med else 0
        var_sal = (linha["Salario_Medio"] / sal_med - 1) * 100 if sal_med else 0
        cl_hc = "" if abs(var_hc) < 3 else ("atencao" if abs(var_hc) < 8 else "alerta")
        cl_sal = "" if abs(var_sal) < 3 else ("atencao" if abs(var_sal) < 8 else "alerta")

        self.card_hc.atualizar(titulo=f"HC verba 0020 em {self.mes_alvo}",
                                valor=fmt_int(linha["HC"]),
                                delta=f"{fmt_pct(var_hc)} vs media baseline",
                                classe=cl_hc + (" negativo" if var_hc < 0 else ""))
        self.card_sal.atualizar(titulo=f"Salario medio em {self.mes_alvo}",
                                 valor=fmt_brl(linha["Salario_Medio"]),
                                 delta=f"{fmt_pct(var_sal)} vs media baseline",
                                 classe=cl_sal + (" negativo" if var_sal < 0 else ""))
        self.card_total20.atualizar(titulo=f"Total verba 0020 em {self.mes_alvo}",
                                     valor=fmt_brl(linha["Salario_Total"]),
                                     delta="")

        # Grafico HC + salario medio
        self.chart_hc.limpar()
        ax1 = self.chart_hc.fig.add_subplot(111)
        x = np.arange(len(self.salario_v20["Mes"]))
        bars = ax1.bar(x, self.salario_v20["HC"], color=IG_AZUL_CLR,
                        edgecolor=IG_AZUL_ESC, linewidth=0.6)
        # Margem para rotulos no topo das barras
        max_hc = self.salario_v20["HC"].max() if len(self.salario_v20) else 1
        ax1.set_ylim(0, max_hc * 1.15)
        for bar, v in zip(bars, self.salario_v20["HC"]):
            ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max_hc * 0.015,
                     fmt_int(v), ha="center", va="bottom",
                     fontsize=9, fontweight="bold", color=IG_AZUL_ESC)
        ax1.set_xticks(x)
        ax1.set_xticklabels(self.salario_v20["Mes"], rotation=45, ha="right", fontsize=9)
        ax1.set_ylabel("HC (quantidade)", color="#1A1A1A", fontsize=10, fontweight="medium")
        if hc_med:
            ax1.axhline(hc_med, color=IG_VERDE_ESC, linestyle=":", linewidth=1.5)
            ax1.text(0, hc_med, f" HC medio baseline: {int(hc_med)}", color=IG_VERDE_ESC,
                     fontsize=9, fontweight="bold", va="bottom")
        estilo_eixos(ax1, "Headcount e salario medio mensal")

        ax2 = ax1.twinx()
        ax2.plot(x, self.salario_v20["Salario_Medio"], color=IG_AZUL_ESC,
                 marker="o", linewidth=2.4, markersize=8, markeredgecolor="white", markeredgewidth=1)
        # Rotulos em cada ponto do salario medio
        for xi, yi in zip(x, self.salario_v20["Salario_Medio"]):
            ax2.text(xi, yi, f" {fmt_eixo_brl(yi, None)}",
                      fontsize=8.5, fontweight="bold", color=IG_AZUL_ESC,
                      ha="left", va="bottom")
        ax2.set_ylabel("Salario medio (R$)", color=IG_AZUL_ESC, fontsize=10, fontweight="medium")
        ax2.yaxis.set_major_formatter(FuncFormatter(fmt_eixo_brl))
        ax2.spines["top"].set_visible(False)
        ax2.tick_params(colors="#1A1A1A", labelsize=9)

        idx_alvo = list(self.salario_v20["Mes"]).index(self.mes_alvo)
        ax1.axvspan(idx_alvo - 0.5, idx_alvo + 0.5, color=IG_AZUL_CLR, alpha=0.20)

        self.chart_hc.fig.tight_layout()
        self.chart_hc.render()

    # ---------- Renderizacao - Outliers ----------
    def _render_outliers(self):
        # ===== DISPERSAO =====
        self.chart_dispersao.limpar()
        # Layout em 2 colunas: grafico + caixa de legenda dos top criticos
        self.chart_dispersao.fig.set_size_inches(11, 6.5)
        gs = self.chart_dispersao.fig.add_gridspec(1, 2, width_ratios=[3.2, 1.0],
                                                     left=0.07, right=0.98, top=0.92, bottom=0.10,
                                                     wspace=0.05)
        ax = self.chart_dispersao.fig.add_subplot(gs[0, 0])
        ax_leg = self.chart_dispersao.fig.add_subplot(gs[0, 1])
        ax_leg.set_axis_off()

        # Filtra baseline minima para nao esmagar a escala log
        BASE_MIN_DISP = 100.0
        valido = self.liquido_analise[self.liquido_analise["MEDIA_BASELINE"] >= BASE_MIN_DISP].copy()
        n_filtrados = (self.liquido_analise["MEDIA_BASELINE"] > 0).sum() - len(valido)

        if len(valido):
            valido["LIQ_PLOT"] = valido["LIQUIDO_ALVO"].clip(lower=100)
            valido["EH_CRITICO"] = valido["STATUS"].isin(list(STATUS_CRITICOS))
            lim_max = max(valido["MEDIA_BASELINE"].max(), valido["LIQUIDO_ALVO"].max()) * 1.5
            xs = np.logspace(np.log10(BASE_MIN_DISP), np.log10(lim_max), 80)
            ax.fill_between(xs, xs * 0.7, xs * 1.3, color=IG_AZUL_CLR, alpha=0.20, label="Banda +/-30%")
            ax.plot(xs, xs, linestyle="--", color="#444444", linewidth=1.2, label="y=x")

            # Plota nao-criticos primeiro, criticos depois (sobrepostos)
            nc = valido[~valido["EH_CRITICO"]]
            cr = valido[valido["EH_CRITICO"]]
            if len(nc):
                ax.scatter(nc["MEDIA_BASELINE"], nc["LIQ_PLOT"],
                            c=nc["STATUS"].map(cor_status).fillna(GRID_CINZA),
                            s=38, edgecolors="white", linewidths=0.8, alpha=0.85, zorder=2)
            if len(cr):
                ax.scatter(cr["MEDIA_BASELINE"], cr["LIQ_PLOT"],
                            c=cr["STATUS"].map(cor_status).fillna(IG_VERMELHO),
                            s=110, edgecolors="white", linewidths=1.2, alpha=0.95, zorder=3,
                            label="Criticos")

                # Top 5 criticos por impacto absoluto: numera no plot + legenda lateral
                cr_top = cr.copy()
                cr_top["IMPACTO_ABS"] = cr_top["VAR_ABS"].abs()
                cr_top = cr_top.sort_values("IMPACTO_ABS", ascending=False).head(5).reset_index()

                for i, row in cr_top.iterrows():
                    n = i + 1
                    # Numero pequeno em circulo branco com borda vermelha sobre o ponto
                    ax.annotate(
                        f"{n}",
                        xy=(row["MEDIA_BASELINE"], min(row["LIQUIDO_ALVO"] if row["LIQUIDO_ALVO"] > 0 else 100, lim_max)),
                        xytext=(0, 0), textcoords="offset points",
                        ha="center", va="center",
                        fontsize=10, fontweight="bold", color=IG_VERMELHO,
                        bbox=dict(boxstyle="circle,pad=0.25", fc="white",
                                   ec=IG_VERMELHO, lw=1.5),
                        zorder=10,
                    )

                # Caixa de legenda lateral
                ax_leg.text(0.02, 0.98, "Top 5 outliers",
                              transform=ax_leg.transAxes,
                              fontsize=11, fontweight="bold", color=IG_AZUL_ESC,
                              ha="left", va="top")
                y_pos = 0.92
                for i, row in cr_top.iterrows():
                    n = i + 1
                    mat = row["Matrícula"]
                    nome_curto = str(row["Nome"])[:24]
                    var_pct = row["VAR_PCT"] if pd.notna(row["VAR_PCT"]) else 0
                    impacto = row["IMPACTO_ABS"]
                    status = row["STATUS"]

                    # Bolinha numerada
                    ax_leg.text(0.05, y_pos, f"{n}",
                                  transform=ax_leg.transAxes,
                                  fontsize=10, fontweight="bold", color=IG_VERMELHO,
                                  ha="center", va="top",
                                  bbox=dict(boxstyle="circle,pad=0.25", fc="white",
                                             ec=IG_VERMELHO, lw=1.3))
                    ax_leg.text(0.13, y_pos, f"{mat} - {nome_curto}",
                                  transform=ax_leg.transAxes,
                                  fontsize=9, fontweight="bold", color="#1A1A1A",
                                  ha="left", va="top")
                    y_pos -= 0.045
                    ax_leg.text(0.13, y_pos, f"{status}  |  {fmt_brl(impacto)}  ({fmt_pct(var_pct)})",
                                  transform=ax_leg.transAxes,
                                  fontsize=8, color=TEXTO_SECUNDARIO,
                                  ha="left", va="top")
                    y_pos -= 0.085

                # Aviso de filtragem se aplicavel
                if n_filtrados > 0:
                    ax_leg.text(0.02, 0.02,
                                  f"({n_filtrados} func. omitidos:\nbaseline < R$ 100)",
                                  transform=ax_leg.transAxes,
                                  fontsize=8, color=TEXTO_SECUNDARIO,
                                  ha="left", va="bottom", style="italic")

            ax.set_xscale("log"); ax.set_yscale("log")
            ax.set_xlim(BASE_MIN_DISP * 0.8, lim_max)
            ax.set_xlabel("Media liquido baseline (R$) - log", fontsize=10,
                          color="#1A1A1A", fontweight="medium")
            ax.set_ylabel(f"Liquido {self.mes_alvo} (R$) - log", fontsize=10,
                          color="#1A1A1A", fontweight="medium")
            ax.legend(loc="lower right", fontsize=9, frameon=True, facecolor="white",
                       edgecolor=GRID_CINZA)

            # Destaque opcional quando o usuario veio da aba Analise de Verbas
            self._destacar_outlier_foco(ax, ax_leg, valido, lim_max)

        estilo_eixos(ax, "Dispersao Baseline vs Folha calculada")
        self.chart_dispersao.render()

        # Top outliers
        self.chart_top.limpar()
        ax = self.chart_top.fig.add_subplot(111)
        criticos = self.liquido_analise[self.liquido_analise["STATUS"].isin(list(STATUS_CRITICOS))].copy()
        if not criticos.empty:
            criticos["IMPACTO_ABS"] = criticos["VAR_ABS"].abs()
            top = criticos.sort_values("IMPACTO_ABS", ascending=False).head(15)
            rotulos = [f"{idx[0]} - {idx[1][:30]}" for idx in top.index]
            bars = ax.barh(rotulos, top["VAR_ABS"], color=IG_VERMELHO, edgecolor="white", linewidth=0.8)
            # Margem a direita para os rotulos (texto SEMPRE no lado positivo)
            x_max_pos = max(top["VAR_ABS"].max(), 0)
            x_min_neg = min(top["VAR_ABS"].min(), 0)
            margem_dir = max(abs(x_max_pos), abs(x_min_neg)) * 0.45
            ax.set_xlim(x_min_neg * 1.05 if x_min_neg < 0 else -margem_dir*0.05,
                          x_max_pos + margem_dir)
            # Estrategia: texto SEMPRE no lado positivo, alinhado a esquerda
            # - Barra positiva: texto comeca no fim da barra (a direita)
            # - Barra negativa: texto comeca a direita do zero (nao colide com label do eixo Y)
            offset_zero = margem_dir * 0.04
            for bar, v, st in zip(bars, top["VAR_ABS"], top["STATUS"]):
                if v >= 0:
                    x_text = v + offset_zero
                else:
                    x_text = offset_zero  # logo a direita do eixo zero
                ax.text(x_text, bar.get_y() + bar.get_height()/2,
                        f"{fmt_brl(v)}  [{st}]",
                        va="center", ha="left",
                        fontsize=9, fontweight="bold",
                        color="#1A1A1A")
            ax.invert_yaxis()
            ax.xaxis.set_major_formatter(FuncFormatter(fmt_eixo_brl))
            ax.set_xlabel(f"Variacao ({self.mes_alvo} - baseline)", fontsize=10,
                          color="#1A1A1A", fontweight="medium")
            ax.axvline(0, color="#444444", linewidth=0.8)
            estilo_eixos(ax, "Top 15 outliers - Maior impacto em reais")
        else:
            ax.text(0.5, 0.5, f"Nenhum outlier critico em {self.mes_alvo}.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=12, color=IG_VERDE_ESC, fontweight="bold")
            ax.set_axis_off()
        self.chart_top.fig.tight_layout()
        self.chart_top.render()

    # ---------- Renderizacao - Verbas Zeradas ----------
    def _render_verbas_zeradas(self):
        def _plot(chart, dados, cor, titulo):
            chart.limpar()
            ax = chart.fig.add_subplot(111)
            if dados.empty:
                ax.text(0.5, 0.5, f"Nenhuma verba regular zerou em {self.mes_alvo}.",
                        transform=ax.transAxes, ha="center", va="center",
                        fontsize=12, color=IG_VERDE_ESC, fontweight="bold")
                ax.set_axis_off()
            else:
                top = dados.head(15)
                rotulos = [f"{idx[0]} - {idx[1][:25]}" for idx in top.index]
                bars = ax.barh(rotulos, top["MEDIA_BASELINE"], color=cor,
                                edgecolor="white", linewidth=0.8)
                # Margem para rotulos
                max_v = top["MEDIA_BASELINE"].max()
                ax.set_xlim(0, max_v * 1.30)
                for bar, v in zip(bars, top["MEDIA_BASELINE"]):
                    ax.text(bar.get_width() + max_v * 0.015, bar.get_y() + bar.get_height()/2,
                            fmt_brl(v), va="center", ha="left",
                            fontsize=9, fontweight="bold", color="#1A1A1A")
                ax.invert_yaxis()
                ax.xaxis.set_major_formatter(FuncFormatter(fmt_eixo_brl))
                ax.set_xlabel("Media baseline (R$)", fontsize=10,
                              color="#1A1A1A", fontweight="medium")
                estilo_eixos(ax, titulo)
            chart.fig.tight_layout()
            chart.render()

        _plot(self.chart_zerados_pgto, self.zerados_pgto, IG_AZUL,
              f"PGTO regulares zeradas em {self.mes_alvo}")
        _plot(self.chart_zerados_desc, self.zerados_desc, IG_VINHO,
              f"DESC regulares zeradas em {self.mes_alvo}")

    # ---------- Tabela detalhada ----------
    def _popular_combos_tabela(self):
        statuses = ["(todos)"] + sorted(self.liquido_analise["STATUS"].unique().tolist())
        self.combo_status_tab.configure(values=statuses)
        self.combo_status_tab.set("(todos)")

    def _limpar_filtros_tabela(self):
        self.combo_status_tab.set("(todos)")
        self.combo_ordem_tab.set("VAR_ABS (impacto)")
        self.entry_busca.delete(0, "end")
        self._renderizar_tabela()

    def _renderizar_tabela(self):
        if self.liquido_analise is None:
            return
        for it in self.tree.get_children():
            self.tree.delete(it)

        tab = self.liquido_analise.reset_index().copy()
        st_filt = self.combo_status_tab.get()
        if st_filt and st_filt != "(todos)":
            tab = tab[tab["STATUS"] == st_filt]

        busca = self.entry_busca.get().strip().upper()
        if busca:
            tab = tab[tab["Nome"].str.upper().str.contains(busca, na=False) |
                       tab["Matrícula"].astype(str).str.contains(busca, na=False)]

        ordem = self.combo_ordem_tab.get()
        ordem_map = {
            "VAR_ABS (impacto)": ("VAR_ABS", True),
            "VAR_PCT (variacao %)": ("VAR_PCT", True),
            "LIQUIDO_ALVO": ("LIQUIDO_ALVO", False),
            "MEDIA_BASELINE": ("MEDIA_BASELINE", False),
        }
        col_ord, ascendente = ordem_map.get(ordem, ("VAR_ABS", True))
        if "impacto" in ordem or "variacao" in ordem:
            tab = tab.reindex(tab[col_ord].abs().sort_values(ascending=False).index)
        else:
            tab = tab.sort_values(col_ord, ascending=ascendente)

        # Limita exibicao para nao travar (mantem todos no estado, exibe primeiros 1000)
        for _, row in tab.head(1000).iterrows():
            tags = ()
            if row["STATUS"] in STATUS_CRITICOS:
                tags = ("critico",)
            elif row["STATUS"] == "DESLIGADO":
                tags = ("desligado",)
            self.tree.insert("", "end", values=(
                row["Matrícula"],
                row["Nome"][:60],
                fmt_brl(row["MEDIA_BASELINE"]),
                fmt_brl(row["LIQUIDO_ALVO"]),
                fmt_brl(row["VAR_ABS"]),
                fmt_pct(row["VAR_PCT"]),
                f"{row['Z_SCORE']:+.2f}" if pd.notna(row["Z_SCORE"]) else "-",
                row["STATUS"],
                row["AUDITORIA"][:120],
            ), tags=tags)

        total = len(tab)
        exib = min(1000, total)
        self.lbl_count_tabela.configure(
            text=f"Linhas exibidas: {exib} de {total} (total no CSV: {len(self.liquido)})"
                 + (f" | Mostrando primeiras 1000" if total > 1000 else ""))

    # ---------- Drill-down ----------
    def _on_func_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        valores = self.tree.item(sel[0], "values")
        mat = valores[0]
        nome = valores[1]
        # Encontrar o nome completo
        match = self.liquido.reset_index()
        match = match[(match["Matrícula"] == mat) & (match["Nome"].str.startswith(nome[:30]))]
        if match.empty:
            return
        nome_full = match.iloc[0]["Nome"]
        self.func_selecionado = (mat, nome_full)
        self._render_drill()

    def _render_drill(self):
        if self.func_selecionado is None:
            return
        mat, nome = self.func_selecionado
        try:
            serie = self.liquido.loc[(mat, nome)]
        except KeyError:
            return

        self.lbl_drill_func.configure(
            text=f"Mat. {mat} - {nome} | Status: {serie['STATUS']} | "
                 f"Baseline: {fmt_brl(serie['MEDIA_BASELINE'])} | "
                 f"Liquido {self.mes_alvo}: {fmt_brl(serie['LIQUIDO_ALVO'])} ({fmt_pct(serie['VAR_PCT'])}) | "
                 f"{serie['AUDITORIA']}"
        )

        # Grafico
        self.chart_drill.limpar()
        ax = self.chart_drill.fig.add_subplot(111)
        x = np.arange(len(self.meses_detectados))
        valores = [serie[m] for m in self.meses_detectados]
        critico = serie["STATUS"] in STATUS_CRITICOS
        cor = IG_VERMELHO if critico else IG_AZUL
        ax.plot(x, valores, color=cor, linewidth=2.6, marker="o", markersize=9,
                markerfacecolor=cor, markeredgecolor="white", markeredgewidth=1.2)
        # Rotulos em cada ponto
        for xi, yi in zip(x, valores):
            if yi != 0 or xi == self.meses_detectados.index(self.mes_alvo):
                ax.text(xi, yi, f" {fmt_eixo_brl(yi, None)}",
                        fontsize=9, fontweight="bold", color="#1A1A1A",
                        ha="left", va="bottom")
        ax.set_xticks(x)
        ax.set_xticklabels(self.meses_detectados, rotation=45, ha="right", fontsize=9)
        if pd.notna(serie["MEDIA_BASELINE"]):
            ax.axhline(serie["MEDIA_BASELINE"], color=IG_VERDE_ESC, linestyle="--", linewidth=1.5)
            ax.text(0, serie["MEDIA_BASELINE"], f" Baseline: {fmt_brl(serie['MEDIA_BASELINE'])}",
                    color=IG_VERDE_ESC, fontsize=9, fontweight="bold", va="bottom")
        idx_alvo = self.meses_detectados.index(self.mes_alvo)
        ax.axvspan(idx_alvo - 0.5, idx_alvo + 0.5, color=IG_AZUL_CLR, alpha=0.22)
        ax.yaxis.set_major_formatter(FuncFormatter(fmt_eixo_brl))
        ax.set_ylabel("Liquido (R$)", fontsize=10, color="#1A1A1A", fontweight="medium")
        estilo_eixos(ax, f"Evolucao do liquido")
        self.chart_drill.fig.tight_layout()
        self.chart_drill.render()

        # Filtros do drill
        func_df = self.df_filtrado[(self.df_filtrado["Matrícula"] == mat) &
                                     (self.df_filtrado["Nome"] == nome)].copy()
        cr_col = "C.R." if "C.R." in func_df.columns else None
        crs = sorted([str(x) for x in func_df[cr_col].dropna().astype(str).unique()
                       if str(x).strip()]) if cr_col else []
        clss = sorted([str(x) for x in func_df["Clas."].dropna().astype(str).unique() if str(x).strip()])
        procs = sorted([str(x) for x in func_df["Processo"].dropna().astype(str).unique() if str(x).strip()])

        self.combo_dd_cr.configure(values=["(todos)"] + crs); self.combo_dd_cr.set("(todos)")
        self.combo_dd_classe.configure(values=["(todos)"] + clss); self.combo_dd_classe.set("(todos)")
        self.combo_dd_proc.configure(values=["(todos)"] + procs); self.combo_dd_proc.set("(todos)")
        self.entry_dd_verba.delete(0, "end")

        self._renderizar_drill_verbas()

    def _limpar_filtros_drill(self):
        self.combo_dd_cr.set("(todos)")
        self.combo_dd_classe.set("(todos)")
        self.combo_dd_proc.set("(todos)")
        self.entry_dd_verba.delete(0, "end")
        self._renderizar_drill_verbas()

    def _renderizar_drill_verbas(self):
        for it in self.tree_drill.get_children():
            self.tree_drill.delete(it)
        if self.func_selecionado is None:
            return
        mat, nome = self.func_selecionado
        func_df = self.df_filtrado[(self.df_filtrado["Matrícula"] == mat) &
                                     (self.df_filtrado["Nome"] == nome)].copy()

        col_alvo = f"{self.mes_alvo} - Valor"
        func_df = func_df[func_df[col_alvo] != 0].copy()

        idx_alvo = self.meses_detectados.index(self.mes_alvo)
        meses_3 = self.meses_detectados[max(0, idx_alvo - 3):idx_alvo]
        cols_3m = [f"{m} - Valor" for m in meses_3]

        # Atualiza cabecalhos das colunas com os meses reais
        nomes_meses = list(meses_3)
        while len(nomes_meses) < 3:
            nomes_meses.insert(0, "-")
        self.tree_drill.heading("M3", text=nomes_meses[0])
        self.tree_drill.heading("M2", text=nomes_meses[1])
        self.tree_drill.heading("M1", text=nomes_meses[2])
        self.tree_drill.heading("Alvo", text=self.mes_alvo)

        cr_col = "C.R." if "C.R." in func_df.columns else None
        # Filtros
        f_cr = self.combo_dd_cr.get()
        f_classe = self.combo_dd_classe.get()
        f_proc = self.combo_dd_proc.get()
        busca_v = self.entry_dd_verba.get().strip().upper()

        if cr_col and f_cr and f_cr != "(todos)":
            func_df = func_df[func_df[cr_col].astype(str) == f_cr]
        if f_classe and f_classe != "(todos)":
            func_df = func_df[func_df["Clas."].astype(str) == f_classe]
        if f_proc and f_proc != "(todos)":
            func_df = func_df[func_df["Processo"].astype(str) == f_proc]
        if busca_v:
            func_df = func_df[
                func_df["Código"].astype(str).str.upper().str.contains(busca_v, na=False) |
                func_df["Descrição"].astype(str).str.upper().str.contains(busca_v, na=False)
            ]

        ordem_clas = ["PGTO", "DESC", "OUTRO"]
        outros = [c for c in func_df["Clas."].astype(str).unique() if c not in ordem_clas]
        ordem_completa = ordem_clas + sorted(outros)
        func_df["Clas."] = pd.Categorical(func_df["Clas."], categories=ordem_completa, ordered=True)
        sort_cols = ["Clas.", cr_col, "Código"] if cr_col else ["Clas.", "Código"]
        func_df = func_df.sort_values(sort_cols).reset_index(drop=True)

        # Garante 3 colunas de meses passados (preenche com '-' se ausente)
        for m_col in cols_3m:
            if m_col not in func_df.columns:
                func_df[m_col] = 0

        for _, row in func_df.iterrows():
            valores_3m = [fmt_brl(row[c]) if c in row else "-" for c in cols_3m]
            # Padding caso meses_3 tenha menos que 3
            while len(valores_3m) < 3:
                valores_3m.insert(0, "-")
            self.tree_drill.insert("", "end", values=(
                row["Código"],
                str(row["Descrição"])[:50],
                str(row["Clas."]),
                str(row[cr_col]) if cr_col else "-",
                str(row["Processo"])[:25],
                valores_3m[0], valores_3m[1], valores_3m[2],
                fmt_brl(row[col_alvo]),
            ))

    # ---------- Renderizacao - Analise de Verbas ----------
    def _render_analise_verbas(self):
        # Para esta aba precisamos do df SEM o filtro de verbas (queremos ver TODAS,
        # inclusive as excluidas, para auditoria). Reconstroi a partir do df base.
        df_base = self.df[self.df["Empresa"].astype(str).isin(self.empresas_selecionadas)].copy()
        if self.processos_disponiveis:
            df_base = df_base[df_base["Processo"].astype(str).isin(self.processos_selecionados)].copy()
        if df_base.empty:
            return

        self.df_av_base = df_base  # guarda para o drill-down individual

        # ===== Top verbas por impacto =====
        top_verbas = impacto_por_verba(df_base, self.meses_detectados, self.mes_alvo,
                                         self.meses_baseline, top_n=20)
        self.top_verbas_df = top_verbas

        self.chart_top_verbas.limpar()
        self.av_top_verbas_bars = []
        ax = self.chart_top_verbas.fig.add_subplot(111)
        if not top_verbas.empty:
            rotulos = [f"{r['Código']} - {str(r['Descrição'])[:35]} [{r['Clas.']}]"
                       for _, r in top_verbas.iterrows()]
            valores = top_verbas["VAR_ABS"].fillna(top_verbas["VALOR_ALVO"]).values

            # Cores: verbas excluidas em cinza listrado, demais em vermelho
            cods_excluidos = set(self.verbas_excluidas)
            cores = []
            hatches = []
            for _, r in top_verbas.iterrows():
                cod = str(r["Código"]).strip()
                if cod in cods_excluidos:
                    cores.append("#999999")
                    hatches.append("//")
                else:
                    cores.append(IG_VERMELHO)
                    hatches.append("")

            bars = ax.barh(rotulos, valores, color=cores, edgecolor="white", linewidth=0.8)
            self.av_top_verbas_bars = [
                (bar, str(r["Código"]).strip())
                for bar, (_, r) in zip(bars, top_verbas.iterrows())
            ]
            for bar, h in zip(bars, hatches):
                if h:
                    bar.set_hatch(h)

            x_max_pos = max(max(valores), 0)
            x_min_neg = min(min(valores), 0)
            margem = max(abs(x_max_pos), abs(x_min_neg)) * 0.45
            ax.set_xlim(x_min_neg * 1.05 if x_min_neg < 0 else -margem*0.05,
                          x_max_pos + margem)
            offset_zero = margem * 0.04
            for bar, v, (_, r) in zip(bars, valores, top_verbas.iterrows()):
                cod = str(r["Código"]).strip()
                marca = "  [EXCLUIDA]" if cod in cods_excluidos else ""
                pct = r["VAR_PCT"] if pd.notna(r["VAR_PCT"]) else None
                pct_txt = f"  ({fmt_pct(pct)})" if pct is not None else ""
                if v >= 0:
                    x_text = v + offset_zero
                else:
                    x_text = offset_zero
                ax.text(x_text, bar.get_y() + bar.get_height()/2,
                        f"{fmt_brl(v)}{pct_txt}{marca}",
                        va="center", ha="left",
                        fontsize=9, fontweight="bold",
                        color="#1A1A1A")
            ax.invert_yaxis()
            ax.xaxis.set_major_formatter(FuncFormatter(fmt_eixo_brl))
            ax.set_xlabel(f"Variacao no mes-alvo ({self.mes_alvo}) vs media baseline",
                          fontsize=10, color="#1A1A1A", fontweight="medium")
            ax.axvline(0, color="#444444", linewidth=0.8)
            estilo_eixos(ax, "Top 20 verbas por impacto absoluto")
        else:
            ax.text(0.5, 0.5, "Sem dados suficientes para o ranking de verbas.",
                    transform=ax.transAxes, ha="center", va="center",
                    fontsize=12, color=TEXTO_SECUNDARIO)
            ax.set_axis_off()
        self.chart_top_verbas.fig.tight_layout()
        self.chart_top_verbas.render()
        self._conectar_click_top_verbas()

        # ===== Lista de verbas para drill-down (ordenada PGTO -> DESC -> OUTRO) =====
        val_cols = [f"{m} - Valor" for m in self.meses_detectados]
        agg_full = df_base.groupby(["Código", "Descrição", "Clas."])[val_cols].sum().reset_index()
        agg_full["TOTAL_PERIODO_ABS"] = agg_full[val_cols].abs().sum(axis=1)
        agg_full = agg_full[agg_full["TOTAL_PERIODO_ABS"] > 0]
        # Ordem fixa: PGTO -> DESC -> OUTRO -> resto. Dentro da classe, por valor desc.
        ordem_clas = {"PGTO": 0, "DESC": 1, "OUTRO": 2}
        agg_full["_ord_clas"] = agg_full["Clas."].map(lambda c: ordem_clas.get(str(c), 99))
        agg_full = agg_full.sort_values(["_ord_clas", "TOTAL_PERIODO_ABS"],
                                           ascending=[True, False]).reset_index(drop=True)

        # Guarda lista completa para filtros e busca
        self.verbas_av_completa = agg_full

        # Popula listbox completa
        self._filtrar_listbox_verbas()

        # Default: seleciona a primeira verba (maior impacto da PGTO)
        if self.listbox_verbas_av.size() > 0 and not self.listbox_verbas_av.curselection():
            self.listbox_verbas_av.selection_set(0)
            self.listbox_verbas_av.see(0)

        self._render_drill_verba()

    def _filtrar_listbox_verbas(self):
        """Refiltra a listbox de verbas com base no termo de busca."""
        if not hasattr(self, "verbas_av_completa"):
            return
        # Preserva a selecao atual antes de repopular
        sel_atual = set()
        for i in self.listbox_verbas_av.curselection():
            try:
                cod_sel = self._extrair_codigo_listbox(i)
                if cod_sel:
                    sel_atual.add(cod_sel)
            except Exception:
                pass

        busca = self.entry_busca_verba_av.get().strip().upper() if hasattr(self, "entry_busca_verba_av") else ""
        df = self.verbas_av_completa.copy()
        if busca:
            df = df[
                df["Código"].astype(str).str.upper().str.contains(busca, na=False) |
                df["Descrição"].astype(str).str.upper().str.contains(busca, na=False) |
                df["Clas."].astype(str).str.upper().str.contains(busca, na=False)
            ]

        self.listbox_verbas_av.delete(0, "end")
        # Cores alternadas por classe para visual claro
        ultima_clas = None
        for _, r in df.iterrows():
            cod = str(r["Código"])
            cls = str(r["Clas."])
            desc = str(r["Descrição"])[:55]
            tot = fmt_brl(r["TOTAL_PERIODO_ABS"])
            # Indicador de classe na linha
            label = f"[{cls:5}] {cod:>7} - {desc:<58} {tot:>16}"
            self.listbox_verbas_av.insert("end", label)
            # Tinge as classes (alterna leve)
            if cls != ultima_clas:
                ultima_clas = cls

        # Restaura selecao para os codigos que ainda existem na listbox
        if sel_atual:
            for i in range(self.listbox_verbas_av.size()):
                cod_atual = self._extrair_codigo_listbox(i)
                if cod_atual in sel_atual:
                    self.listbox_verbas_av.selection_set(i)

    def _extrair_codigo_listbox(self, idx):
        """Extrai o codigo da verba a partir do texto da listbox no indice idx."""
        try:
            txt = self.listbox_verbas_av.get(idx)
            # Formato: "[CLAS ] CODIGO - DESCRICAO    TOTAL"
            partes = txt.split("]", 1)
            if len(partes) < 2:
                return None
            depois = partes[1].strip()
            cod_e_resto = depois.split(" - ", 1)
            return cod_e_resto[0].strip()
        except Exception:
            return None

    def _limpar_busca_verba_av(self):
        if hasattr(self, "entry_busca_verba_av"):
            self.entry_busca_verba_av.delete(0, "end")
            self._filtrar_listbox_verbas()

    def _selecionar_classe_verbas(self, classe):
        """Seleciona todas as verbas visiveis de uma classe (substitui selecao atual)."""
        self.listbox_verbas_av.selection_clear(0, "end")
        chave = f"[{classe:5}]"
        for i in range(self.listbox_verbas_av.size()):
            txt = self.listbox_verbas_av.get(i)
            if txt.startswith(chave):
                self.listbox_verbas_av.selection_set(i)
        self._render_drill_verba()

    def _limpar_selecao_verbas_av(self):
        self.listbox_verbas_av.selection_clear(0, "end")
        self._render_drill_verba()

    def _limpar_tree_av_colab(self, mensagem=""):
        """Limpa a tabela de colaboradores da aba Analise de Verbas."""
        if hasattr(self, "tree_av_colab"):
            for it in self.tree_av_colab.get_children():
                self.tree_av_colab.delete(it)
        self.av_colab_df = pd.DataFrame()
        self.av_colab_selecionado = None
        if hasattr(self, "lbl_av_colab_status"):
            self.lbl_av_colab_status.set_text(mensagem or "(sem colaboradores para exibir)")

    def _selecionar_verba_por_codigo_av(self, codigo):
        """Seleciona uma verba na listbox da aba Analise de Verbas e atualiza o drill."""
        if not hasattr(self, "listbox_verbas_av"):
            return
        cod_alvo = str(codigo).strip()
        self.listbox_verbas_av.selection_clear(0, "end")
        achou = False
        for i in range(self.listbox_verbas_av.size()):
            cod_atual = self._extrair_codigo_listbox(i)
            if str(cod_atual).strip() == cod_alvo:
                self.listbox_verbas_av.selection_set(i)
                self.listbox_verbas_av.see(i)
                achou = True
                break
        if achou:
            self.av_mes_clicado = self.mes_alvo
            self._render_drill_verba()

    def _conectar_click_top_verbas(self):
        """Conecta clique no grafico Top Verbas para selecionar a verba clicada."""
        if not hasattr(self, "chart_top_verbas"):
            return
        try:
            if self.av_top_verbas_cid is not None:
                self.chart_top_verbas.canvas.mpl_disconnect(self.av_top_verbas_cid)
        except Exception:
            pass
        self.av_top_verbas_cid = self.chart_top_verbas.canvas.mpl_connect(
            "button_press_event", self._on_top_verbas_click
        )

    def _on_top_verbas_click(self, event):
        """Ao clicar em uma barra do ranking, seleciona a verba correspondente."""
        if event.inaxes is None or not getattr(self, "av_top_verbas_bars", None):
            return
        for bar, codigo in self.av_top_verbas_bars:
            try:
                contem, _info = bar.contains(event)
            except Exception:
                contem = False
            if contem:
                self._selecionar_verba_por_codigo_av(codigo)
                break

    def _conectar_click_grafico_av(self):
        """Conecta clique no grafico HC/Media para abrir colaboradores do mes clicado."""
        if not hasattr(self, "chart_av_hc"):
            return
        try:
            if self.av_hc_cid is not None:
                self.chart_av_hc.canvas.mpl_disconnect(self.av_hc_cid)
        except Exception:
            pass
        self.av_hc_cid = self.chart_av_hc.canvas.mpl_connect(
            "button_press_event", self._on_av_hc_click
        )

    def _on_av_hc_click(self, event):
        """Identifica o mes clicado no grafico da verba e lista os colaboradores."""
        if event.inaxes is None or event.xdata is None:
            return
        if self.av_dados_plot is None or self.av_dados_plot.empty:
            return
        try:
            idx = int(round(event.xdata))
        except Exception:
            return
        if idx < 0 or idx >= len(self.av_dados_plot):
            return
        mes = str(self.av_dados_plot.iloc[idx]["Mes"])
        self.av_mes_clicado = mes
        self._render_colaboradores_verba_mes(mes=mes)

    def _render_colaboradores_verba_mes(self, codigos_sel=None, mes=None):
        """
        Lista os colaboradores que compoem a verba selecionada no mes clicado.
        Tambem mostra a media individual da verba na baseline e cruza com o status
        liquido do colaborador, permitindo ir para Tabela & Drill-down ou Outliers.
        """
        if not hasattr(self, "tree_av_colab"):
            return

        codigos_sel = list(codigos_sel or getattr(self, "av_codigos_sel", []))
        mes = mes or getattr(self, "av_mes_clicado", None) or self.mes_alvo

        for it in self.tree_av_colab.get_children():
            self.tree_av_colab.delete(it)
        self.av_colab_selecionado = None

        if not codigos_sel:
            self._limpar_tree_av_colab("Selecione uma ou mais verbas para ver os colaboradores.")
            return
        if not hasattr(self, "df_av_base") or self.df_av_base is None or self.df_av_base.empty:
            self._limpar_tree_av_colab("Base da Analise de Verbas ainda nao carregada.")
            return

        col_mes = f"{mes} - Valor"
        cols_base = [f"{m} - Valor" for m in self.meses_baseline if f"{m} - Valor" in self.df_av_base.columns]
        if col_mes not in self.df_av_base.columns:
            self._limpar_tree_av_colab(f"Mes {mes} nao encontrado na base.")
            return

        cods_norm = {str(c).strip().lstrip("0") or "0" for c in codigos_sel}
        tmp = self.df_av_base.copy()
        tmp["_COD_NORM_AV"] = tmp["Código"].astype(str).str.strip().str.lstrip("0").replace("", "0")
        sub = tmp[tmp["_COD_NORM_AV"].isin(cods_norm)].copy()
        if sub.empty:
            self._limpar_tree_av_colab("Nenhuma linha encontrada para a(s) verba(s) selecionada(s).")
            return

        cols_agg = list(dict.fromkeys(cols_base + [col_mes]))
        agg = sub.groupby(["Matrícula", "Nome"])[cols_agg].sum().reset_index()
        agg["VALOR_MES_VERBA"] = agg[col_mes]

        if cols_base:
            base = agg[cols_base].replace(0, np.nan)
            agg["MEDIA_BASE_VERBA"] = base.mean(axis=1)
        else:
            agg["MEDIA_BASE_VERBA"] = np.nan

        agg["VAR_ABS_VERBA"] = agg["VALOR_MES_VERBA"] - agg["MEDIA_BASE_VERBA"].fillna(0)
        agg["VAR_PCT_VERBA"] = np.where(
            agg["MEDIA_BASE_VERBA"].notna() & (agg["MEDIA_BASE_VERBA"] != 0),
            ((agg["VALOR_MES_VERBA"] / agg["MEDIA_BASE_VERBA"]) - 1) * 100,
            np.nan
        )

        # Mantem quem teve verba no mes OU possuia historico na baseline
        agg = agg[(agg["VALOR_MES_VERBA"] != 0) | agg["MEDIA_BASE_VERBA"].notna()].copy()
        if agg.empty:
            self._limpar_tree_av_colab(f"Nenhum colaborador com movimento/historico da verba em {mes}.")
            return

        # Cruza com a auditoria liquida por funcionario
        if self.liquido_analise is not None:
            liq = self.liquido_analise.reset_index().copy()
            liq["Matrícula"] = liq["Matrícula"].astype(str)
            agg["Matrícula"] = agg["Matrícula"].astype(str)
            cols_liq = ["Matrícula", "Nome", "STATUS", "LIQUIDO_ALVO", "VAR_ABS", "VAR_PCT", "Z_SCORE", "AUDITORIA"]
            liq = liq[[c for c in cols_liq if c in liq.columns]]
            agg = agg.merge(liq, on=["Matrícula", "Nome"], how="left")
        else:
            agg["STATUS"] = "-"
            agg["LIQUIDO_ALVO"] = np.nan
            agg["AUDITORIA"] = ""

        agg["IMPACTO_ABS_VERBA"] = agg["VAR_ABS_VERBA"].abs()
        agg["VALOR_ABS_VERBA"] = agg["VALOR_MES_VERBA"].abs()
        agg = agg.sort_values(["IMPACTO_ABS_VERBA", "VALOR_ABS_VERBA"], ascending=[False, False]).reset_index(drop=True)
        self.av_colab_df = agg

        total = len(agg)
        qtd_com_valor = int((agg["VALOR_MES_VERBA"] != 0).sum())
        qtd_zerou = int(((agg["VALOR_MES_VERBA"] == 0) & agg["MEDIA_BASE_VERBA"].notna()).sum())
        cods_txt = ", ".join(codigos_sel) if len(codigos_sel) <= 8 else f"{len(codigos_sel)} verbas"
        self.lbl_av_colab_status.set_text(
            f"Mes analisado: {mes} | Verba(s): {cods_txt} | "
            f"{qtd_com_valor} colaborador(es) com valor no mes; {qtd_zerou} zerado(s) com historico; "
            f"{total} linha(s) exibiveis. Duplo clique abre o drill-down do colaborador."
        )

        for _, row in agg.head(1000).iterrows():
            status = str(row.get("STATUS", "-"))
            tags = ()
            if status in STATUS_CRITICOS:
                tags = ("critico",)
            elif status == "DESLIGADO":
                tags = ("desligado",)

            self.tree_av_colab.insert("", "end", values=(
                str(row.get("Matrícula", "")),
                str(row.get("Nome", ""))[:60],
                status,
                fmt_brl(row.get("VALOR_MES_VERBA", np.nan)),
                fmt_brl(row.get("MEDIA_BASE_VERBA", np.nan)),
                fmt_brl(row.get("VAR_ABS_VERBA", np.nan)),
                fmt_pct(row.get("VAR_PCT_VERBA", np.nan)),
                fmt_brl(row.get("LIQUIDO_ALVO", np.nan)),
                str(row.get("AUDITORIA", ""))[:140],
            ), tags=tags)

        if total > 1000:
            self.lbl_av_colab_status.set_text(self.lbl_av_colab_status.cget("text") + " | Mostrando primeiras 1000 linhas.")

    def _on_av_colab_select(self, _event=None):
        """Guarda o colaborador selecionado na tabela de Analise de Verbas."""
        if not hasattr(self, "tree_av_colab"):
            return
        sel = self.tree_av_colab.selection()
        if not sel:
            self.av_colab_selecionado = None
            return
        vals = self.tree_av_colab.item(sel[0], "values")
        if not vals:
            self.av_colab_selecionado = None
            return
        mat = str(vals[0])
        nome_curto = str(vals[1])
        nome_full = nome_curto
        if self.av_colab_df is not None and not self.av_colab_df.empty:
            m = self.av_colab_df[self.av_colab_df["Matrícula"].astype(str) == mat]
            if not m.empty:
                nome_full = str(m.iloc[0]["Nome"])
        self.av_colab_selecionado = (mat, nome_full)

    def _selecionar_colaborador_na_tabela_principal(self, mat, nome):
        """Filtra a Tabela & Drill-down pela matricula e abre o drill do colaborador."""
        if self.liquido is None:
            return False
        self.combo_status_tab.set("(todos)")
        self.entry_busca.delete(0, "end")
        self.entry_busca.insert(0, str(mat))
        self._renderizar_tabela()

        alvo_item = None
        for item in self.tree.get_children():
            vals = self.tree.item(item, "values")
            if vals and str(vals[0]) == str(mat):
                alvo_item = item
                break
        if alvo_item is not None:
            self.tree.selection_set(alvo_item)
            self.tree.focus(alvo_item)
            self.tree.see(alvo_item)

        # Usa o nome completo da base liquida quando possivel
        try:
            liq_reset = self.liquido.reset_index()
            m = liq_reset[liq_reset["Matrícula"].astype(str) == str(mat)]
            if not m.empty:
                nome = str(m.iloc[0]["Nome"])
        except Exception:
            pass

        self.func_selecionado = (str(mat), nome)
        self._render_drill()
        return True

    def _ir_para_tabela_drill_colaborador_av(self):
        """Navega da Analise de Verbas para a Tabela & Drill-down do colaborador."""
        if self.av_colab_selecionado is None:
            messagebox.showinfo("Analise de Verbas", "Selecione um colaborador na tabela de colaboradores da verba.")
            return
        mat, nome = self.av_colab_selecionado
        self.tabview.set("Tabela & Drill-down")
        self._selecionar_colaborador_na_tabela_principal(mat, nome)

    def _ir_para_outliers_colaborador_av(self):
        """Navega para a aba Outliers com destaque visual no colaborador selecionado."""
        if self.av_colab_selecionado is None:
            messagebox.showinfo("Analise de Verbas", "Selecione um colaborador na tabela de colaboradores da verba.")
            return
        mat, nome = self.av_colab_selecionado
        self.outlier_foco = (str(mat), nome)
        self._render_outliers()
        self.tabview.set("Outliers")

    def _destacar_outlier_foco(self, ax, ax_leg, valido, lim_max):
        """Desenha destaque no ponto do colaborador quando vier da Analise de Verbas."""
        foco = getattr(self, "outlier_foco", None)
        if not foco or valido is None or valido.empty:
            return
        mat, nome = foco
        try:
            v = valido.reset_index().copy()
            v["Matrícula"] = v["Matrícula"].astype(str)
            alvo = v[v["Matrícula"] == str(mat)]
            if alvo.empty:
                ax_leg.text(0.02, 0.24,
                              f"Foco: {mat}\nNao exibido no grafico\n(baseline menor que R$ 100 ou sem dados).",
                              transform=ax_leg.transAxes, fontsize=8.5, color=IG_LARANJA,
                              ha="left", va="top", fontweight="bold")
                return
            row = alvo.iloc[0]
            x = row["MEDIA_BASELINE"]
            y = row.get("LIQ_PLOT", row.get("LIQUIDO_ALVO", 100))
            y = max(float(y), 100.0)
            ax.scatter([x], [y], s=280, facecolors="none", edgecolors=IG_LARANJA,
                       linewidths=2.6, zorder=20)
            ax.annotate("Selecionado", xy=(x, y), xytext=(12, 12), textcoords="offset points",
                        fontsize=9, fontweight="bold", color=IG_LARANJA,
                        bbox=dict(boxstyle="round,pad=0.25", fc="white", ec=IG_LARANJA, lw=1.2),
                        arrowprops=dict(arrowstyle="->", color=IG_LARANJA, lw=1.2),
                        zorder=21)
            nome_txt = str(row.get("Nome", nome))[:28]
            ax_leg.text(0.02, 0.24,
                          f"Foco vindo da Analise de Verbas\n{mat} - {nome_txt}\n"
                          f"Status: {row.get('STATUS', '-')}\n"
                          f"Impacto liquido: {fmt_brl(row.get('VAR_ABS', np.nan))}",
                          transform=ax_leg.transAxes, fontsize=8.5, color=IG_LARANJA,
                          ha="left", va="top", fontweight="bold")
        except Exception:
            return

    def _render_drill_verba(self):
        if not hasattr(self, "df_av_base") or self.df_av_base is None:
            return
        if not hasattr(self, "listbox_verbas_av"):
            return

        # Coleta os codigos das verbas selecionadas
        codigos_sel = []
        labels_sel = []
        for i in self.listbox_verbas_av.curselection():
            cod = self._extrair_codigo_listbox(i)
            if cod:
                codigos_sel.append(cod)
                labels_sel.append(self.listbox_verbas_av.get(i))

        # Atualiza contador lateral
        if not codigos_sel:
            self.lbl_contador_verbas_av.set_text("(nenhuma)")
            self.chart_av_hc.limpar(); self.chart_av_hc.render()
            self.card_av_hc.atualizar(titulo="HC", valor="-", delta="")
            self.card_av_total.atualizar(titulo="Total", valor="-", delta="")
            self.card_av_media.atualizar(titulo="Valor medio", valor="-", delta="")
            self.card_av_status.atualizar(titulo="Status", valor="-", delta="")
            self.lbl_av_aviso.set_text("Selecione uma ou mais verbas na listbox para ver os graficos.")
            self.av_codigos_sel = []
            self.av_dados_plot = None
            self._limpar_tree_av_colab("Selecione uma ou mais verbas para ver os colaboradores.")
            return

        # Atualiza painel lateral com a lista das selecionadas
        if len(codigos_sel) <= 8:
            txt_contador = f"{len(codigos_sel)} verba(s):\n" + "\n".join(f"- {c}" for c in codigos_sel)
        else:
            primeiras = "\n".join(f"- {c}" for c in codigos_sel[:6])
            txt_contador = f"{len(codigos_sel)} verbas:\n{primeiras}\n  (... +{len(codigos_sel)-6})"
        self.lbl_contador_verbas_av.set_text(txt_contador)

        # Calcula HC e total agregados
        self.av_codigos_sel = list(codigos_sel)
        dados = hc_e_total_por_verba(self.df_av_base, self.meses_detectados, codigos_sel)
        self.av_dados_plot = dados.copy() if dados is not None else None
        if dados.empty:
            self._limpar_tree_av_colab("Sem dados para a(s) verba(s) selecionada(s).")
            return

        # Cards
        try:
            linha = dados[dados["Mes"] == self.mes_alvo].iloc[0]
            base = dados[dados["Mes"].isin(self.meses_baseline)]
            hc_med = base["HC"].mean() if not base.empty else 0
            media_med = base["Media"].mean() if not base.empty else 0
            total_med = base["Total"].mean() if not base.empty else 0

            var_hc = (linha["HC"] / hc_med - 1) * 100 if hc_med else 0
            var_med = (linha["Media"] / media_med - 1) * 100 if media_med else 0
            var_tot = (linha["Total"] / total_med - 1) * 100 if total_med else 0

            cl_hc = "" if abs(var_hc) < 5 else ("atencao" if abs(var_hc) < 15 else "alerta")
            cl_med = "" if abs(var_med) < 5 else ("atencao" if abs(var_med) < 15 else "alerta")
            cl_tot = "" if abs(var_tot) < 5 else ("atencao" if abs(var_tot) < 15 else "alerta")

            qtd_label = f"({len(codigos_sel)} verbas)" if len(codigos_sel) > 1 else "(1 verba)"
            self.card_av_hc.atualizar(titulo=f"HC {self.mes_alvo} {qtd_label}",
                                        valor=fmt_int(linha["HC"]),
                                        delta=f"{fmt_pct(var_hc)} vs media baseline",
                                        classe=cl_hc + (" negativo" if var_hc < 0 else ""))
            self.card_av_total.atualizar(titulo=f"Total agregado ({self.mes_alvo})",
                                           valor=fmt_brl(linha["Total"]),
                                           delta=f"{fmt_pct(var_tot)} vs media baseline",
                                           classe=cl_tot + (" negativo" if var_tot < 0 else ""))
            self.card_av_media.atualizar(titulo="Valor medio por matricula",
                                           valor=fmt_brl(linha["Media"]),
                                           delta=f"{fmt_pct(var_med)} vs media baseline",
                                           classe=cl_med + (" negativo" if var_med < 0 else ""))

            # Status: quantas das selecionadas estao excluidas
            excl_count = sum(1 for c in codigos_sel if c in self.verbas_excluidas)
            if excl_count == 0:
                status_txt = "Todas incluidas"
                classe_st = ""
            elif excl_count == len(codigos_sel):
                status_txt = "Todas excluidas"
                classe_st = "atencao"
            else:
                status_txt = f"{excl_count}/{len(codigos_sel)} excluidas"
                classe_st = "atencao"
            self.card_av_status.atualizar(titulo="Status no calculo",
                                            valor=status_txt,
                                            delta=f"{len(codigos_sel)} verba(s)",
                                            classe=classe_st)
        except (IndexError, KeyError):
            pass

        # Grafico HC (barras) + Valor medio (linha)
        self.chart_av_hc.limpar()
        ax1 = self.chart_av_hc.fig.add_subplot(111)
        x = np.arange(len(dados["Mes"]))
        bars = ax1.bar(x, dados["HC"], color=IG_AZUL_CLR,
                        edgecolor=IG_AZUL_ESC, linewidth=0.6)
        max_hc = dados["HC"].max() if len(dados) else 1
        ax1.set_ylim(0, max(max_hc * 1.18, 1))
        for bar, v in zip(bars, dados["HC"]):
            if v > 0:
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max_hc * 0.015,
                         fmt_int(v), ha="center", va="bottom",
                         fontsize=9, fontweight="bold", color=IG_AZUL_ESC)
        ax1.set_xticks(x)
        ax1.set_xticklabels(dados["Mes"], rotation=45, ha="right", fontsize=9)
        ax1.set_ylabel("HC matriculas unicas (quantidade)", color="#1A1A1A",
                        fontsize=10, fontweight="medium")

        base = dados[dados["Mes"].isin(self.meses_baseline)]
        if len(base) and base["HC"].mean() > 0:
            hc_med = base["HC"].mean()
            ax1.axhline(hc_med, color=IG_VERDE_ESC, linestyle=":", linewidth=1.5)
            ax1.text(0, hc_med, f" HC medio baseline: {int(hc_med)}",
                      color=IG_VERDE_ESC, fontsize=9, fontweight="bold", va="bottom")

        # Titulo dinamico
        if len(codigos_sel) == 1:
            titulo_grafico = labels_sel[0].strip() if labels_sel else f"Verba {codigos_sel[0]}"
            if len(titulo_grafico) > 80:
                titulo_grafico = titulo_grafico[:78] + ".."
        else:
            titulo_grafico = f"{len(codigos_sel)} verbas agregadas: {', '.join(codigos_sel[:4])}"
            if len(codigos_sel) > 4:
                titulo_grafico += f", +{len(codigos_sel)-4}"
        estilo_eixos(ax1, titulo_grafico)

        ax2 = ax1.twinx()
        ax2.plot(x, dados["Media"], color=IG_AZUL_ESC,
                  marker="o", linewidth=2.4, markersize=8,
                  markeredgecolor="white", markeredgewidth=1)
        for xi, yi in zip(x, dados["Media"]):
            if yi != 0:
                ax2.text(xi, yi, f" {fmt_eixo_brl(yi, None)}",
                          fontsize=8.5, fontweight="bold", color=IG_AZUL_ESC,
                          ha="left", va="bottom")
        ax2.set_ylabel("Valor medio agregado por matricula (R$)", color=IG_AZUL_ESC,
                        fontsize=10, fontweight="medium")
        ax2.yaxis.set_major_formatter(FuncFormatter(fmt_eixo_brl))
        ax2.spines["top"].set_visible(False)
        ax2.tick_params(colors="#1A1A1A", labelsize=9)

        idx_alvo = list(dados["Mes"]).index(self.mes_alvo) if self.mes_alvo in list(dados["Mes"]) else -1
        if idx_alvo >= 0:
            ax1.axvspan(idx_alvo - 0.5, idx_alvo + 0.5, color=IG_AZUL_CLR, alpha=0.20)

        # Realce sutil no mes selecionado por clique, quando houver
        if getattr(self, "av_mes_clicado", None) in list(dados["Mes"]):
            try:
                idx_sel = list(dados["Mes"]).index(self.av_mes_clicado)
                ax1.axvspan(idx_sel - 0.5, idx_sel + 0.5, color=IG_LARANJA, alpha=0.10)
            except Exception:
                pass

        self.chart_av_hc.fig.tight_layout()
        self.chart_av_hc.render()
        self._conectar_click_grafico_av()

        # Aviso explicativo
        try:
            linha_alvo_av = dados[dados["Mes"] == self.mes_alvo].iloc[0]
            cods_str = ", ".join(codigos_sel) if len(codigos_sel) <= 8 else f"{len(codigos_sel)} verbas"
            self.lbl_av_aviso.set_text(
                f"Verba(s) [{cods_str}] em {self.mes_alvo}: "
                f"{fmt_int(linha_alvo_av['HC'])} matriculas, "
                f"total agregado {fmt_brl(linha_alvo_av['Total'])}, "
                f"valor medio por matricula {fmt_brl(linha_alvo_av['Media'])}.  "
                f"O HC nao e a soma dos HCs por verba: e a quantidade de matriculas "
                f"unicas que tem valor != 0 em qualquer das verbas selecionadas."
            )
        except (IndexError, KeyError):
            self.lbl_av_aviso.set_text("")

        # Atualiza automaticamente a tabela de colaboradores.
        # Por padrao usa o mes-alvo; se o usuario clicou em outro mes, preserva o mes clicado.
        mes_colab = getattr(self, "av_mes_clicado", None) or self.mes_alvo
        if mes_colab not in list(dados["Mes"]):
            mes_colab = self.mes_alvo
        self._render_colaboradores_verba_mes(codigos_sel=codigos_sel, mes=mes_colab)

    # ---------- Filtro de verbas ----------
    def acao_filtrar_verbas(self):
        if self.df is None:
            return
        # Aplica os filtros de empresa/processo antes (visao consistente com o reprocessamento)
        df_f = self.df[self.df["Empresa"].astype(str).isin(self.empresas_selecionadas)].copy()
        if self.processos_disponiveis:
            df_f = df_f[df_f["Processo"].astype(str).isin(self.processos_selecionados)].copy()
        if df_f.empty:
            messagebox.showinfo("Filtrar verbas", "Nenhuma linha disponivel para filtrar.")
            return
        dlg = FiltroVerbasDialog(self, df_f, self.meses_detectados, self.verbas_excluidas)
        self.wait_window(dlg)
        # Apos fechar, dlg.resultado contem o novo set ou None (cancelado)
        if dlg.resultado is not None:
            self.verbas_excluidas = dlg.resultado
            if self.verbas_excluidas:
                self.lbl_verbas_status.configure(
                    text=f"{len(self.verbas_excluidas)} verba(s) excluida(s) dos calculos",
                    text_color=IG_VINHO)
            else:
                self.lbl_verbas_status.configure(
                    text="(todas as verbas incluidas)", text_color=TEXTO_SECUNDARIO)
            # Reprocessa automaticamente
            self.acao_reprocessar()

    # ---------- Exportar ----------
    def acao_exportar_excel(self):
        if self.liquido is None:
            messagebox.showinfo("Exportar", "Carregue e processe um arquivo antes de exportar.")
            return
        path = filedialog.asksaveasfilename(
            title="Salvar relatorio Excel",
            defaultextension=".xlsx",
            initialfile=f"auditoria_folha_{self.mes_alvo.replace('/', '-')}.xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return

        self.lbl_exportar_status.configure(text="Gerando arquivo...", text_color=IG_AZUL_ESC)
        self.update_idletasks()
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                self.resumo.to_excel(w, sheet_name="Resumo_Mensal", index=False)
                liq = self.liquido.reset_index()
                ordem = ["AUSENTE", "NEGATIVO", "ZERO_SUSPEITO", "EXTREMA", "EXTREMA_IQR",
                         "Z_2SIGMA", "ALTA", "ALTA_IQR", "MATERIAL", "NOVO_FUNC", "OK",
                         "SEM_DADOS", "DESLIGADO"]
                liq["STATUS"] = pd.Categorical(liq["STATUS"], categories=ordem, ordered=True)
                liq = liq.sort_values(["STATUS", "VAR_ABS"], na_position="last")
                liq.to_excel(w, sheet_name="Liquido_Funcionario", index=False)
                self.zerados_pgto.reset_index().to_excel(w, sheet_name="Verbas_PGTO_Zeradas", index=False)
                self.zerados_desc.reset_index().to_excel(w, sheet_name="Verbas_DESC_Zeradas", index=False)
                # Aba de metadados/parametros
                meta = pd.DataFrame([
                    {"chave": "metodo", "valor": self.metadata_audit.get("metodo", "")},
                    {"chave": "config", "valor": self.config_param_nome},
                    {"chave": "mes_alvo", "valor": self.mes_alvo},
                    {"chave": "meses_baseline", "valor": ", ".join(self.meses_baseline)},
                    {"chave": "folha_total_alvo", "valor": self.metadata_audit.get("folha_total_alvo", 0)},
                    {"chave": "materialidade", "valor": self.metadata_audit.get("materialidade", 0)},
                    {"chave": "Q1_var_abs", "valor": self.metadata_audit.get("q1", "")},
                    {"chave": "Q3_var_abs", "valor": self.metadata_audit.get("q3", "")},
                    {"chave": "IQR_var_abs", "valor": self.metadata_audit.get("iqr", "")},
                    {"chave": "gerado_em", "valor": datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
                ])
                meta.to_excel(w, sheet_name="Parametros", index=False)
            self.lbl_exportar_status.configure(
                text=f"Arquivo salvo: {os.path.basename(path)}",
                text_color=IG_VERDE_ESC)
        except Exception as e:
            self.lbl_exportar_status.configure(text=f"Erro: {e}", text_color=IG_VERMELHO)
            messagebox.showerror("Erro ao exportar", str(e))


# -----------------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    app = AuditoriaFolhaApp()
    app.mainloop()
