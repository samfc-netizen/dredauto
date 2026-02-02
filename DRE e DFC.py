import os
import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

# ============================================================
# CONFIG
# ============================================================
# No Streamlit Cloud, o Excel deve estar no próprio repositório.
# Para atualizar os dados, basta substituir/commitar o Excel no GitHub.
ARQUIVO_EXCEL = "BASE DRE DAUTO TINTAS.xlsx"

# Caminho absoluto do diretório do app (compatível com Streamlit Cloud)
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
CAMINHO_LOCAL = os.path.join(BASE_DIR, ARQUIVO_EXCEL)

SHEET_FAT = "CMV E FATURAMENTO"
SHEET_IMP = "IMPOSTOS E FOLHA"
SHEET_DRE = "DRE"

MESES = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

# >>> Regra de competência: IMPOSTOS/FOLHA (DTA.PAG) do mês anterior
# Ex.: DTA.PAG em DEZ -> apropriar em JAN na DRE
IMP_MONTH_OFFSET = +1  # DEZ -> JAN

# CMV E FATURAMENTO
COL_RECEITA = "VR.TOTAL"
COL_CMV = "CUSTO"

# IMPOSTOS E FOLHA
COL_CONTA_IMP = "CONTA DE RESULTADO"
COL_DESPESA_IMP = "DESPESA"
COL_VAL_IMP = "VAL.PAG"
COL_DATA_IMP = "DTA.PAG"   # <<< FIXO conforme você informou

# DRE
COL_CONTA_DRE = "CONTA DE RESULTADO"
COL_DESPESA_DRE = "DESPESA"
COL_VAL_DRE = "VAL.PAG"

# Auto-detect (se falhar, travar manualmente aqui)
COL_LOJA_FAT = None
COL_LOJA_IMP = None
COL_LOJA_DRE = None

COL_MES_FAT = None
COL_MES_DRE = None

COL_HIST_IMP = None
COL_HIST_DRE = None

# Contas
COD_00004 = "00004"  # DEDUÇÕES (IMPOSTOS SOBRE VENDAS) - IMPOSTOS E FOLHA
COD_00006 = "00006"  # DESPESAS COM PESSOAL             - IMPOSTOS E FOLHA
COD_00007 = "00007"  # DESPESAS ADMINISTRATIVAS         - DRE
COD_00009 = "00009"  # DESPESAS COMERCIAIS              - DRE
COD_00011 = "00011"  # DESPESAS FINANCEIRAS             - DRE
COD_00017 = "00017"  # DESPESAS OPERACIONAIS            - DRE


# ============================================================
# HELPERS
# ============================================================
def format_brl(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def format_pct(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    try:
        return f"{float(v) * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def coerce_number(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s.astype(float)
    tmp = (
        s.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.strip()
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(tmp, errors="coerce")


def detect_col(df: pd.DataFrame, manual: str | None, candidates: list[str]) -> str | None:
    if manual and manual in df.columns:
        return manual
    low = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in low:
            return low[cand.lower()]
    return None


def detect_loja_col(df: pd.DataFrame, manual: str | None):
    return detect_col(df, manual, ["LOJA", "FILIAL", "UNIDADE", "PONTO DE VENDA", "PDV", "EMPRESA"])


def detect_mes_col(df: pd.DataFrame, manual: str | None):
    return detect_col(df, manual, ["MÊS", "MES", "COMPETÊNCIA", "COMPETENCIA", "DATA", "DT", "DATA EMISSÃO", "DATA EMISSÃO NF"])


def detect_hist_col(df: pd.DataFrame, manual: str | None):
    return detect_col(df, manual, ["HISTÓRICO", "HISTORICO", "HIST", "OBS", "OBSERVAÇÃO", "OBSERVACAO", "DESCRIÇÃO", "DESCRICAO"])


def safe_div(a: float, b: float) -> float:
    if b == 0:
        return np.nan
    return a / b


def parse_mes(v) -> str | None:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None

    # tenta data primeiro (para DTA.PAG)
    dt = pd.to_datetime(v, errors="coerce", dayfirst=True)
    if pd.notna(dt):
        return MESES[int(dt.month) - 1]

    s = str(v).strip().upper()
    for m in MESES:
        if m in s:
            return m

    mnum = re.findall(r"\b(1[0-2]|[1-9])\b", s)
    if mnum:
        mn = int(mnum[0])
        return MESES[mn - 1]

    return None


def shift_mes(mes: str, offset: int) -> str:
    if mes not in MESES:
        return mes
    i = MESES.index(mes)
    j = (i + offset) % 12
    return MESES[j]


def build_month_field(df: pd.DataFrame, col_mes: str | None, month_offset: int = 0) -> pd.Series:
    if col_mes is None or col_mes not in df.columns:
        base = pd.Series(["JAN"] * len(df), index=df.index)
    else:
        base = df[col_mes].apply(parse_mes).fillna("JAN")
    if month_offset != 0:
        return base.apply(lambda m: shift_mes(m, month_offset))
    return base


def filter_lojas(df: pd.DataFrame, col_loja: str | None, lojas_sel: list[str]) -> pd.DataFrame:
    if not col_loja or col_loja not in df.columns:
        return df
    if not lojas_sel:
        return df
    return df[df[col_loja].astype(str).str.strip().isin(lojas_sel)]


def conta_codigo(series_conta: pd.Series) -> pd.Series:
    """
    Extrai o código 00004/00006 etc. de CONTA DE RESULTADO,
    independentemente de vir como texto "00004 - ..." ou número 4.
    """
    s = series_conta.astype(str).str.strip()

    # remove .0 quando vem de Excel como número
    s = s.str.replace(r"\.0$", "", regex=True)

    # pega 1 a 5 dígitos do começo
    m = s.str.extract(r"^\s*(\d{1,5})")[0]
    # zfill para 5 dígitos
    out = m.fillna("").apply(lambda x: x.zfill(5) if x != "" else "")
    return out


def sum_monthly_total(df: pd.DataFrame, col_val: str, col_loja: str | None, lojas_sel: list[str], col_mes: str | None, month_offset: int = 0) -> pd.Series:
    x = df.copy()
    x = filter_lojas(x, col_loja, lojas_sel)
    x[col_val] = coerce_number(x[col_val])
    x = x.dropna(subset=[col_val])
    x["MES__"] = build_month_field(x, col_mes, month_offset=month_offset)

    g = x.groupby("MES__", as_index=True)[col_val].sum()
    out = pd.Series(0.0, index=MESES)
    for m in MESES:
        if m in g.index:
            out[m] = float(g.loc[m])
    return out


def sum_monthly_conta_codigo(df: pd.DataFrame, col_conta: str, cod: str, col_val: str,
                             col_loja: str | None, lojas_sel: list[str], col_mes: str | None,
                             month_offset: int = 0) -> pd.Series:
    x = df.copy()
    x = filter_lojas(x, col_loja, lojas_sel)

    if col_conta not in x.columns or col_val not in x.columns:
        return pd.Series(0.0, index=MESES)

    x[col_val] = coerce_number(x[col_val])
    x = x.dropna(subset=[col_val])

    x["COD__"] = conta_codigo(x[col_conta])
    y = x[x["COD__"] == cod]
    if y.empty:
        return pd.Series(0.0, index=MESES)

    y["MES__"] = build_month_field(y, col_mes, month_offset=month_offset)
    g = y.groupby("MES__", as_index=True)[col_val].sum()

    out = pd.Series(0.0, index=MESES)
    for m in MESES:
        if m in g.index:
            out[m] = float(g.loc[m])
    return out


def drill_despesas_unique_codigo(
    df: pd.DataFrame,
    col_conta: str,
    col_despesa: str,
    col_val: str,
    col_loja: str | None,
    lojas_sel: list[str],
    cod: str,
    col_mes: str | None,
    meses_ref: list[str],
    col_hist: str | None,
    month_offset: int = 0,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    needed = [col_conta, col_despesa, col_val]
    if any(c not in df.columns for c in needed):
        return (
            pd.DataFrame(columns=["DESPESA", "VALOR", "%_SOBRE_CONTA"]),
            pd.DataFrame(columns=["DESPESA", "LOJA", "HIST"]),
        )

    x = df.copy()
    x = filter_lojas(x, col_loja, lojas_sel)
    x[col_val] = coerce_number(x[col_val])
    x = x.dropna(subset=[col_val])

    x["COD__"] = conta_codigo(x[col_conta])
    y = x[x["COD__"] == cod].copy()
    if y.empty:
        return (
            pd.DataFrame(columns=["DESPESA", "VALOR", "%_SOBRE_CONTA"]),
            pd.DataFrame(columns=["DESPESA", "LOJA", "HIST"]),
        )

    y["MES__"] = build_month_field(y, col_mes, month_offset=month_offset)
    y = y[y["MES__"].isin(meses_ref)]
    if y.empty:
        return (
            pd.DataFrame(columns=["DESPESA", "VALOR", "%_SOBRE_CONTA"]),
            pd.DataFrame(columns=["DESPESA", "LOJA", "HIST"]),
        )

    if col_loja and col_loja in y.columns:
        y["LOJA__"] = y[col_loja].astype(str).str.strip()
    else:
        y["LOJA__"] = "TODAS"

    if col_hist and col_hist in y.columns:
        y["HIST__"] = y[col_hist].astype(str).str.strip()
    else:
        y["HIST__"] = ""

    y["DESPESA__"] = y[col_despesa].astype(str).str.strip()

    agg = y.groupby("DESPESA__", as_index=False)[col_val].sum().rename(columns={"DESPESA__": "DESPESA", col_val: "VALOR"})
    total = float(agg["VALOR"].sum()) if len(agg) else 0.0
    agg["%_SOBRE_CONTA"] = agg["VALOR"].apply(lambda v: safe_div(v, total))
    agg = agg.sort_values("VALOR", ascending=False).reset_index(drop=True)

    raw = y[["DESPESA__", "LOJA__", "HIST__"]].rename(columns={"DESPESA__": "DESPESA", "LOJA__": "LOJA", "HIST__": "HIST"})
    raw["HIST"] = raw["HIST"].astype(str).str.strip()
    raw = raw[(raw["HIST"] != "") & (raw["HIST"].str.lower() != "nan")]
    raw = raw.drop_duplicates(subset=["DESPESA", "LOJA", "HIST"]).reset_index(drop=True)

    return agg, raw


# ============================================================
# APP
# ============================================================
st.set_page_config(page_title="DRE Lojas- Dauto Tintas", layout="wide")
st.title("DRE Lojas Dauto Tintas")

# ====== EXCEL (ARQUIVO NO REPOSITÓRIO) ======
if not os.path.exists(CAMINHO_LOCAL):
    st.error("Arquivo Excel não encontrado no repositório do app.")
    st.code(f"Esperado em: {CAMINHO_LOCAL}")
    try:
        st.write("Arquivos disponíveis na pasta do app:")
        st.write(sorted(os.listdir(BASE_DIR)))
    except Exception:
        pass
    st.stop()

try:
    EXCEL_XLS = pd.ExcelFile(CAMINHO_LOCAL)
df_fat = pd.read_excel(EXCEL_XLS, sheet_name=SHEET_FAT)
    df_imp = pd.read_excel(EXCEL_XLS, sheet_name=SHEET_IMP)
    df_dre = pd.read_excel(EXCEL_XLS, sheet_name=SHEET_DRE)
except Exception as e:
    st.error(f"Erro lendo as abas do Excel: {e}")
    st.stop()

df_fat.columns = [str(c).strip() for c in df_fat.columns]
df_imp.columns = [str(c).strip() for c in df_imp.columns]
df_dre.columns = [str(c).strip() for c in df_dre.columns]

# Detect colunas
loja_fat = detect_loja_col(df_fat, COL_LOJA_FAT)
loja_imp = detect_loja_col(df_imp, COL_LOJA_IMP)
loja_dre = detect_loja_col(df_dre, COL_LOJA_DRE)

mes_fat = detect_mes_col(df_fat, COL_MES_FAT)
mes_dre = detect_mes_col(df_dre, COL_MES_DRE)

# >>> FIXO: na aba IMPOSTOS E FOLHA, a "data/competência" é DTA.PAG
mes_imp = COL_DATA_IMP

hist_imp = detect_hist_col(df_imp, COL_HIST_IMP)
hist_dre = detect_hist_col(df_dre, COL_HIST_DRE)

# Validar colunas essenciais
missing = []
for col in [COL_RECEITA, COL_CMV]:
    if col not in df_fat.columns:
        missing.append(f"{SHEET_FAT}.{col}")

for col in [COL_CONTA_IMP, COL_DESPESA_IMP, COL_VAL_IMP, COL_DATA_IMP]:
    if col not in df_imp.columns:
        missing.append(f"{SHEET_IMP}.{col}")

for col in [COL_CONTA_DRE, COL_DESPESA_DRE, COL_VAL_DRE]:
    if col not in df_dre.columns:
        missing.append(f"{SHEET_DRE}.{col}")

if missing:
    st.error("Colunas obrigatórias não encontradas:\n- " + "\n- ".join(missing))
    st.stop()

# Lojas disponíveis
lojas_set = set()
for df, col in [(df_fat, loja_fat), (df_imp, loja_imp), (df_dre, loja_dre)]:
    if col and col in df.columns:
        lojas_set |= set(df[col].astype(str).str.strip().dropna().unique())
lojas_all = sorted([l for l in lojas_set if l and l.upper() != "NAN"])

# ============================================================
# FILTROS
# ============================================================
st.sidebar.header("Filtros")

TODAS_OPT = "TODAS"
lojas_opt = [TODAS_OPT] + lojas_all

default_sel_lojas = st.session_state.get("lojas_multi", [TODAS_OPT])
sel_lojas = st.sidebar.multiselect("Lojas:", options=lojas_opt, default=default_sel_lojas, key="lojas_multi")

if (TODAS_OPT in sel_lojas) or (len(sel_lojas) == 0):
    lojas_sel = lojas_all[:]
else:
    lojas_sel = [x for x in sel_lojas if x != TODAS_OPT]

default_meses = st.session_state.get("meses_multi", ["JAN"])
sel_meses = st.sidebar.multiselect("Meses:", options=MESES, default=default_meses, key="meses_multi")
meses_ref = sel_meses[:] if sel_meses else MESES[:]
periodo_label = " + ".join(meses_ref) if len(meses_ref) <= 3 else f"{len(meses_ref)} meses"

# ============================================================
# CÁLCULOS (JAN..DEZ)
# ============================================================
# FAT (sem offset)
receita_m = sum_monthly_total(df_fat, COL_RECEITA, loja_fat, lojas_sel, mes_fat, month_offset=0)
cmv_m = sum_monthly_total(df_fat, COL_CMV, loja_fat, lojas_sel, mes_fat, month_offset=0)

# IMP (DTA.PAG com offset +1: DEZ vira JAN)
dedu_m = sum_monthly_conta_codigo(df_imp, COL_CONTA_IMP, COD_00004, COL_VAL_IMP, loja_imp, lojas_sel, mes_imp, month_offset=IMP_MONTH_OFFSET)
pessoal_m = sum_monthly_conta_codigo(df_imp, COL_CONTA_IMP, COD_00006, COL_VAL_IMP, loja_imp, lojas_sel, mes_imp, month_offset=IMP_MONTH_OFFSET)

margem_m = receita_m - cmv_m - dedu_m

# DRE (sem offset)
adm_m = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00007, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)
com_m = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00009, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)
fin_m = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00011, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)
oper_m = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00017, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)

resultado_m = margem_m - (pessoal_m + adm_m + com_m + fin_m + oper_m)
markup_m = receita_m / cmv_m.replace({0: np.nan})

# percentuais por mês
pct_receita_m = pd.Series(1.0, index=MESES)
pct_cmv_m = cmv_m / receita_m.replace({0: np.nan})
pct_dedu_m = dedu_m / receita_m.replace({0: np.nan})
pct_margem_m = margem_m / receita_m.replace({0: np.nan})
pct_pessoal_m = pessoal_m / receita_m.replace({0: np.nan})
pct_adm_m = adm_m / receita_m.replace({0: np.nan})
pct_com_m = com_m / receita_m.replace({0: np.nan})
pct_fin_m = fin_m / receita_m.replace({0: np.nan})
pct_oper_m = oper_m / receita_m.replace({0: np.nan})
pct_resultado_m = resultado_m / receita_m.replace({0: np.nan})

# ============================================================
# TABELA
# ============================================================
linhas = [
    "RECEITA",
    "CMV",
    "MARKUP",
    "DEDUÇÕES IMPOSTOS",
    "MARGEM BRUTA",
    "DESPESAS COM PESSOAL",
    "DESPESAS ADMINISTRATIVAS",
    "DESPESAS COMERCIAIS",
    "DESPESAS FINANCEIRAS",
    "DESPESAS OPERACIONAIS",
    "RESULTADO OPERACIONAL",
]

val_map = {
    "RECEITA": receita_m,
    "CMV": cmv_m,
    "MARKUP": markup_m,
    "DEDUÇÕES IMPOSTOS": dedu_m,
    "MARGEM BRUTA": margem_m,
    "DESPESAS COM PESSOAL": pessoal_m,
    "DESPESAS ADMINISTRATIVAS": adm_m,
    "DESPESAS COMERCIAIS": com_m,
    "DESPESAS FINANCEIRAS": fin_m,
    "DESPESAS OPERACIONAIS": oper_m,
    "RESULTADO OPERACIONAL": resultado_m,
}

pct_map = {
    "RECEITA": pct_receita_m,
    "CMV": pct_cmv_m,
    "MARKUP": pd.Series(np.nan, index=MESES),
    "DEDUÇÕES IMPOSTOS": pct_dedu_m,
    "MARGEM BRUTA": pct_margem_m,
    "DESPESAS COM PESSOAL": pct_pessoal_m,
    "DESPESAS ADMINISTRATIVAS": pct_adm_m,
    "DESPESAS COMERCIAIS": pct_com_m,
    "DESPESAS FINANCEIRAS": pct_fin_m,
    "DESPESAS OPERACIONAIS": pct_oper_m,
    "RESULTADO OPERACIONAL": pct_resultado_m,
}

cols = ["LINHA"]
for m in meses_ref:
    cols += [f"{m} (R$)", f"{m} (%)"]
cols += ["TOTAL (R$)", "TOTAL (%)"]

table = pd.DataFrame(columns=cols)
table["LINHA"] = linhas

for m in meses_ref:
    table[f"{m} (R$)"] = [val_map[l][m] for l in linhas]
    table[f"{m} (%)"] = [pct_map[l][m] for l in linhas]

# TOTAL (período)
total_receita = float(receita_m[meses_ref].sum())
total_cmv = float(cmv_m[meses_ref].sum())
total_dedu = float(dedu_m[meses_ref].sum())
total_mb = float(margem_m[meses_ref].sum())
total_pessoal = float(pessoal_m[meses_ref].sum())
total_adm = float(adm_m[meses_ref].sum())
total_com = float(com_m[meses_ref].sum())
total_fin = float(fin_m[meses_ref].sum())
total_oper = float(oper_m[meses_ref].sum())
total_result = float(resultado_m[meses_ref].sum())
total_markup = safe_div(total_receita, total_cmv)

total_pct_map = {
    "RECEITA": 1.0,
    "CMV": safe_div(total_cmv, total_receita),
    "MARKUP": np.nan,
    "DEDUÇÕES IMPOSTOS": safe_div(total_dedu, total_receita),
    "MARGEM BRUTA": safe_div(total_mb, total_receita),
    "DESPESAS COM PESSOAL": safe_div(total_pessoal, total_receita),
    "DESPESAS ADMINISTRATIVAS": safe_div(total_adm, total_receita),
    "DESPESAS COMERCIAIS": safe_div(total_com, total_receita),
    "DESPESAS FINANCEIRAS": safe_div(total_fin, total_receita),
    "DESPESAS OPERACIONAIS": safe_div(total_oper, total_receita),
    "RESULTADO OPERACIONAL": safe_div(total_result, total_receita),
}

total_val_map = {
    "RECEITA": total_receita,
    "CMV": total_cmv,
    "MARKUP": total_markup,
    "DEDUÇÕES IMPOSTOS": total_dedu,
    "MARGEM BRUTA": total_mb,
    "DESPESAS COM PESSOAL": total_pessoal,
    "DESPESAS ADMINISTRATIVAS": total_adm,
    "DESPESAS COMERCIAIS": total_com,
    "DESPESAS FINANCEIRAS": total_fin,
    "DESPESAS OPERACIONAIS": total_oper,
    "RESULTADO OPERACIONAL": total_result,
}

table["TOTAL (R$)"] = [total_val_map[l] for l in linhas]
table["TOTAL (%)"] = [total_pct_map[l] for l in linhas]

# Visual
view = table.copy()

for m in meses_ref:
    view[f"{m} (R$)"] = view.apply(
        lambda r: (
            "—" if pd.isna(r[f"{m} (R$)"]) else
            (
                f"{float(r[f'{m} (R$)']):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                if r["LINHA"] == "MARKUP"
                else f"R$ {format_brl(r[f'{m} (R$)'])}"
            )
        ),
        axis=1
    )
    view[f"{m} (%)"] = view.apply(
        lambda r: "—" if r["LINHA"] == "MARKUP" else format_pct(r[f"{m} (%)"]),
        axis=1
    )

view["TOTAL (R$)"] = view.apply(
    lambda r: (
        "—" if pd.isna(r["TOTAL (R$)"]) else
        (
            f"{float(r['TOTAL (R$)']):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if r["LINHA"] == "MARKUP"
            else f"R$ {format_brl(r['TOTAL (R$)'])}"
        )
    ),
    axis=1
)
view["TOTAL (%)"] = view.apply(
    lambda r: "—" if r["LINHA"] == "MARKUP" else format_pct(r["TOTAL (%)"]),
    axis=1
)

# ============================================================
# CARDS TOPO
# ============================================================
c1, c2, c3, c4 = st.columns(4)
c1.metric("Receita (Acumulado)", f"R$ {format_brl(total_receita)}")
c2.metric("CMV (Acumulado)", f"R$ {format_brl(total_cmv)}")
c3.metric("Margem Bruta (Acumulado)", f"R$ {format_brl(total_mb)}")

ro_pct_total = safe_div(total_result, total_receita)
c4.metric("Resultado Operacional (Acumulado)", f"R$ {format_brl(total_result)}", f"{format_pct(ro_pct_total)} s/ Receita")

cor_ro = "#0a3d91" if total_result >= 0 else "#c62828"
st.markdown(
    f"""
<div style="padding:10px 14px;border-radius:12px;border:1px solid rgba(0,0,0,.08);margin-top:8px;">
  <div style="font-size:13px;opacity:.75;">Resultado Operacional — Período: {periodo_label}</div>
  <div style="font-size:22px;font-weight:900;color:{cor_ro};">
    R$ {format_brl(total_result)} <span style="font-size:16px;font-weight:800;">({format_pct(ro_pct_total)})</span>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown("---")
st.subheader("DRE — Valores e Percentuais (Meses selecionados + Total)")

def style_view(df: pd.DataFrame):
    styles = pd.DataFrame("", index=df.index, columns=df.columns)

    mb_idx = df.index[df["LINHA"] == "MARGEM BRUTA"]
    for i in mb_idx:
        styles.loc[i, :] = "font-weight: 800;"

    ro_idx = df.index[df["LINHA"] == "RESULTADO OPERACIONAL"]
    if len(ro_idx) > 0:
        i = ro_idx[0]
        color_total = "#0a3d91" if total_result >= 0 else "#c62828"
        styles.loc[i, "TOTAL (R$)"] = f"font-weight: 900; color: {color_total};"
        styles.loc[i, "TOTAL (%)"] = f"font-weight: 900; color: {color_total};"
        for m in meses_ref:
            val = float(val_map["RESULTADO OPERACIONAL"][m])
            color = "#0a3d91" if val >= 0 else "#c62828"
            styles.loc[i, f"{m} (R$)"] = f"font-weight: 900; color: {color};"
            styles.loc[i, f"{m} (%)"] = f"font-weight: 900; color: {color};"

    return styles

st.dataframe(view.style.apply(style_view, axis=None), use_container_width=True, height=560, hide_index=True)

st.markdown("---")
st.subheader(f"Análise com Drilldown por Despesa — Período: {periodo_label}")

def conta_total_periodo(series: pd.Series) -> float:
    return float(series[meses_ref].sum())

def render_drill(titulo: str, cod: str, origem: str, valor_conta_periodo: float, month_offset: int):
    pct_conta_periodo = safe_div(valor_conta_periodo, total_receita)
    st.markdown(f"**{titulo}** — % sobre Receita (período): **{format_pct(pct_conta_periodo)}**")

    if origem == "IMP":
        agg, raw = drill_despesas_unique_codigo(
            df_imp, COL_CONTA_IMP, COL_DESPESA_IMP, COL_VAL_IMP,
            loja_imp, lojas_sel, cod,
            mes_imp, meses_ref, hist_imp,
            month_offset=month_offset,
        )
    else:
        agg, raw = drill_despesas_unique_codigo(
            df_dre, COL_CONTA_DRE, COL_DESPESA_DRE, COL_VAL_DRE,
            loja_dre, lojas_sel, cod,
            mes_dre, meses_ref, hist_dre,
            month_offset=month_offset,
        )

    if agg.empty:
        st.info("Sem despesas encontradas para essa conta com os filtros atuais.")
        return

    tv = agg.copy()
    tv["VALOR (R$)"] = tv["VALOR"].apply(lambda v: f"R$ {format_brl(v)}")
    tv["% sobre a conta"] = tv["%_SOBRE_CONTA"].apply(format_pct)
    tv = tv[["DESPESA", "VALOR (R$)", "% sobre a conta"]]

    left, right = st.columns([1.25, 1])

    with left:
        st.dataframe(tv, use_container_width=True, height=320, hide_index=True)

        desp_sel = st.selectbox(
            "Selecione uma despesa para ver histórico:",
            options=agg["DESPESA"].tolist(),
            key=f"sel_{origem}_{cod}_{'_'.join(meses_ref)}"
        )

        st.markdown("**Histórico da despesa (por loja)**")
        if raw.empty:
            st.info("Nenhum histórico encontrado (coluna de histórico ausente/vazia).")
        else:
            r = raw[raw["DESPESA"] == desp_sel].copy()
            if r.empty:
                st.info("Sem histórico para essa despesa.")
            else:
                for loja in sorted(r["LOJA"].unique().tolist()):
                    st.markdown(f"- **{loja}**")
                    for h in r[r["LOJA"] == loja]["HIST"].tolist():
                        st.write(f"  • {h}")

    with right:
        fig = px.pie(agg, values="VALOR", names="DESPESA", title="Distribuição das despesas (sobre a conta)")
        st.plotly_chart(fig, use_container_width=True)

with st.expander("DEDUÇÕES IMPOSTOS", expanded=False):
    render_drill("DEDUÇÕES IMPOSTOS", COD_00004, "IMP", conta_total_periodo(dedu_m), month_offset=IMP_MONTH_OFFSET)

with st.expander("DESPESAS COM PESSOAL", expanded=False):
    render_drill("DESPESAS COM PESSOAL", COD_00006, "IMP", conta_total_periodo(pessoal_m), month_offset=IMP_MONTH_OFFSET)

with st.expander("DESPESAS ADMINISTRATIVAS", expanded=False):
    render_drill("DESPESAS ADMINISTRATIVAS", COD_00007, "DRE", conta_total_periodo(adm_m), month_offset=0)

with st.expander("DESPESAS COMERCIAIS", expanded=False):
    render_drill("DESPESAS COMERCIAIS", COD_00009, "DRE", conta_total_periodo(com_m), month_offset=0)

with st.expander("DESPESAS FINANCEIRAS", expanded=False):
    render_drill("DESPESAS FINANCEIRAS", COD_00011, "DRE", conta_total_periodo(fin_m), month_offset=0)

with st.expander("DESPESAS OPERACIONAIS", expanded=False):
    render_drill("DESPESAS OPERACIONAIS", COD_00017, "DRE", conta_total_periodo(oper_m), month_offset=0)

with st.expander("Diagnóstico (competência IMPOSTOS/FOLHA)", expanded=False):
    st.write({
        "IMPOSTOS E FOLHA -> coluna de data": COL_DATA_IMP,
        "Offset aplicado (mês anterior apropriado no mês vigente)": IMP_MONTH_OFFSET,
        "Mês (FAT)": mes_fat,
        "Mês (IMP fixo)": mes_imp,
        "Mês (DRE)": mes_dre,
        "LOJA (FAT)": loja_fat,
        "LOJA (IMP)": loja_imp,
        "LOJA (DRE)": loja_dre,
        "Meses selecionados": meses_ref,
    })


# ============================================================
# >>> NOVO BLOCO: DFC (COMPLETO)
# ============================================================
st.markdown("---")
st.title("DFC DAUTO TINTAS")

# ------------------------------------------------------------
# Config DFC
# ------------------------------------------------------------
DFC_ANO_BASE = 2026  # Dashboard iniciando em 2026

DFC_LOJA_ORDER = [
    "ADE",
    "GAMA",
    "SOFNORTE",
    "CEILÂNDIA",
    "S IA",
    "UNAÍ",
    "AG LINDAS",
    "GUARÁ",
    "LUZIÂNIA",
]

# ------------------------------------------------------------
# Helpers locais DFC
# ------------------------------------------------------------
def _strip_accents_upper(s: str) -> str:
    try:
        import unicodedata
        s = unicodedata.normalize("NFKD", str(s))
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
    except Exception:
        s = str(s)
    return re.sub(r"\s+", " ", s).strip().upper()

def _canon_loja_name(col: str) -> str:
    u = _strip_accents_upper(col)
    u = u.replace(".", " ").replace("-", " ")
    u = re.sub(r"\s+", " ", u).strip()

    if u in {"SIA", "S I A"}:
        return "S IA"
    if u in {"AGUAS LINDAS", "AGUASLINDAS", "AG LINDAS", "AGLINDAS", "AG. LINDAS"}:
        return "AG LINDAS"
    if u in {"CEILANDIA"}:
        return "CEILÂNDIA"
    if u in {"GUARA"}:
        return "GUARÁ"
    if u in {"UNAI"}:
        return "UNAÍ"
    if u in {"LUZIANIA"}:
        return "LUZIÂNIA"

    for x in DFC_LOJA_ORDER:
        if _strip_accents_upper(x) == u:
            return x

    return u

_MESES_PT_FULL_TO_NUM = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}
_MESES_PT_ABBR_TO_NUM = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12
}

def _parse_mes_ano(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None

    dt = pd.to_datetime(v, errors="coerce", dayfirst=True)
    if pd.notna(dt):
        return (int(dt.year), int(dt.month))

    s = str(v).strip()
    if s == "":
        return None
    su = _strip_accents_upper(s)

    m = re.search(r"\b([A-Z]{3})\s*/\s*(\d{2,4})\b", su)
    if m:
        mon = m.group(1)
        yy = m.group(2)
        if mon in _MESES_PT_ABBR_TO_NUM:
            mm = _MESES_PT_ABBR_TO_NUM[mon]
            aa = 2000 + int(yy) if len(yy) == 2 else int(yy)
            return (aa, mm)

    m2 = re.search(r"\b([A-ZÇ]{4,9})\s*/\s*(\d{2,4})\b", su)
    if m2:
        mon = m2.group(1)
        yy = m2.group(2)
        if mon in _MESES_PT_FULL_TO_NUM:
            mm = _MESES_PT_FULL_TO_NUM[mon]
            aa = 2000 + int(yy) if len(yy) == 2 else int(yy)
            return (aa, mm)

    if su in _MESES_PT_ABBR_TO_NUM:
        return (DFC_ANO_BASE, _MESES_PT_ABBR_TO_NUM[su])
    if su in _MESES_PT_FULL_TO_NUM:
        return (DFC_ANO_BASE, _MESES_PT_FULL_TO_NUM[su])

    for abbr, mm in _MESES_PT_ABBR_TO_NUM.items():
        if re.search(rf"\b{abbr}\b", su):
            y = re.search(r"\b(20\d{2}|\d{2})\b", su)
            if y:
                yy = y.group(1)
                aa = 2000 + int(yy) if len(yy) == 2 else int(yy)
            else:
                aa = DFC_ANO_BASE
            return (aa, mm)

    return None

def _target_month_tuples_from_meses_ref(meses_ref_list):
    out = []
    for m in meses_ref_list:
        mu = _strip_accents_upper(m)[:3]
        if mu in _MESES_PT_ABBR_TO_NUM:
            out.append((DFC_ANO_BASE, _MESES_PT_ABBR_TO_NUM[mu]))
    return out

def _shift_month_tuple(ano, mes, offset):
    idx = (ano * 12 + (mes - 1)) + offset
    aa = idx // 12
    mm = (idx % 12) + 1
    return (aa, mm)

def _mes_tuple_to_label(ano, mes):
    abbr = MESES[mes - 1]
    yy = str(ano)[-2:]
    return f"{abbr}/{yy}"

def _coerce_money_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s.astype(float).fillna(0.0)
    tmp = s.astype(str).str.strip()
    tmp = tmp.replace({"-": "", "—": "", "nan": "", "None": ""})
    tmp = tmp.str.replace("R$", "", regex=False).str.replace("r$", "", regex=False)
    tmp = tmp.str.replace("\u00a0", " ", regex=False).str.replace(" ", "", regex=False)
    tmp = tmp.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(tmp, errors="coerce").fillna(0.0)

def _read_wide_sheet(sheet_name: str, xls: pd.ExcelFile) -> pd.DataFrame:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _build_wide_table_and_totals(
    df_wide: pd.DataFrame,
    lojas_sel_list: list[str],
    target_months: list[tuple],
    month_offset_lookup: int,
    loja_order: list[str]
):
    if df_wide is None or df_wide.empty:
        return pd.DataFrame(), {}, 0.0, []

    x = df_wide.copy()
    col_mes = x.columns[0]
    x["MES_TUP__"] = x[col_mes].apply(_parse_mes_ano)
    x = x[x["MES_TUP__"].notna()].copy()
    if x.empty:
        return pd.DataFrame(), {}, 0.0, []

    loja_cols_raw = [c for c in x.columns if c not in {col_mes, "MES_TUP__"}]
    loja_cols_raw = [c for c in loja_cols_raw if _strip_accents_upper(c) not in {"TOTAL", "TOTAL GERAL"}]

    rename_map = {c: _canon_loja_name(c) for c in loja_cols_raw}
    x = x.rename(columns=rename_map)
    loja_cols = [rename_map[c] for c in loja_cols_raw]

    lojas_sel_set = set([str(l).strip() for l in lojas_sel_list]) if lojas_sel_list else set()
    loja_cols_use = [c for c in loja_cols if c in lojas_sel_set] if lojas_sel_set else loja_cols[:]
    if not loja_cols_use:
        loja_cols_use = loja_cols[:]

    ordered = [c for c in loja_order if c in loja_cols_use]
    extras = [c for c in loja_cols_use if c not in ordered]
    lojas_use = ordered + extras

    for c in lojas_use:
        x[c] = _coerce_money_series(x[c])

    by_month = {}
    for (aa, mm), grp in x.groupby("MES_TUP__", dropna=True):
        by_month[(int(aa), int(mm))] = grp[lojas_use].sum(axis=0)

    rows = []
    totals_by_label = {}
    total_periodo = 0.0

    for (aa, mm) in target_months:
        lk = _shift_month_tuple(aa, mm, month_offset_lookup)
        series_vals = by_month.get(lk, pd.Series(0.0, index=lojas_use))
        linha_total = float(series_vals.sum())
        lbl = _mes_tuple_to_label(aa, mm)

        totals_by_label[lbl] = linha_total
        total_periodo += linha_total

        row = {"MÊS": lbl}
        for c in lojas_use:
            row[c] = float(series_vals.get(c, 0.0))
        row["TOTAL"] = linha_total
        rows.append(row)

    df_out = pd.DataFrame(rows)

    df_view = df_out.copy()
    for c in [c for c in df_view.columns if c != "MÊS"]:
        df_view[c] = df_view[c].apply(lambda v: f"R$ {format_brl(v)}")

    return df_view, totals_by_label, total_periodo, lojas_use


# ------------------------------------------------------------
# Estilo (cores e negrito)
# ------------------------------------------------------------
def _style_pos_neg(val):
    try:
        v = float(val)
    except Exception:
        return ""
    if v < 0:
        return "color: #c62828;"  # vermelho
    if v > 0:
        return "color: #0a3d91;"  # azul
    return "color: #555;"

def _style_bold_resultado(row):
    if str(row.get("LINHA", "")).strip().upper() == "RESULTADO CAIXA":
        return ["font-weight: 700;"] * len(row)
    return [""] * len(row)


# ------------------------------------------------------------
# Meses selecionados (do sidebar do DRE)
# ------------------------------------------------------------
target_months = _target_month_tuples_from_meses_ref(meses_ref)

if not target_months:
    st.warning("Selecione ao menos 1 mês no filtro para visualizar o DFC.")
else:
    # --------------------------------------------------------
    # 1) RECEBIMENTOS (Drill) + Totais
    # --------------------------------------------------------
    try:
        df_receb_wide = _read_wide_sheet("RECEBIMENTOS", EXCEL_XLS)
    except Exception as e:
        st.error(f"Erro ao ler a aba RECEBIMENTOS: {e}")
        df_receb_wide = None

    receb_view, receb_totals_by_lbl, receb_total_periodo, _ = _build_wide_table_and_totals(
        df_wide=df_receb_wide,
        lojas_sel_list=lojas_sel,
        target_months=target_months,
        month_offset_lookup=0,
        loja_order=DFC_LOJA_ORDER
    )

    c1, c2 = st.columns([1, 2])
    with c1:
        st.metric("Recebimentos (Total do período)", f"R$ {format_brl(receb_total_periodo)}")
    with c2:
        st.caption("Recebimentos puxados da aba RECEBIMENTOS e somados por loja.")

    with st.expander("Drill — RECEBIMENTOS (Mês × Loja)", expanded=False):
        if receb_view is None or receb_view.empty:
            st.info("Sem dados para exibir em RECEBIMENTOS com os filtros atuais.")
        else:
            st.dataframe(receb_view, use_container_width=True, hide_index=True, height=420)

    # --------------------------------------------------------
    # 2) COMPRAS LÍQUIDAS (Drill) = COMPRAS - DEVOLUÇÕES (competência M-1)
    # --------------------------------------------------------
    try:
        df_compras_wide = _read_wide_sheet("COMPRAS", EXCEL_XLS)
    except Exception as e:
        st.error(f"Erro ao ler a aba COMPRAS: {e}")
        df_compras_wide = None

    try:
        df_devol_wide = _read_wide_sheet("DEVOLUÇÕES", EXCEL_XLS)
    except Exception as e:
        st.error(f"Erro ao ler a aba DEVOLUÇÕES: {e}")
        df_devol_wide = None

    compras_view, compras_totals_by_lbl, _, _ = _build_wide_table_and_totals(
        df_wide=df_compras_wide,
        lojas_sel_list=lojas_sel,
        target_months=target_months,
        month_offset_lookup=-1,
        loja_order=DFC_LOJA_ORDER
    )

    devol_view, devol_totals_by_lbl, _, _ = _build_wide_table_and_totals(
        df_wide=df_devol_wide,
        lojas_sel_list=lojas_sel,
        target_months=target_months,
        month_offset_lookup=-1,
        loja_order=DFC_LOJA_ORDER
    )

    compras_liq_totals_by_lbl = {}
    for (aa, mm) in target_months:
        lbl = _mes_tuple_to_label(aa, mm)
        compras_liq_totals_by_lbl[lbl] = float(compras_totals_by_lbl.get(lbl, 0.0)) - float(devol_totals_by_lbl.get(lbl, 0.0))

    compras_liq_total_periodo = float(sum(compras_liq_totals_by_lbl.values()))

    c3, c4 = st.columns([1, 2])
    with c3:
        st.metric("Compras Líquidas (Total do período)", f"R$ {format_brl(compras_liq_total_periodo)}")
    with c4:
        st.caption("Compras Líquidas = COMPRAS − DEVOLUÇÕES (competência: mês anterior).")

    with st.expander("Drill — COMPRAS (líquido) = COMPRAS − DEVOLUÇÕES (Competência: mês anterior)", expanded=False):
        def _brl_to_float(s):
            if s is None or (isinstance(s, float) and np.isnan(s)):
                return 0.0
            t = str(s).strip().replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            try:
                return float(t)
            except Exception:
                return 0.0

        cols_liq = ["MÊS"] + [c for c in DFC_LOJA_ORDER if (compras_view is not None and c in compras_view.columns) or (devol_view is not None and c in devol_view.columns)] + ["TOTAL"]

        cv = compras_view.copy() if compras_view is not None and not compras_view.empty else pd.DataFrame(columns=cols_liq)
        dv = devol_view.copy() if devol_view is not None and not devol_view.empty else pd.DataFrame(columns=cols_liq)

        for c in cols_liq:
            if c not in cv.columns:
                cv[c] = "R$ 0,00"
            if c not in dv.columns:
                dv[c] = "R$ 0,00"
        cv = cv[cols_liq]
        dv = dv[cols_liq]

        liq = cv.copy()
        for c in cols_liq:
            if c == "MÊS":
                continue
            liq[c] = cv[c].apply(_brl_to_float) - dv[c].apply(_brl_to_float)

        liq_view = liq.copy()
        for c in [c for c in liq_view.columns if c != "MÊS"]:
            liq_view[c] = liq_view[c].apply(lambda v: f"R$ {format_brl(v)}")

        st.dataframe(liq_view, use_container_width=True, hide_index=True, height=420)

    # --------------------------------------------------------
    # 3) TABELA GERAL DO DFC (layout igual ao DRE)
    # --------------------------------------------------------
    st.markdown("---")
    st.subheader("DFC — Tabela Geral (Valores e Percentuais sobre Recebimento)")

    if "df_dre" not in globals() or df_dre is None or df_dre.empty:
        st.error("A aba DRE (df_dre) não está carregada. O DFC precisa dela.")
    else:
        dedu_dfc_m = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00004, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)
        pess_dfc_m = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00006, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)
        adm_dfc_m  = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00007, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)
        com_dfc_m  = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00009, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)
        fin_dfc_m  = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00011, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)
        ope_dfc_m  = sum_monthly_conta_codigo(df_dre, COL_CONTA_DRE, COD_00017, COL_VAL_DRE, loja_dre, lojas_sel, mes_dre, month_offset=0)

        receb_m = pd.Series(0.0, index=MESES)
        forn_m  = pd.Series(0.0, index=MESES)

        for (aa, mm) in target_months:
            abbr = MESES[mm - 1]
            lbl = _mes_tuple_to_label(aa, mm)
            receb_m[abbr] = float(receb_totals_by_lbl.get(lbl, 0.0))
            forn_m[abbr]  = float(compras_liq_totals_by_lbl.get(lbl, 0.0))

        resultado_caixa_m = receb_m - (forn_m + dedu_dfc_m + pess_dfc_m + adm_dfc_m + com_dfc_m + fin_dfc_m + ope_dfc_m)

        denom = receb_m.replace({0: np.nan})
        pct_receb_m = pd.Series(1.0, index=MESES)
        pct_forn_m  = forn_m / denom
        pct_dedu_m  = dedu_dfc_m / denom
        pct_pess_m  = pess_dfc_m / denom
        pct_adm_m   = adm_dfc_m / denom
        pct_com_m   = com_dfc_m / denom
        pct_fin_m   = fin_dfc_m / denom
        pct_ope_m   = ope_dfc_m / denom
        pct_res_m   = resultado_caixa_m / denom

        linhas_dfc = [
            "RECEBIMENTO",
            "FORNECEDOR (COMPRAS LÍQUIDAS)",
            "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)",
            "DESPESAS COM PESSOAL",
            "DESPESAS ADMINISTRATIVAS",
            "DESPESAS COMERCIAIS",
            "DESPESAS FINANCEIRAS",
            "DESPESAS OPERACIONAIS",
            "RESULTADO CAIXA",
        ]

        val_map_dfc = {
            "RECEBIMENTO": receb_m,
            "FORNECEDOR (COMPRAS LÍQUIDAS)": forn_m,
            "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": dedu_dfc_m,
            "DESPESAS COM PESSOAL": pess_dfc_m,
            "DESPESAS ADMINISTRATIVAS": adm_dfc_m,
            "DESPESAS COMERCIAIS": com_dfc_m,
            "DESPESAS FINANCEIRAS": fin_dfc_m,
            "DESPESAS OPERACIONAIS": ope_dfc_m,
            "RESULTADO CAIXA": resultado_caixa_m,
        }

        pct_map_dfc = {
            "RECEBIMENTO": pct_receb_m,
            "FORNECEDOR (COMPRAS LÍQUIDAS)": pct_forn_m,
            "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": pct_dedu_m,
            "DESPESAS COM PESSOAL": pct_pess_m,
            "DESPESAS ADMINISTRATIVAS": pct_adm_m,
            "DESPESAS COMERCIAIS": pct_com_m,
            "DESPESAS FINANCEIRAS": pct_fin_m,
            "DESPESAS OPERACIONAIS": pct_ope_m,
            "RESULTADO CAIXA": pct_res_m,
        }

        cols = ["LINHA"]
        for m in meses_ref:
            cols += [f"{m} (R$)", f"{m} (%)"]
        cols += ["TOTAL (R$)", "TOTAL (%)"]

        table_dfc = pd.DataFrame(columns=cols)
        table_dfc["LINHA"] = linhas_dfc

        for m in meses_ref:
            table_dfc[f"{m} (R$)"] = [val_map_dfc[l][m] for l in linhas_dfc]
            table_dfc[f"{m} (%)"] = [pct_map_dfc[l][m] for l in linhas_dfc]

        total_receb = float(receb_m[meses_ref].sum())
        total_forn  = float(forn_m[meses_ref].sum())
        total_dedu  = float(dedu_dfc_m[meses_ref].sum())
        total_pess  = float(pess_dfc_m[meses_ref].sum())
        total_adm   = float(adm_dfc_m[meses_ref].sum())
        total_com   = float(com_dfc_m[meses_ref].sum())
        total_fin   = float(fin_dfc_m[meses_ref].sum())
        total_ope   = float(ope_dfc_m[meses_ref].sum())
        total_res   = float(resultado_caixa_m[meses_ref].sum())

        table_dfc["TOTAL (R$)"] = [
            total_receb, total_forn, total_dedu, total_pess, total_adm, total_com, total_fin, total_ope, total_res
        ]

        def _pct_over_receb(v):
            return safe_div(v, total_receb)

        table_dfc["TOTAL (%)"] = [
            1.0 if total_receb != 0 else np.nan,
            _pct_over_receb(total_forn),
            _pct_over_receb(total_dedu),
            _pct_over_receb(total_pess),
            _pct_over_receb(total_adm),
            _pct_over_receb(total_com),
            _pct_over_receb(total_fin),
            _pct_over_receb(total_ope),
            _pct_over_receb(total_res),
        ]

        # Cards topo
        k1, k2, k3 = st.columns(3)
        k1.metric("Recebimento (Período)", f"R$ {format_brl(total_receb)}")
        k2.metric("Fornecedor (Período)", f"R$ {format_brl(total_forn)}", f"{format_pct(_pct_over_receb(total_forn))} s/ receb.")
        k3.metric("Resultado Caixa (Período)", f"R$ {format_brl(total_res)}", f"{format_pct(_pct_over_receb(total_res))} s/ receb.")

             # --------- STYLER (apenas Resultado Caixa colorido + negrito) ---------
        value_cols = [c for c in table_dfc.columns if "(R$)" in c] + ["TOTAL (R$)"]
        pct_cols = [c for c in table_dfc.columns if "(%)" in c] + ["TOTAL (%)"]

        def _style_resultado_caixa_colors(row):
            """
            Aplica cor SOMENTE na linha RESULTADO CAIXA e SOMENTE nas colunas de valor (R$).
            """
            is_res = str(row.get("LINHA", "")).strip().upper() == "RESULTADO CAIXA"
            styles = [""] * len(row)

            if not is_res:
                return styles

            # mapeia índice de colunas para aplicar estilo
            col_index = {col: i for i, col in enumerate(row.index)}

            for c in value_cols:
                if c in col_index:
                    try:
                        v = float(row[c])
                    except Exception:
                        v = 0.0

                    if v < 0:
                        styles[col_index[c]] = "color: #c62828;"  # vermelho
                    elif v > 0:
                        styles[col_index[c]] = "color: #0a3d91;"  # azul
                    else:
                        styles[col_index[c]] = "color: #555;"

            # negrito em toda a linha Resultado Caixa (inclui % também)
            styles = [s + " font-weight: 700;" if s else "font-weight: 700;" for s in styles]
            return styles

        table_dfc = table_dfc.reset_index(drop=True)

        sty = table_dfc.style

        # aplica cor apenas no Resultado Caixa (colunas R$) + negrito na linha inteira
        sty = sty.apply(_style_resultado_caixa_colors, axis=1)

        # formatação BRL e %
        fmt_map = {}
        for c in value_cols:
            fmt_map[c] = lambda v: f"R$ {format_brl(v)}"
        for c in pct_cols:
            fmt_map[c] = lambda v: format_pct(v)

        sty = sty.format(fmt_map)

        st.dataframe(sty, use_container_width=True, height=520, hide_index=True)


        # ----------------------------------------------------
        # 4) DRILL DFC — (mesmo do DRE) com keys únicas
        # ----------------------------------------------------
        st.markdown("---")
        st.subheader("DFC — Drill por Contas de Resultado)")

        def _conta_total_periodo(series: pd.Series) -> float:
            return float(series[meses_ref].sum())

        def _render_drill_dfc(titulo: str, cod: str, valor_conta_periodo: float):
            pct_conta_periodo = safe_div(valor_conta_periodo, total_receb)
            st.markdown(f"**{titulo}** — % sobre Recebimento (período): **{format_pct(pct_conta_periodo)}**")

            agg, raw = drill_despesas_unique_codigo(
                df_dre,
                COL_CONTA_DRE,
                COL_DESPESA_DRE,
                COL_VAL_DRE,
                loja_dre,
                lojas_sel,
                cod,
                mes_dre,
                meses_ref,
                hist_dre,
                month_offset=0,
            )

            if agg.empty:
                st.info("Sem despesas encontradas para essa conta com os filtros atuais.")
                return

            tv = agg.copy()
            tv["VALOR (R$)"] = tv["VALOR"].apply(lambda v: f"R$ {format_brl(v)}")
            tv["% sobre a conta"] = tv["%_SOBRE_CONTA"].apply(format_pct)
            tv = tv[["DESPESA", "VALOR (R$)", "% sobre a conta"]]

            left, right = st.columns([1.25, 1])

            with left:
                st.dataframe(tv, use_container_width=True, height=340, hide_index=True)

                desp_sel = st.selectbox(
                    "Selecione uma despesa para ver histórico:",
                    options=agg["DESPESA"].tolist(),
                    key=f"dfc_selbox_{cod}_{'_'.join(meses_ref)}"
                )

                st.markdown("**Histórico da despesa (por loja)**")
                if raw.empty:
                    st.info("Nenhum histórico encontrado (coluna HISTÓRICO ausente/vazia).")
                else:
                    r = raw[raw["DESPESA"] == desp_sel].copy()
                    if r.empty:
                        st.info("Sem histórico para essa despesa.")
                    else:
                        for loja in sorted(r["LOJA"].unique().tolist()):
                            st.markdown(f"- **{loja}**")
                            for h in r[r["LOJA"] == loja]["HIST"].tolist():
                                st.write(f"  • {h}")

            with right:
                fig = px.pie(agg, values="VALOR", names="DESPESA", title="Distribuição das despesas (sobre a conta)")
                st.plotly_chart(fig, use_container_width=True, key=f"dfc_pie_{cod}_{'_'.join(meses_ref)}")

        with st.expander("00004 — DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", expanded=False):
            _render_drill_dfc("00004 — DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", COD_00004, _conta_total_periodo(dedu_dfc_m))

        with st.expander("00006 — DESPESAS COM PESSOAL", expanded=False):
            _render_drill_dfc("00006 — DESPESAS COM PESSOAL", COD_00006, _conta_total_periodo(pess_dfc_m))

        with st.expander("00007 — DESPESAS ADMINISTRATIVAS", expanded=False):
            _render_drill_dfc("00007 — DESPESAS ADMINISTRATIVAS", COD_00007, _conta_total_periodo(adm_dfc_m))

        with st.expander("00009 — DESPESAS COMERCIAIS", expanded=False):
            _render_drill_dfc("00009 — DESPESAS COMERCIAIS", COD_00009, _conta_total_periodo(com_dfc_m))

        with st.expander("00011 — DESPESAS FINANCEIRAS", expanded=False):
            _render_drill_dfc("00011 — DESPESAS FINANCEIRAS", COD_00011, _conta_total_periodo(fin_dfc_m))

        with st.expander("00017 — DESPESAS OPERACIONAIS", expanded=False):
            _render_drill_dfc("00017 — DESPESAS OPERACIONAIS", COD_00017, _conta_total_periodo(ope_dfc_m))

        with st.expander("Diagnóstico (DFC)", expanded=False):
            st.write({
                "Meses selecionados (meses_ref)": meses_ref,
                "Target months (ano,mes)": target_months,
                "Recebimento total período": total_receb,
                "Fornecedor total período (compras líquidas)": total_forn,
            })
