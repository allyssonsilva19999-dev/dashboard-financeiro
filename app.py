import math
import sqlite3
import textwrap
from datetime import date
from html import escape
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

# ====================== CONFIG ======================
st.set_page_config(
    page_title="Dashboard Financeiro",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="auto",
)

DB_FILE = "financeiro.db"


def get_conn():
    """Conexao SQLite resiliente."""
    conn = sqlite3.connect(DB_FILE, timeout=30, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=DELETE")
    conn.execute("PRAGMA busy_timeout=30000")
    return conn


# ====================== UTILITÁRIOS ======================


def brl(valor) -> str:
    try:
        if valor is None:
            numero = 0.0
        else:
            numero = float(valor)
        if numero != numero:  # NaN check
            numero = 0.0
    except Exception:
        numero = 0.0
    texto = f"R$ {numero:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def brl_curto(valor) -> str:
    try:
        numero = float(valor or 0)
    except (TypeError, ValueError):
        numero = 0.0
    absoluto = abs(numero)
    if absoluto >= 1_000_000:
        texto = f"R$ {numero / 1_000_000:.1f} mi"
    elif absoluto >= 1_000:
        texto = f"R$ {numero / 1_000:.1f} mil"
    else:
        texto = f"R$ {numero:.0f}"
    return texto.replace(".", ",")


def pct(valor, casas: int = 0) -> str:
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        numero = 0.0
    return f"{numero:.{casas}f}%".replace(".", ",")


def data_br(valor) -> str:
    if valor is None:
        return "Sem data"
    try:
        texto = str(valor).strip()
        # ISO dates (YYYY-MM-DD) from SQLite — no dayfirst needed
        if len(texto) >= 10 and texto[4] == "-" and texto[7] == "-":
            data_convertida = pd.to_datetime(texto[:10], errors="coerce", format="%Y-%m-%d")
        else:
            data_convertida = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    except Exception:
        return "Sem data"
    try:
        if data_convertida is None or pd.isna(data_convertida):
            return "Sem data"
    except Exception:
        return "Sem data"
    return data_convertida.strftime("%d/%m/%Y")


def limpar_texto(valor, padrao: str = "") -> str:
    if valor is None:
        return padrao
    try:
        # Avoid pd.isna on exotic objects that raise
        if isinstance(valor, float) and math.isnan(valor):
            return padrao
        if isinstance(valor, (pd.Series, pd.DataFrame)):
            return padrao
        # Scalar check without exploding
        try:
            if valor is pd.NA or valor is pd.NaT:
                return padrao
        except Exception:
            pass
        if str(valor).lower() in ("nan", "nat", "none", "<na>"):
            return padrao
    except Exception:
        return padrao
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return padrao
    return texto


def normalizar(valor) -> str:
    texto = limpar_texto(valor).lower()
    trocas = {
        "á": "a", "à": "a", "ã": "a", "â": "a",
        "é": "e", "ê": "e", "í": "i",
        "ó": "o", "ô": "o", "õ": "o",
        "ú": "u", "ç": "c",
    }
    for origem, destino in trocas.items():
        texto = texto.replace(origem, destino)
    return " ".join(texto.replace("_", " ").replace("-", " ").split())


def texto_meses(meses: int) -> str:
    if meses <= 0:
        return "agora"
    anos = meses // 12
    meses_restantes = meses % 12
    partes = []
    if anos:
        partes.append(f"{anos} ano" if anos == 1 else f"{anos} anos")
    if meses_restantes:
        partes.append(f"{meses_restantes} mês" if meses_restantes == 1 else f"{meses_restantes} meses")
    return " e ".join(partes) if partes else "menos de 1 mês"


def limitar_percentual(valor) -> float:
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        numero = 0.0
    return max(0.0, min(numero, 100.0))


def mensagem_erro_usuario(erro) -> str:
    texto = str(erro or "").strip()
    if not texto or len(texto) > 180:
        return "Não conseguimos concluir agora. Tente novamente."
    return texto


# ====================== ESTILO (mix Influency + Lodgify + Mobile) ======================
st.markdown(
    """
<style>
    :root {
        --ink: #1c1f26;
        --muted: #8a90a0;
        --lime: #d9ff00;
        --lime-soft: #f3ff9a;
        --mint: #b8f0d8;
        --mint-soft: #e8faf3;
        --mint-mid: #7ed9b0;
        --blue: #a8c4ff;
        --blue-soft: #eaf0ff;
        --bg: #f3f4f7;
        --card: #ffffff;
        --sidebar: #ffffff;
        --shadow: 0 12px 40px rgba(28,31,38,0.07);
        --shadow-soft: 0 6px 18px rgba(28,31,38,0.05);
        --radius: 20px;
        --radius-sm: 14px;
        --radius-lg: 24px;
    }

    html, body, [class*="css"] {
        font-family: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        -webkit-text-size-adjust: 100%;
        -webkit-tap-highlight-color: transparent;
    }

    [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(circle at 0% 0%, rgba(184,240,216,0.35), transparent 28rem),
            radial-gradient(circle at 100% 0%, rgba(217,255,0,0.12), transparent 22rem),
            radial-gradient(circle at 90% 90%, rgba(168,196,255,0.18), transparent 26rem),
            var(--bg) !important;
        color: var(--ink);
    }

    [data-testid="stHeader"],
    [data-testid="stToolbar"] {
        background: transparent !important;
    }

    .main .block-container {
        max-width: 1200px;
        padding-top: 1.25rem;
        padding-bottom: 2.5rem;
        padding-left: 1.4rem;
        padding-right: 1.4rem;
    }

    /* ===== SIDEBAR ===== */
    [data-testid="stSidebar"] {
        background: var(--sidebar) !important;
        border-right: 1px solid rgba(28,31,38,0.06);
        box-shadow: 4px 0 24px rgba(28,31,38,0.03);
    }
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 1.2rem;
    }
    section[data-testid="stSidebar"] .stMarkdown h1,
    section[data-testid="stSidebar"] .stMarkdown h2,
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: var(--ink) !important;
    }
    [data-testid="stSidebar"] .stRadio > label {
        display: none !important;
    }
    [data-testid="stSidebar"] .stRadio [role="radiogroup"] label {
        padding: 0.55rem 0.95rem !important;
        border-radius: 12px !important;
        font-weight: 650 !important;
        color: #6b7280 !important;
        margin-bottom: 0.25rem !important;
        transition: 0.15s ease;
        min-height: 2.6rem;
    }
    [data-testid="stSidebar"] .stRadio [role="radiogroup"] label:hover {
        background: #f6f7f9 !important;
        color: var(--ink) !important;
    }
    [data-testid="stSidebar"] .stRadio [role="radiogroup"] label[data-checked="true"],
    [data-testid="stSidebar"] .stRadio [aria-checked="true"] {
        background: var(--lime) !important;
        color: var(--ink) !important;
        font-weight: 800 !important;
        box-shadow: 0 4px 14px rgba(217,255,0,0.25);
    }

    h1, h2, h3 {
        color: var(--ink) !important;
        letter-spacing: -0.025em;
        font-weight: 800 !important;
        word-wrap: break-word;
        overflow-wrap: break-word;
    }

    /* ===== GRIDS ===== */
    .metric-grid,
    .indicator-grid,
    .answer-grid,
    .goal-grid,
    .debt-grid,
    .investment-grid {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.95rem;
        margin: 1rem 0 1.2rem;
    }

    /* ===== CARDS ===== */
    .metric-card,
    .indicator-card,
    .answer-card,
    .goal-card,
    .history-item,
    .investment-item,
    .debt-item {
        position: relative;
        overflow: hidden;
        min-height: 7.8rem;
        padding: 1.15rem 1.2rem;
        border-radius: var(--radius);
        background: var(--card);
        border: 1px solid rgba(28,31,38,0.04);
        box-shadow: var(--shadow-soft);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }

    .metric-card:hover,
    .indicator-card:hover {
        transform: translateY(-2px);
        box-shadow: var(--shadow);
    }

    .metric-card.accent {
        background: linear-gradient(145deg, #d9ff00 0%, #e8ff4a 100%);
        border: none;
        box-shadow: 0 10px 28px rgba(217,255,0,0.28);
    }
    .metric-card.accent .metric-label,
    .metric-card.accent .metric-foot,
    .metric-card.accent .metric-value {
        color: var(--ink) !important;
    }

    .metric-card.mint {
        background: linear-gradient(145deg, #e8faf3 0%, #d4f5e8 100%);
        border: 1px solid rgba(126,217,176,0.25);
    }

    .metric-card.blue {
        background: var(--card);
    }

    .metric-card.dark {
        background: linear-gradient(145deg, #1c1f26 0%, #2a2e38 100%);
        color: white;
        border: none;
    }
    .metric-card.dark .metric-label,
    .metric-card.dark .metric-foot,
    .metric-card.dark .metric-value {
        color: white !important;
    }

    .metric-label,
    .indicator-top,
    .answer-question {
        color: var(--muted);
        font-size: 0.76rem;
        font-weight: 700;
        letter-spacing: 0.01em;
    }

    .metric-value,
    .indicator-value,
    .answer-value {
        margin-top: 0.7rem;
        color: var(--ink);
        font-size: clamp(1.25rem, 2.4vw, 1.9rem);
        line-height: 1.15;
        font-weight: 800;
        letter-spacing: -0.03em;
        word-break: break-word;
    }

    .metric-foot,
    .indicator-note,
    .answer-action,
    .goal-meta,
    .debt-meta,
    .investment-meta {
        margin-top: 0.35rem;
        color: var(--muted);
        font-size: 0.8rem;
        line-height: 1.4;
        font-weight: 500;
    }

    .positive { color: #0d9f6e; font-weight: 800; }
    .negative { color: #e04b5a; font-weight: 800; }

    .goal-progress,
    .progress-track {
        overflow: hidden;
        height: 0.48rem;
        margin-top: 0.65rem;
        border-radius: 999px;
        background: rgba(28,31,38,0.07);
    }
    .goal-progress span,
    .progress-track span {
        display: block;
        height: 100%;
        border-radius: inherit;
        background: linear-gradient(90deg, var(--lime) 0%, var(--mint-mid) 100%);
    }

    .chart-intro,
    .history-summary {
        margin: 0.85rem 0;
        padding: 1rem 1.15rem;
        border-radius: var(--radius-sm);
        background: var(--card);
        border: 1px solid rgba(28,31,38,0.04);
        box-shadow: var(--shadow-soft);
        color: var(--ink);
        font-size: 0.93rem;
    }

    .answer-good {
        background: var(--mint-soft);
        border: 1px solid rgba(126,217,176,0.3);
    }
    .answer-care {
        background: #fff8e8;
        border: 1px solid rgba(230,180,60,0.25);
    }
    .answer-risk {
        background: #ffecee;
        border: 1px solid rgba(224,75,90,0.2);
    }

    /* ===== TABS ===== */
    div[data-testid="stTabs"] {
        margin-top: 0.4rem;
    }
    div[data-testid="stTabs"] [data-baseweb="tab-list"] {
        gap: 0.4rem;
        background: transparent !important;
        flex-wrap: wrap !important;
        overflow-x: auto !important;
        -webkit-overflow-scrolling: touch;
        scrollbar-width: none;
        padding-bottom: 0.35rem;
    }
    div[data-testid="stTabs"] [data-baseweb="tab-list"]::-webkit-scrollbar {
        display: none;
    }
    div[data-testid="stTabs"] button {
        border-radius: 999px !important;
        color: #6b7280 !important;
        font-weight: 700 !important;
        background: #ffffff !important;
        border: 1px solid rgba(28,31,38,0.06) !important;
        box-shadow: var(--shadow-soft) !important;
        padding: 0.42rem 1.05rem !important;
        transition: 0.15s ease;
        white-space: nowrap !important;
        flex-shrink: 0 !important;
        min-height: 2.5rem;
        font-size: 0.88rem !important;
    }
    div[data-testid="stTabs"] button:hover {
        background: #f8f9fb !important;
        color: var(--ink) !important;
    }
    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: var(--ink) !important;
        background: var(--lime) !important;
        border-color: transparent !important;
        font-weight: 800 !important;
        box-shadow: 0 4px 16px rgba(217,255,0,0.3) !important;
    }
    div[data-testid="stTabs"] [data-baseweb="tab-highlight"],
    div[data-testid="stTabs"] [data-baseweb="tab-border"] {
        background: transparent !important;
        display: none !important;
    }

    /* Charts / forms / expanders */
    div[data-testid="stPlotlyChart"],
    div[data-testid="stDataFrame"],
    div[data-testid="stExpander"],
    div[data-testid="stForm"],
    [data-testid="stAlert"] {
        border-radius: var(--radius) !important;
        background: var(--card) !important;
        border: 1px solid rgba(28,31,38,0.04) !important;
        box-shadow: var(--shadow-soft) !important;
    }

    /* Plotly container overflow on mobile */
    div[data-testid="stPlotlyChart"] {
        overflow-x: auto !important;
        max-width: 100% !important;
    }
    .js-plotly-plot,
    .plotly {
        max-width: 100% !important;
    }

    /* Buttons - touch friendly */
    .stButton > button,
    [data-testid="stFormSubmitButton"] button {
        min-height: 2.75rem;
        border: 0 !important;
        border-radius: 999px !important;
        color: white !important;
        background: #1c1f26 !important;
        font-weight: 750 !important;
        box-shadow: var(--shadow-soft) !important;
        transition: 0.18s ease;
        width: 100%;
        padding: 0.55rem 1.1rem !important;
    }
    .stButton > button:hover,
    [data-testid="stFormSubmitButton"] button:hover {
        background: #2d323c !important;
        color: white !important;
        transform: translateY(-1px);
        box-shadow: 0 8px 20px rgba(28,31,38,0.12) !important;
    }

    div[data-testid="stDownloadButton"] button {
        min-height: 2.75rem;
        border-radius: 999px !important;
        color: var(--ink) !important;
        background: #ffffff !important;
        border: 1px solid rgba(28,31,38,0.08) !important;
        font-weight: 750 !important;
        box-shadow: var(--shadow-soft) !important;
        width: 100%;
    }
    div[data-testid="stDownloadButton"] button:hover {
        border-color: rgba(217,255,0,0.6) !important;
        background: #fefff5 !important;
    }

    /* Inputs */
    div[data-baseweb="input"],
    div[data-baseweb="select"] > div,
    textarea {
        border-radius: 12px !important;
        border: 1px solid rgba(28,31,38,0.08) !important;
        background: #ffffff !important;
        box-shadow: none !important;
        min-height: 2.6rem !important;
    }
    div[data-baseweb="input"]:focus-within,
    div[data-baseweb="select"] > div:focus-within,
    textarea:focus {
        border-color: rgba(126,217,176,0.7) !important;
        box-shadow: 0 0 0 3px rgba(184,240,216,0.35) !important;
    }

    /* Radio horizontal - wrap on mobile */
    div[role="radiogroup"] {
        flex-wrap: wrap !important;
        gap: 0.4rem !important;
    }

    /* Header */
    .page-header {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 1rem;
        margin-bottom: 1.15rem;
        flex-wrap: wrap;
    }
    .page-header h1 {
        margin: 0.1rem 0 0.2rem !important;
        font-size: clamp(1.55rem, 5vw, 2.25rem) !important;
        font-weight: 800 !important;
        letter-spacing: -0.03em;
    }
    .page-header p {
        margin: 0;
        color: var(--muted);
        font-size: 0.92rem;
        font-weight: 500;
    }
    .user-chip {
        display: flex;
        align-items: center;
        gap: 0.55rem;
        padding: 0.32rem 0.85rem 0.32rem 0.38rem;
        border-radius: 999px;
        background: #ffffff;
        box-shadow: var(--shadow-soft);
        border: 1px solid rgba(28,31,38,0.05);
        font-size: 0.84rem;
        font-weight: 700;
        color: var(--ink);
        white-space: nowrap;
        flex-shrink: 0;
    }
    .avatar-dot {
        width: 1.85rem;
        height: 1.85rem;
        display: grid;
        place-items: center;
        border-radius: 50%;
        background: linear-gradient(135deg, #1c1f26, #3a404c);
        color: white;
        font-size: 0.72rem;
        font-weight: 800;
        flex-shrink: 0;
    }
    .badge-new {
        display: inline-flex;
        align-items: center;
        padding: 0.22rem 0.65rem;
        border-radius: 999px;
        background: var(--mint);
        color: #0d5c3d;
        font-size: 0.74rem;
        font-weight: 800;
    }
    .badge-lime {
        display: inline-flex;
        align-items: center;
        padding: 0.22rem 0.65rem;
        border-radius: 999px;
        background: var(--lime);
        color: var(--ink);
        font-size: 0.74rem;
        font-weight: 800;
    }

    .card-icon {
        position: absolute;
        top: 0.95rem;
        right: 0.95rem;
        width: 2rem;
        height: 2rem;
        display: grid;
        place-items: center;
        border-radius: 50%;
        border: 1.5px solid rgba(28,31,38,0.1);
        font-size: 0.85rem;
        color: var(--ink);
        background: rgba(255,255,255,0.5);
    }
    .metric-card.accent .card-icon {
        border-color: rgba(28,31,38,0.18);
        background: rgba(255,255,255,0.35);
    }
    .metric-card.dark .card-icon {
        border-color: rgba(255,255,255,0.2);
        color: white;
        background: rgba(255,255,255,0.08);
    }

    .status-pill {
        display: inline-flex;
        align-items: center;
        padding: 0.2rem 0.6rem;
        border-radius: 999px;
        font-size: 0.72rem;
        font-weight: 750;
    }
    .status-ok { background: var(--mint-soft); color: #0d5c3d; }
    .status-warn { background: #fff4d6; color: #8a5a00; }
    .status-bad { background: #ffe4e8; color: #a31d2e; }

    /* History / list items flex layout */
    .history-item,
    .investment-item,
    .debt-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 0.85rem;
        min-height: auto;
        margin-bottom: 0.65rem;
    }
    .history-item > div:first-child,
    .investment-item > div:first-child,
    .debt-item > div:first-child {
        min-width: 0;
        flex: 1;
    }

    /* Dataframe horizontal scroll */
    div[data-testid="stDataFrame"] {
        overflow-x: auto !important;
    }

    /* Columns stack naturally via Streamlit, but tighten gap */
    [data-testid="column"] {
        min-width: 0 !important;
    }

    /* ========== TABLET (<= 980px) ========== */
    @media (max-width: 980px) {
        .main .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
            padding-top: 1rem;
        }

        .metric-grid,
        .indicator-grid,
        .answer-grid,
        .goal-grid,
        .debt-grid,
        .investment-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.75rem;
        }

        .metric-card,
        .indicator-card,
        .answer-card,
        .goal-card {
            min-height: 7rem;
            padding: 1rem;
        }

        .metric-value,
        .indicator-value,
        .answer-value {
            font-size: clamp(1.2rem, 4vw, 1.65rem);
        }

        .page-header {
            gap: 0.75rem;
        }
    }

    /* ========== MOBILE (<= 640px) ========== */
    @media (max-width: 640px) {
        .main .block-container {
            padding-left: 0.7rem !important;
            padding-right: 0.7rem !important;
            padding-top: 0.75rem !important;
            padding-bottom: 1.5rem !important;
        }

        .metric-grid,
        .indicator-grid,
        .answer-grid,
        .goal-grid,
        .debt-grid,
        .investment-grid {
            grid-template-columns: 1fr !important;
            gap: 0.65rem;
            margin: 0.75rem 0 1rem;
        }

        .metric-card,
        .indicator-card,
        .answer-card,
        .goal-card,
        .history-item,
        .investment-item,
        .debt-item {
            min-height: auto;
            padding: 0.95rem 1rem;
            border-radius: 16px;
        }

        .metric-value,
        .indicator-value,
        .answer-value {
            font-size: 1.45rem;
            margin-top: 0.5rem;
        }

        .metric-label,
        .indicator-top,
        .answer-question {
            font-size: 0.72rem;
        }

        .metric-foot,
        .indicator-note,
        .answer-action,
        .goal-meta,
        .debt-meta,
        .investment-meta {
            font-size: 0.76rem;
        }

        .card-icon {
            width: 1.75rem;
            height: 1.75rem;
            top: 0.8rem;
            right: 0.8rem;
            font-size: 0.78rem;
        }

        .page-header {
            flex-direction: column;
            align-items: stretch;
            gap: 0.65rem;
            margin-bottom: 0.9rem;
        }

        .page-header h1 {
            font-size: 1.55rem !important;
        }

        .page-header p {
            font-size: 0.85rem;
        }

        .user-chip {
            align-self: flex-start;
            font-size: 0.8rem;
            padding: 0.28rem 0.7rem 0.28rem 0.32rem;
        }

        .badge-new,
        .badge-lime {
            font-size: 0.7rem;
            padding: 0.18rem 0.55rem;
        }

        /* Tabs: horizontal scroll, smaller pills */
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {
            gap: 0.3rem !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            padding-bottom: 0.4rem;
            margin: 0 -0.2rem;
            padding-left: 0.2rem;
            padding-right: 0.2rem;
        }
        div[data-testid="stTabs"] button {
            padding: 0.38rem 0.8rem !important;
            font-size: 0.78rem !important;
            min-height: 2.35rem !important;
        }

        .chart-intro,
        .history-summary {
            padding: 0.85rem 0.95rem;
            font-size: 0.88rem;
            border-radius: 14px;
        }

        /* Full-width touch buttons */
        .stButton > button,
        [data-testid="stFormSubmitButton"] button,
        div[data-testid="stDownloadButton"] button {
            min-height: 3rem !important;
            font-size: 0.95rem !important;
        }

        /* Forms: stack fields tighter */
        div[data-testid="stForm"] {
            padding: 0.5rem 0 !important;
        }

        /* History items stack label/value better */
        .history-item,
        .investment-item,
        .debt-item {
            flex-direction: column;
            align-items: flex-start;
            gap: 0.45rem;
        }
        .history-item > div:last-child,
        .investment-item > div:last-child,
        .debt-item > div:last-child {
            align-self: flex-end;
            font-size: 1.05rem;
        }

        /* Expander headers more tappable */
        div[data-testid="stExpander"] summary,
        div[data-testid="stExpander"] details summary {
            min-height: 2.8rem;
            padding: 0.6rem 0.8rem !important;
        }

        /* Reduce hover lift on touch devices */
        .metric-card:hover,
        .indicator-card:hover {
            transform: none;
        }

        h2, h3, .stSubheader {
            font-size: 1.15rem !important;
        }

        /* Caption / footers */
        .stCaption, [data-testid="stCaptionContainer"] {
            font-size: 0.78rem !important;
        }
    }

    /* ========== SMALL PHONES (<= 380px) ========== */
    @media (max-width: 380px) {
        .main .block-container {
            padding-left: 0.55rem !important;
            padding-right: 0.55rem !important;
        }

        .metric-value,
        .indicator-value,
        .answer-value {
            font-size: 1.3rem;
        }

        div[data-testid="stTabs"] button {
            padding: 0.32rem 0.65rem !important;
            font-size: 0.72rem !important;
        }

        .page-header h1 {
            font-size: 1.4rem !important;
        }
    }

    /* Prefer reduced motion */
    @media (prefers-reduced-motion: reduce) {
        .metric-card,
        .indicator-card,
        .stButton > button {
            transition: none !important;
        }
        .metric-card:hover,
        .indicator-card:hover {
            transform: none !important;
        }
    }

    /* ===== FIXES MOBILE: labels, forms, wrap, overflow ===== */

    /* Labels always visible */
    label,
    [data-testid="stWidgetLabel"],
    [data-testid="stWidgetLabel"] p,
    [data-testid="stMarkdownContainer"] label,
    .stSelectbox label,
    .stTextInput label,
    .stNumberInput label,
    .stDateInput label,
    .stTextArea label,
    .stRadio label,
    .stMultiSelect label {
        color: #1c1f26 !important;
        opacity: 1 !important;
        visibility: visible !important;
        font-weight: 650 !important;
        font-size: 0.88rem !important;
        margin-bottom: 0.25rem !important;
    }

    /* Form fields: light bg + contrast */
    div[data-baseweb="input"],
    div[data-baseweb="input"] > div,
    div[data-baseweb="select"] > div,
    div[data-baseweb="base-input"],
    input,
    textarea,
    [data-baseweb="textarea"],
    .stTextInput input,
    .stNumberInput input,
    .stDateInput input,
    .stTextArea textarea {
        background: #ffffff !important;
        color: #1c1f26 !important;
        border: 1.5px solid rgba(28,31,38,0.12) !important;
        border-radius: 12px !important;
        min-height: 2.7rem !important;
        font-size: 0.95rem !important;
        caret-color: #1c1f26 !important;
    }
    div[data-baseweb="input"] input,
    div[data-baseweb="select"] input {
        color: #1c1f26 !important;
        background: transparent !important;
        -webkit-text-fill-color: #1c1f26 !important;
    }
    /* placeholder contrast */
    input::placeholder,
    textarea::placeholder {
        color: #8a90a0 !important;
        opacity: 1 !important;
        -webkit-text-fill-color: #8a90a0 !important;
    }
    /* select value text */
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div {
        color: #1c1f26 !important;
    }

    /* Cards: never clip long values */
    .metric-card,
    .indicator-card,
    .answer-card,
    .goal-card,
    .history-item,
    .investment-item,
    .debt-item {
        overflow: visible !important;
    }
    .metric-value,
    .indicator-value,
    .answer-value {
        overflow: visible !important;
        text-overflow: unset !important;
        white-space: normal !important;
        word-break: break-word !important;
        overflow-wrap: anywhere !important;
        line-height: 1.2 !important;
        max-width: 100% !important;
    }
    .metric-label,
    .metric-foot {
        white-space: normal !important;
        overflow: visible !important;
        word-break: break-word !important;
    }
    .card-icon {
        flex-shrink: 0;
    }

    /* Tabs: allow wrap on small screens */
    div[data-testid="stTabs"] [data-baseweb="tab-list"] {
        flex-wrap: wrap !important;
        row-gap: 0.4rem !important;
        overflow-x: visible !important;
    }
    div[data-testid="stTabs"] button {
        white-space: normal !important;
        line-height: 1.25 !important;
        height: auto !important;
        max-width: 100%;
    }

    /* Buttons wrap text instead of overflowing */
    .stButton > button,
    [data-testid="stFormSubmitButton"] button,
    div[data-testid="stDownloadButton"] button {
        white-space: normal !important;
        line-height: 1.25 !important;
        height: auto !important;
        min-height: 2.75rem !important;
        word-break: break-word !important;
    }

    /* Horizontal radio wraps */
    div[role="radiogroup"] {
        flex-wrap: wrap !important;
        gap: 0.45rem !important;
    }
    div[role="radiogroup"] label {
        white-space: normal !important;
    }

    /* Columns don't crush content */
    [data-testid="column"] {
        overflow: visible !important;
    }
    [data-testid="column"] > div {
        overflow: visible !important;
    }

    /* Streamlit metric fallback */
    [data-testid="stMetricValue"] {
        white-space: normal !important;
        overflow: visible !important;
        word-break: break-word !important;
    }
    [data-testid="stMetricLabel"] {
        color: #1c1f26 !important;
        opacity: 1 !important;
    }

    /* Subheaders / captions visible */
    h1, h2, h3, h4,
    .stSubheader,
    [data-testid="stCaptionContainer"],
    .stCaption {
        color: #1c1f26 !important;
        opacity: 1 !important;
    }

    /* Expander title visible */
    div[data-testid="stExpander"] summary,
    div[data-testid="stExpander"] summary p,
    div[data-testid="stExpander"] summary span {
        color: #1c1f26 !important;
        opacity: 1 !important;
        font-weight: 700 !important;
    }

    @media (max-width: 640px) {
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {
            flex-wrap: wrap !important;
            overflow-x: visible !important;
            justify-content: flex-start !important;
        }
        div[data-testid="stTabs"] button {
            flex: 1 1 auto !important;
            min-width: calc(50% - 0.3rem) !important;
            max-width: 100% !important;
            text-align: center !important;
            padding: 0.45rem 0.55rem !important;
            font-size: 0.76rem !important;
        }

        .metric-value,
        .indicator-value,
        .answer-value {
            font-size: 1.35rem !important;
            letter-spacing: -0.02em;
        }

        /* form labels a bit larger on phone */
        label, [data-testid="stWidgetLabel"] p {
            font-size: 0.9rem !important;
        }

        div[data-baseweb="input"],
        div[data-baseweb="select"] > div,
        input, textarea {
            min-height: 2.85rem !important;
            font-size: 16px !important; /* evita zoom iOS */
        }
    }

</style>
""",
    unsafe_allow_html=True,
)

# ====================== BANCO DE DADOS ======================
def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS transacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT NOT NULL,
            descricao TEXT,
            categoria TEXT,
            valor REAL NOT NULL,
            tipo TEXT,
            cartao TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS investimentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT,
            tipo TEXT,
            valor REAL,
            rentabilidade TEXT,
            descricao TEXT,
            status TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS dividas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT,
            credor TEXT,
            tipo TEXT,
            saldo_original REAL,
            desconto REAL,
            saldo_negociado REAL,
            parcela_possivel REAL,
            vencimento TEXT,
            prioridade TEXT,
            consequencia TEXT,
            status TEXT,
            proxima_acao TEXT,
            anotacoes TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS metas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT,
            nome TEXT,
            valor_meta REAL,
            valor_atual REAL,
            aporte_mensal REAL,
            prazo TEXT,
            status TEXT,
            anotacoes TEXT
        )
    """)
    conn.commit()
    conn.close()


def carregar_dados() -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query("SELECT * FROM transacoes ORDER BY data DESC, id DESC", conn)
    except Exception:
        df = pd.DataFrame(columns=["id", "data", "descricao", "categoria", "valor", "tipo", "cartao"])
    conn.close()
    if len(df) and "tipo" in df.columns:
        df["tipo"] = df["tipo"].replace({"Saida": "Saída", "saida": "Saída", "Entrada": "Entrada", "entrada": "Entrada"})
    return df


def carregar_investimentos() -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query("SELECT * FROM investimentos ORDER BY data DESC, id DESC", conn)
    except Exception:
        df = pd.DataFrame(columns=["id", "data", "tipo", "valor", "rentabilidade", "descricao", "status"])
    conn.close()
    return df


def carregar_dividas() -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query("SELECT * FROM dividas ORDER BY data DESC, id DESC", conn)
    except Exception:
        df = pd.DataFrame(
            columns=[
                "id", "data", "credor", "tipo", "saldo_original", "desconto",
                "saldo_negociado", "parcela_possivel", "vencimento", "prioridade",
                "consequencia", "status", "proxima_acao", "anotacoes",
            ]
        )
    conn.close()
    return df


def carregar_metas() -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query("SELECT * FROM metas ORDER BY data DESC, id DESC", conn)
    except Exception:
        df = pd.DataFrame(
            columns=[
                "id", "data", "nome", "valor_meta", "valor_atual",
                "aporte_mensal", "prazo", "status", "anotacoes",
            ]
        )
    conn.close()
    return df


def salvar_transacao(data, descricao, categoria, valor, tipo, cartao):
    conn = get_conn()
    conn.execute(
        "INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao) VALUES (?, ?, ?, ?, ?, ?)",
        (str(data), descricao, categoria, float(valor), tipo, cartao),
    )
    conn.commit()
    conn.close()


def excluir_transacao(id_):
    conn = get_conn()
    conn.execute("DELETE FROM transacoes WHERE id = ?", (int(id_),))
    conn.commit()
    conn.close()


def limpar_historico():
    conn = get_conn()
    conn.execute("DELETE FROM transacoes")
    conn.commit()
    conn.close()


def salvar_investimento(data, tipo, valor, rentabilidade, descricao, status):
    conn = get_conn()
    conn.execute(
        "INSERT INTO investimentos (data, tipo, valor, rentabilidade, descricao, status) VALUES (?, ?, ?, ?, ?, ?)",
        (str(data), tipo, float(valor), rentabilidade, descricao, status),
    )
    conn.commit()
    conn.close()


def excluir_investimento(id_):
    conn = get_conn()
    conn.execute("DELETE FROM investimentos WHERE id = ?", (int(id_),))
    conn.commit()
    conn.close()


def salvar_divida(
    data, credor, tipo, saldo_original, desconto, saldo_negociado,
    parcela_possivel, vencimento, prioridade, consequencia, status,
    proxima_acao, anotacoes,
):
    conn = get_conn()
    conn.execute(
        """
        INSERT INTO dividas (
            data, credor, tipo, saldo_original, desconto, saldo_negociado,
            parcela_possivel, vencimento, prioridade, consequencia, status,
            proxima_acao, anotacoes
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            str(data), credor, tipo, float(saldo_original), float(desconto),
            float(saldo_negociado), float(parcela_possivel), str(vencimento),
            prioridade, consequencia, status, proxima_acao, anotacoes,
        ),
    )
    conn.commit()
    conn.close()


def excluir_divida(id_):
    conn = get_conn()
    conn.execute("DELETE FROM dividas WHERE id = ?", (int(id_),))
    conn.commit()
    conn.close()


def salvar_meta(data, nome, valor_meta, valor_atual, aporte_mensal, prazo, status, anotacoes):
    conn = get_conn()
    conn.execute(
        """
        INSERT INTO metas (data, nome, valor_meta, valor_atual, aporte_mensal, prazo, status, anotacoes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (str(data), nome, float(valor_meta), float(valor_atual), float(aporte_mensal), prazo, status, anotacoes),
    )
    conn.commit()
    conn.close()


def excluir_meta(id_):
    conn = get_conn()
    conn.execute("DELETE FROM metas WHERE id = ?", (int(id_),))
    conn.commit()
    conn.close()


# ====================== FUNÇÕES DE DADOS ======================
def converter_valor(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    if not texto:
        return None

    negativo = (texto.startswith("(") and texto.endswith(")")) or texto.startswith("-")
    texto = (
        texto.replace("R$", "")
        .replace(" ", "")
        .replace("\u00a0", "")
        .replace("(", "")
        .replace(")", "")
    )
    texto = "".join(c for c in texto if c.isdigit() or c in ",.-")

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        numero = float(texto)
    except ValueError:
        return None
    return -abs(numero) if negativo else numero


def calcular_score_financeiro(entradas, saidas, saldo, investimentos, total_dividas, parcelas_dividas, progresso_metas):
    if entradas <= 0 and saidas <= 0:
        return 50

    taxa_sobra = (saldo / entradas) * 100 if entradas > 0 else -35
    comprometimento = ((saidas + parcelas_dividas) / entradas) * 100 if entradas > 0 else 100
    meses_reserva = investimentos / (saidas / 3) if saidas > 0 else (3 if investimentos > 0 else 0)

    score = 52
    score += max(-24, min(taxa_sobra, 24))
    score -= max(0, min(comprometimento - 70, 22))
    score += min(meses_reserva, 3) * 4
    score += min(progresso_metas / 10, 10)
    if total_dividas > 0:
        score -= min((total_dividas / max(entradas, 1)) * 8, 16)
    if saldo >= 0:
        score += 5
    return int(max(0, min(round(score), 100)))


def preparar_fluxo_mensal(df_transacoes: pd.DataFrame) -> pd.DataFrame:
    if df_transacoes.empty:
        return pd.DataFrame(columns=["mes", "entradas", "saidas", "saldo"])

    df_fluxo = df_transacoes.copy()
    df_fluxo["data_convertida"] = pd.to_datetime(df_fluxo["data"], errors="coerce")
    df_fluxo = df_fluxo.dropna(subset=["data_convertida"])
    if df_fluxo.empty:
        return pd.DataFrame(columns=["mes", "entradas", "saidas", "saldo"])

    df_fluxo["mes"] = df_fluxo["data_convertida"].dt.to_period("M").dt.to_timestamp()
    df_fluxo["entradas"] = df_fluxo["valor"].clip(lower=0)
    df_fluxo["saidas"] = df_fluxo["valor"].clip(upper=0).abs()
    return (
        df_fluxo.groupby("mes", as_index=False)[["entradas", "saidas"]]
        .sum()
        .assign(saldo=lambda d: d["entradas"] - d["saidas"])
        .sort_values("mes")
    )


def resumo_mes_recente(df_transacoes: pd.DataFrame):
    df_fluxo = df_transacoes.copy()
    df_fluxo["data_convertida"] = pd.to_datetime(df_fluxo["data"], errors="coerce")
    df_fluxo = df_fluxo.dropna(subset=["data_convertida"])
    if df_fluxo.empty:
        return pd.DataFrame(), "Sem mês"

    mes_recente = df_fluxo["data_convertida"].max().to_period("M")
    df_mes = df_fluxo[df_fluxo["data_convertida"].dt.to_period("M") == mes_recente]
    return df_mes, mes_recente.strftime("%m/%Y")


def style_plot(fig, height=320, show_legend=True):
    """Estilo compacto, legivel no mobile, sem toolbar."""
    fig.update_layout(
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(color="#1c1f26", size=11, family="Inter, sans-serif"),
        title=dict(
            font=dict(size=14, color="#1c1f26", family="Inter, sans-serif"),
            x=0.0,
            xanchor="left",
            pad=dict(t=0, b=8),
        ),
        margin=dict(l=28, r=16, t=44, b=64),
        height=height,
        autosize=True,
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.18,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(255,255,255,0)",
            font=dict(size=10, color="#1c1f26"),
            itemwidth=30,
            tracegroupgap=4,
        ) if show_legend else dict(visible=False),
        hoverlabel=dict(bgcolor="#1c1f26", font_color="#ffffff", font_size=12),
        separators=",.",
        uniformtext_minsize=9,
        uniformtext_mode="hide",
        showlegend=show_legend,
    )
    fig.update_xaxes(
        gridcolor="rgba(28,31,38,0.06)",
        tickfont=dict(color="#8a90a0", size=10),
        zeroline=False,
        automargin=True,
    )
    fig.update_yaxes(
        gridcolor="rgba(28,31,38,0.06)",
        tickfont=dict(color="#8a90a0", size=10),
        zeroline=False,
        automargin=True,
    )
    # remove modebar chrome
    fig.update_layout(modebar=dict(remove=["zoom", "pan", "select", "lasso", "zoomIn", "zoomOut", "autoScale", "resetScale"]))
    return fig


PLOTLY_CONFIG = {
    "displayModeBar": False,
    "displaylogo": False,
    "responsive": True,
    "staticPlot": False,
}


# ====================== IMPORTAÇÃO DE PLANILHA ======================
MESES_PLANILHA = {
    "janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
}

PALAVRAS_ENTRADA = {"entrada", "receita", "recebimento", "salario", "renda", "freelance", "reembolso", "venda"}
PALAVRAS_SAIDA = {"saida", "despesa", "debito", "credito", "cartao", "gasto", "fixo", "parcelado"}


def gerar_modelo_excel() -> bytes:
    arquivo = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Movimentações"

    cabecalhos = ["Data", "Descrição", "Categoria", "Valor", "Tipo", "Forma de Pagamento"]
    exemplos = [
        [date.today(), "Salário mensal", "Salário", 3500.00, "Entrada", "Conta corrente"],
        [date.today(), "Compras do mês", "Mercado", 420.00, "Saída", "Cartão"],
        [date.today(), "Internet", "Contas", 99.90, "Saída", "Pix"],
    ]
    ws.append(cabecalhos)
    for ex in exemplos:
        ws.append(ex)

    tabela = Table(displayName="TabelaMovimentacoes", ref=f"A1:F{len(exemplos)+1}")
    tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tabela)
    ws.freeze_panes = "A2"

    for col, largura in {"A": 14, "B": 28, "C": 18, "D": 16, "E": 12, "F": 22}.items():
        ws.column_dimensions[col].width = largura

    for cel in ws["A"][1:]:
        cel.number_format = "dd/mm/yyyy"
    for cel in ws["D"][1:]:
        cel.number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'

    for cel in ws[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor="111318")
        cel.alignment = Alignment(horizontal="center")

    wb.save(arquivo)
    arquivo.seek(0)
    return arquivo.getvalue()


def ler_planilha(arquivo):
    nome = getattr(arquivo, "name", "").lower()
    if nome.endswith(".csv"):
        try:
            return pd.read_csv(arquivo, sep=None, engine="python")
        except UnicodeDecodeError:
            arquivo.seek(0)
            return pd.read_csv(arquivo, sep=None, engine="python", encoding="latin-1")

    dados = arquivo.read()
    if nome.endswith(".xlsx"):
        return pd.read_excel(BytesIO(dados), sheet_name=None, header=None, engine="openpyxl")
    if nome.endswith(".xls"):
        return pd.read_excel(BytesIO(dados), sheet_name=None, header=None)
    raise ValueError("Envie uma planilha CSV, XLSX ou XLS.")


def preparar_importacao(dados):
    """Suporta planilha simples (colunas Data/Descrição/Valor...) ou abas mensais."""
    if isinstance(dados, dict):
        # Tenta modelo de organização financeira (abas Janeiro, Fevereiro...)
        linhas = []
        for nome_aba, df_aba in dados.items():
            mes = MESES_PLANILHA.get(normalizar(nome_aba))
            if not mes:
                continue
            # Procura blocos com "valor" e "descricao/nome"
            for i in range(len(df_aba.index)):
                valores_linha = [normalizar(df_aba.iat[i, c]) for c in range(len(df_aba.columns))]
                if "valor" not in valores_linha:
                    continue
                col_desc = next((c for c, v in enumerate(valores_linha) if v in ("descricao", "nome")), None)
                col_valor = next((c for c, v in enumerate(valores_linha) if v == "valor" and (col_desc is None or c > col_desc)), None)
                if col_desc is None or col_valor is None:
                    continue

                for j in range(i + 1, len(df_aba.index)):
                    descricao = limpar_texto(df_aba.iat[j, col_desc])
                    valor_raw = converter_valor(df_aba.iat[j, col_valor])
                    desc_n = normalizar(descricao)
                    if not descricao and valor_raw is None:
                        break
                    if desc_n.startswith("total") or desc_n in ("saidas", "entradas", "investimentos", "reserva"):
                        break
                    if valor_raw is None or valor_raw == 0:
                        continue

                    tipo = "Entrada" if valor_raw > 0 or any(p in desc_n for p in PALAVRAS_ENTRADA) else "Saída"
                    if any(p in desc_n for p in PALAVRAS_SAIDA):
                        tipo = "Saída"
                    if any(p in desc_n for p in PALAVRAS_ENTRADA):
                        tipo = "Entrada"

                    linhas.append({
                        "data": date(date.today().year, mes, 1).isoformat(),
                        "descricao": descricao or "Importado",
                        "categoria": "Receita" if tipo == "Entrada" else "Outros",
                        "valor": abs(valor_raw) if tipo == "Entrada" else -abs(valor_raw),
                        "tipo": tipo,
                        "cartao": "Planilha",
                    })

        if linhas:
            return pd.DataFrame(linhas).drop_duplicates(), 0

        # Fallback: primeira aba com header
        primeira = next(iter(dados.values()), pd.DataFrame())
        if primeira.empty:
            return pd.DataFrame(), 0
        primeira = primeira.copy()
        primeira.columns = primeira.iloc[0]
        dados = primeira.iloc[1:].reset_index(drop=True)

    # Planilha tabular normal
    aliases = {
        "data": ["data", "dt", "dia", "date"],
        "descricao": ["descricao", "descrição", "historico", "histórico", "nome", "lancamento"],
        "categoria": ["categoria", "grupo", "classificacao"],
        "valor": ["valor", "valor r$", "valor rs", "amount", "preco"],
        "tipo": ["tipo", "natureza", "entrada saida"],
        "cartao": ["forma de pagamento", "pagamento", "cartao", "cartão", "conta", "banco"],
    }
    colunas = {normalizar(c): c for c in dados.columns}
    mapa = {}
    for destino, opcoes in aliases.items():
        for opcao in opcoes:
            if normalizar(opcao) in colunas:
                mapa[destino] = colunas[normalizar(opcao)]
                break

    if "valor" not in mapa:
        raise ValueError("A planilha precisa ter uma coluna de valor.")

    linhas = []
    ignoradas = 0
    for _, row in dados.iterrows():
        valor_original = converter_valor(row.get(mapa["valor"]))
        if valor_original is None or valor_original == 0:
            ignoradas += 1
            continue

        data_conv = pd.to_datetime(row.get(mapa.get("data"), date.today()), errors="coerce", dayfirst=True)
        if pd.isna(data_conv):
            data_conv = pd.Timestamp(date.today())

        tipo_texto = normalizar(row.get(mapa.get("tipo"), ""))
        if any(p in tipo_texto for p in PALAVRAS_ENTRADA):
            tipo = "Entrada"
        elif any(p in tipo_texto for p in PALAVRAS_SAIDA):
            tipo = "Saída"
        else:
            tipo = "Entrada" if valor_original > 0 else "Saída"

        linhas.append({
            "data": data_conv.date().isoformat(),
            "descricao": limpar_texto(row.get(mapa.get("descricao")), "Importado da planilha"),
            "categoria": limpar_texto(row.get(mapa.get("categoria")), "Importado"),
            "valor": abs(valor_original) if tipo == "Entrada" else -abs(valor_original),
            "tipo": tipo,
            "cartao": limpar_texto(row.get(mapa.get("cartao")), "Planilha"),
        })

    return pd.DataFrame(linhas), ignoradas


def chave_movimentacao(registro) -> tuple:
    data_mov = pd.to_datetime(registro.get("data"), errors="coerce", dayfirst=True)
    data_norm = data_mov.date().isoformat() if not pd.isna(data_mov) else limpar_texto(registro.get("data"))
    descricao = normalizar(limpar_texto(registro.get("descricao")))
    tipo = normalizar(limpar_texto(registro.get("tipo")))
    valor = round(float(converter_valor(registro.get("valor")) or 0), 2)
    return data_norm, descricao, valor, tipo


def conciliar_movimentacoes(df_importado: pd.DataFrame, df_existente: pd.DataFrame):
    if df_importado.empty:
        return df_importado.copy(), 0

    chaves_existentes = {chave_movimentacao(r) for r in df_existente.to_dict(orient="records")}
    chaves_novas = set()
    indices = []
    duplicadas = 0

    for idx, reg in df_importado.iterrows():
        chave = chave_movimentacao(reg)
        if chave in chaves_existentes or chave in chaves_novas:
            duplicadas += 1
            continue
        chaves_novas.add(chave)
        indices.append(idx)

    return df_importado.loc[indices].reset_index(drop=True), duplicadas


def importar_movimentacoes(df_importado: pd.DataFrame):
    df_novo, duplicadas = conciliar_movimentacoes(df_importado, carregar_dados())
    if df_novo.empty:
        return 0, duplicadas

    conn = get_conn()
    conn.executemany(
        "INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao) VALUES (?, ?, ?, ?, ?, ?)",
        df_novo[["data", "descricao", "categoria", "valor", "tipo", "cartao"]].itertuples(index=False, name=None),
    )
    conn.commit()
    conn.close()
    return len(df_novo), duplicadas


# ====================== RELATÓRIO PDF ======================
def gerar_pdf(df, investimentos, dividas, metas) -> bytes:
    linhas = []

    def add(texto="", largura=90):
        for parte in textwrap.wrap(str(texto), width=largura) or [""]:
            linhas.append(parte)

    entradas = df[df["valor"] > 0]["valor"].sum() if len(df) else 0
    saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) else 0
    saldo = df["valor"].sum() if len(df) else 0
    total_inv = investimentos["valor"].sum() if len(investimentos) else 0

    add("Dashboard Financeiro - Relatório Detalhado")
    add(f"Emitido em {date.today().strftime('%d/%m/%Y')}")
    add("")
    add(f"Entradas: {brl(entradas)}")
    add(f"Saídas: {brl(saidas)}")
    add(f"Saldo: {brl(saldo)}")
    add(f"Investimentos: {brl(total_inv)}")
    add(f"Movimentações: {len(df)} | Investimentos: {len(investimentos)} | Dívidas: {len(dividas)} | Metas: {len(metas)}")
    add("")
    add("--- Movimentações ---")
    for _, row in df.iterrows():
        add(f"{data_br(row['data'])} | {row['descricao']} | {row['categoria']} | {row['tipo']} | {brl(row['valor'])}")

    add("")
    add("--- Investimentos ---")
    for _, inv in investimentos.iterrows():
        add(f"{data_br(inv['data'])} | {inv['descricao']} | {inv['tipo']} | {brl(inv['valor'])} | {inv['status']}")

    add("")
    add("--- Dívidas ---")
    for _, d in dividas.iterrows():
        saldo_b = d["saldo_negociado"] if d["saldo_negociado"] > 0 else d["saldo_original"]
        add(f"{data_br(d['data'])} | {d['credor']} | {d['tipo']} | {brl(saldo_b)} | {d['status']}")

    add("")
    add("--- Metas ---")
    for _, m in metas.iterrows():
        add(f"{m['nome']} | Meta: {brl(m['valor_meta'])} | Atual: {brl(m['valor_atual'])} | {m['status']}")

    # Gera PDF simples
    paginas = [linhas[i:i + 42] for i in range(0, len(linhas), 42)] or [[]]
    objetos = {
        1: b"<< /Type /Catalog /Pages 2 0 R >>",
        3: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
        4: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>",
    }
    refs = []
    for idx, pagina in enumerate(paginas):
        page_id = 5 + idx * 2
        content_id = page_id + 1
        refs.append(f"{page_id} 0 R")
        comandos = []
        y = 795
        for linha in pagina:
            fonte = "F2" if y == 795 else "F1"
            tamanho = 14 if y == 795 else 9
            texto = (
                str(linha)
                .replace("\\", "\\\\")
                .replace("(", "\\(")
                .replace(")", "\\)")
                .encode("latin-1", errors="replace")
                .decode("latin-1")
            )
            comandos.append(f"BT /{fonte} {tamanho} Tf 50 {y} Td ({texto}) Tj ET")
            y -= 17
        fluxo = "\n".join(comandos).encode("latin-1")
        objetos[page_id] = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
            f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
            f"/Contents {content_id} 0 R >>"
        ).encode("latin-1")
        objetos[content_id] = f"<< /Length {len(fluxo)} >>\nstream\n".encode("latin-1") + fluxo + b"\nendstream"

    objetos[2] = f"<< /Type /Pages /Kids [{' '.join(refs)}] /Count {len(paginas)} >>".encode("latin-1")

    pdf = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for numero in range(1, max(objetos) + 1):
        offsets.append(len(pdf))
        pdf.extend(f"{numero} 0 obj\n".encode("latin-1"))
        pdf.extend(objetos[numero])
        pdf.extend(b"\nendobj\n")

    xref = len(pdf)
    pdf.extend(f"xref\n0 {len(offsets)}\n".encode("latin-1"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode("latin-1"))
    pdf.extend(f"trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode("latin-1"))
    return bytes(pdf)


# ====================== INICIALIZAÇÃO ======================
init_db()
df = carregar_dados()
investimentos = carregar_investimentos()
dividas = carregar_dividas()
metas = carregar_metas()

# Métricas principais
entradas = df[df["valor"] > 0]["valor"].sum() if len(df) else 0.0
saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) else 0.0
saldo = df["valor"].sum() if len(df) else 0.0
total_investido = investimentos["valor"].sum() if len(investimentos) else 0.0

if len(dividas):
    d_base = dividas.copy()
    d_base["saldo_base"] = d_base["saldo_negociado"].where(d_base["saldo_negociado"] > 0, d_base["saldo_original"])
    total_dividas_abertas = d_base[d_base["status"] != "Quitada"]["saldo_base"].sum()
    parcelas_dividas = d_base[d_base["status"] != "Quitada"]["parcela_possivel"].sum()
else:
    total_dividas_abertas = 0.0
    parcelas_dividas = 0.0

progresso_metas = 0.0
if len(metas):
    total_meta = metas["valor_meta"].sum()
    total_atual = metas["valor_atual"].sum()
    progresso_metas = (total_atual / total_meta * 100) if total_meta > 0 else 0.0

score = calcular_score_financeiro(
    entradas, saidas, saldo, total_investido, total_dividas_abertas, parcelas_dividas, progresso_metas
)
taxa_sobra = (saldo / entradas * 100) if entradas > 0 else 0.0
comprometimento = ((saidas + parcelas_dividas) / entradas * 100) if entradas > 0 else 0.0

# ====================== CABEÇALHO ======================
st.markdown(
    f"""
<div class="page-header">
    <div>
        <div style="display:flex; align-items:center; gap:0.6rem; margin-bottom:0.15rem;">
            <span style="color:#8b90a0; font-weight:700; font-size:0.88rem;">Financeiro</span>
            <span class="badge-new">Saúde {score}/100</span>
        </div>
        <h1>Dashboard</h1>
        <p>Veja seu dinheiro, metas e dívidas em um só lugar.</p>
    </div>
    <div class="user-chip">
        <span class="avatar-dot">DF</span>
        Visão pessoal
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# Cards principais (mix Influency lime + Lodgify mint)
st.markdown(
    f"""
<div class="metric-grid">
    <div class="metric-card accent">
        <div class="card-icon">↗</div>
        <div class="metric-label">Saldo atual</div>
        <div class="metric-value">{brl(saldo)}</div>
        <div class="metric-foot">Neste período</div>
    </div>
    <div class="metric-card mint">
        <div class="card-icon">↑</div>
        <div class="metric-label">Entradas</div>
        <div class="metric-value">{brl(entradas)}</div>
        <div class="metric-foot">Receitas registradas</div>
    </div>
    <div class="metric-card">
        <div class="card-icon">↓</div>
        <div class="metric-label">Saídas</div>
        <div class="metric-value">{brl(saidas)}</div>
        <div class="metric-foot">Despesas acumuladas</div>
    </div>
    <div class="metric-card">
        <div class="card-icon">◆</div>
        <div class="metric-label">Patrimônio</div>
        <div class="metric-value">{brl(total_investido)}</div>
        <div class="metric-foot">Investimentos</div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# ====================== ABAS ======================
aba = st.tabs(["➕ Nova Movimentação", "📊 Dashboard", "🎯 Metas", "🤝 Dívidas", "📈 Investimentos", "📋 Histórico"])

# ---------- ABA 1: Nova Movimentação ----------
with aba[0]:
    st.subheader("Adicionar Nova Movimentação")
    tipo_sel = st.radio("Tipo", ["Entrada", "Saída"], horizontal=True)

    with st.form("form_mov"):
        c1, c2 = st.columns(2)
        with c1:
            data_mov = st.date_input("Data", value=date.today(), format="DD/MM/YYYY")
            descricao = st.text_input("Descrição")
            if tipo_sel == "Entrada":
                categoria = st.selectbox("Categoria", ["Salário", "Rendas Extras", "Freelance", "Reembolso", "Outro"])
            else:
                categoria = st.selectbox(
                    "Categoria",
                    ["Mercado", "Aluguel", "Contas", "Lazer", "Roupa", "Beleza", "Transporte", "Dívidas", "Outro"],
                )
        with c2:
            valor = st.number_input("Valor R$", value=0.0, min_value=0.0, step=0.01)
            cartao = st.text_input("Forma de pagamento")

        if st.form_submit_button("Salvar movimentação"):
            if not descricao.strip():
                st.error("Informe uma descrição.")
            elif valor <= 0:
                st.error("Informe um valor maior que zero.")
            else:
                valor_final = valor if tipo_sel == "Entrada" else -abs(valor)
                try:
                    salvar_transacao(data_mov, descricao.strip(), categoria, valor_final, tipo_sel, cartao.strip())
                    st.success("Movimentação salva!")
                    st.rerun()
                except Exception as e:
                    st.error(mensagem_erro_usuario(e))

# ---------- ABA 2: Dashboard ----------
with aba[1]:
    st.subheader("Dashboard em Tempo Real")

    # Indicadores
    st.markdown(
        f"""
<div class="indicator-grid">
    <div class="indicator-card">
        <div class="indicator-top">Score financeiro</div>
        <div class="indicator-value">{score}/100</div>
        <div class="indicator-note">Combina saldo, dívidas, reserva e metas.</div>
        <div class="progress-track"><span style="width:{limitar_percentual(score)}%"></span></div>
    </div>
    <div class="indicator-card">
        <div class="indicator-top">Taxa de sobra</div>
        <div class="indicator-value">{pct(taxa_sobra)}</div>
        <div class="indicator-note">Parte das entradas que virou saldo.</div>
        <div class="progress-track"><span style="width:{limitar_percentual(taxa_sobra)}%"></span></div>
    </div>
    <div class="indicator-card">
        <div class="indicator-top">Comprometimento</div>
        <div class="indicator-value">{pct(comprometimento)}</div>
        <div class="indicator-note">Saídas + parcelas vs entradas.</div>
        <div class="progress-track"><span style="width:{limitar_percentual(comprometimento)}%"></span></div>
    </div>
    <div class="indicator-card">
        <div class="indicator-top">Dívidas abertas</div>
        <div class="indicator-value">{brl(total_dividas_abertas)}</div>
        <div class="indicator-note">Saldo a negociar ou pagar.</div>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )

    # Respostas simples
    df_mes, nome_mes = resumo_mes_recente(df)
    fluxo = preparar_fluxo_mensal(df)
    entradas_mes = df_mes[df_mes["valor"] > 0]["valor"].sum() if len(df_mes) else 0
    saidas_mes = abs(df_mes[df_mes["valor"] < 0]["valor"].sum()) if len(df_mes) else 0
    saldo_mes = entradas_mes - saidas_mes
    media_sobra = fluxo["saldo"].mean() if len(fluxo) else saldo

    if saldo_mes >= 0:
        status_mes, classe_mes = "No azul", "answer-good"
        acao_mes = f"Em {nome_mes} sobrou {brl(saldo_mes)}. Separe uma parte antes de gastar."
    else:
        status_mes, classe_mes = "Atenção", "answer-risk"
        acao_mes = f"Em {nome_mes} faltou {brl(abs(saldo_mes))}. Reduza o maior gasto primeiro."

    despesas_mes = df_mes[df_mes["valor"] < 0].copy() if len(df_mes) else pd.DataFrame()
    if len(despesas_mes):
        despesas_mes["abs"] = despesas_mes["valor"].abs()
        top = despesas_mes.groupby("categoria")["abs"].sum().sort_values(ascending=False)
        maior_gasto = f"{top.index[0]} · {brl(top.iloc[0])}" if len(top) else "Sem gastos"
    else:
        maior_gasto = "Sem gastos no mês"

    st.markdown(
        f"""
<div class="chart-intro"><strong>O que preciso saber agora?</strong><br>Respostas simples com base nos seus números.</div>
<div class="answer-grid">
    <div class="answer-card {classe_mes}">
        <div class="answer-question">Como está meu mês?</div>
        <div class="answer-value">{status_mes}</div>
        <div class="answer-action">{acao_mes}</div>
    </div>
    <div class="answer-card answer-care">
        <div class="answer-question">Maior gasto</div>
        <div class="answer-value">{escape(maior_gasto)}</div>
        <div class="answer-action">Revise essa categoria primeiro.</div>
    </div>
    <div class="answer-card {'answer-good' if media_sobra > 0 else 'answer-risk'}">
        <div class="answer-question">Quanto sobra em média?</div>
        <div class="answer-value">{brl(media_sobra)}</div>
        <div class="answer-action">Esse valor pode virar reserva ou meta.</div>
    </div>
    <div class="answer-card answer-care">
        <div class="answer-question">Próxima ação</div>
        <div class="answer-value">{"Guardar primeiro" if media_sobra > 0 else "Cortar vazamento"}</div>
        <div class="answer-action">{"Reserve uma parte ao receber." if media_sobra > 0 else "Escolha uma categoria para reduzir."}</div>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )

    # Projeção futura
    with st.expander("🔮 Meu dinheiro futuro", expanded=False):
        st.caption("Simule os próximos 6 meses. A projeção usa sua média mensal atual.")
        c1, c2, c3 = st.columns(3)
        with c1:
            reduzir = st.slider("E se eu gastar menos por mês?", 0, 3000, 0, 50, format="R$ %d")
        with c2:
            extra = st.slider("E se eu receber um extra por mês?", 0, 5000, 0, 50, format="R$ %d")
        with c3:
            meta_mes = st.slider("Quero separar para metas por mês", 0, 5000, 0, 50, format="R$ %d")

        media_entradas = fluxo["entradas"].mean() if len(fluxo) else entradas
        media_saidas = fluxo["saidas"].mean() if len(fluxo) else saidas
        fluxo_normal = media_sobra + reduzir + extra - meta_mes
        ajuste_bom = max(media_entradas * 0.05, 100) if media_entradas else 100
        ajuste_apertado = max(media_saidas * 0.08, 100) if media_saidas else 100

        inicio = pd.Timestamp(date.today()).to_period("M").to_timestamp()
        proj = []
        s_normal = s_bom = s_apertado = saldo
        for i in range(1, 7):
            mes = inicio + pd.DateOffset(months=i)
            s_normal += fluxo_normal
            s_bom += fluxo_normal + ajuste_bom
            s_apertado += fluxo_normal - ajuste_apertado
            proj.extend([
                {"Mês": mes, "Cenário": "Normal", "Saldo projetado": s_normal},
                {"Mês": mes, "Cenário": "Bom", "Saldo projetado": s_bom},
                {"Mês": mes, "Cenário": "Apertado", "Saldo projetado": s_apertado},
            ])
        df_proj = pd.DataFrame(proj)

        neg = df_proj[(df_proj["Cenário"] == "Normal") & (df_proj["Saldo projetado"] < 0)]
        if neg.empty:
            resp = "Pelo cenário normal, seu saldo continua positivo nos próximos 6 meses."
        else:
            resp = f"No cenário normal, seu saldo pode ficar negativo em {neg.iloc[0]['Mês'].strftime('%m/%Y')}."

        st.markdown(f"""<div class="history-summary"><strong>Vou ficar sem dinheiro?</strong><br>{resp}</div>""", unsafe_allow_html=True)

        fig = px.line(
            df_proj, x="Mês", y="Saldo projetado", color="Cenário",
            title="Como deve ficar meu dinheiro nos próximos meses?",
            markers=True,
            color_discrete_map={"Bom": "#0d9f6e", "Normal": "#d9ff00", "Apertado": "#e04b5a"},
            labels={"Saldo projetado": "Saldo projetado", "Mês": "Mês"},
        )
        fig.update_traces(
            line=dict(width=3),
            marker=dict(size=8),
            hovertemplate="<b>%{fullData.name}</b><br>%{x|%m/%Y}<br>R$ %{y:,.2f}<extra></extra>",
        )
        fig.update_yaxes(tickprefix="R$ ")
        fig.update_xaxes(tickformat="%m/%Y")
        fig = style_plot(fig, height=340)
        fig.update_layout(
            legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center", title_text=""),
            margin=dict(l=28, r=12, t=44, b=70),
        )
        st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

    # Importação
    with st.expander("📤 Subir planilha de movimentações", expanded=False):
        st.caption("Novos registros entram no mesmo histórico. Duplicados são ignorados.")
        st.download_button(
            "Baixar modelo Excel",
            data=gerar_modelo_excel(),
            file_name="modelo-dashboard-financeiro.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        arquivo = st.file_uploader("Escolha a planilha", type=["csv", "xlsx", "xls"])
        if arquivo is not None:
            try:
                dados = ler_planilha(arquivo)
                df_imp, ignoradas = preparar_importacao(dados)
                df_novo, duplicadas = conciliar_movimentacoes(df_imp, df)

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Novas entradas", brl(df_novo[df_novo["valor"] > 0]["valor"].sum() if len(df_novo) else 0))
                c2.metric("Novas saídas", brl(abs(df_novo[df_novo["valor"] < 0]["valor"].sum() if len(df_novo) else 0)))
                c3.metric("Impacto no saldo", brl(df_novo["valor"].sum() if len(df_novo) else 0))
                c4.metric("Novos registros", len(df_novo))

                if ignoradas:
                    st.caption(f"{ignoradas} linhas ignoradas (sem valor válido).")
                if duplicadas:
                    st.info(f"{duplicadas} lançamentos duplicados foram reconhecidos e ignorados.")

                if len(df_novo):
                    previa = df_novo.rename(columns={
                        "data": "Data", "descricao": "Descrição", "categoria": "Categoria",
                        "valor": "Valor", "tipo": "Tipo", "cartao": "Pagamento",
                    })
                    previa["Data"] = previa["Data"].map(data_br)
                    previa["Valor"] = previa["Valor"].map(brl)
                    st.dataframe(previa, use_container_width=True, hide_index=True)
                    if st.button("Integrar movimentações ao dashboard", type="primary"):
                        total, dup = importar_movimentacoes(df_imp)
                        st.success(f"{total} movimentações integradas. {dup} duplicadas ignoradas.")
                        st.rerun()
                else:
                    st.info("Todos os lançamentos já estão no dashboard ou não foram reconhecidos.")
            except Exception as e:
                st.error(f"Não consegui importar: {mensagem_erro_usuario(e)}")

    # Gráficos
    if len(df):
        df_chart = df.copy()
        df_chart["categoria"] = df_chart["categoria"].replace("", "Sem categoria").fillna("Sem categoria")
        df_chart["cartao"] = df_chart["cartao"].replace("", "Não informado").fillna("Não informado")
        df_chart["valor_abs"] = df_chart["valor"].abs()
        df_chart["data_convertida"] = pd.to_datetime(df_chart["data"], errors="coerce")

        df_tl = df_chart.dropna(subset=["data_convertida"]).sort_values("data_convertida")
        if len(df_tl):
            df_tl["saldo_acumulado"] = df_tl["valor"].cumsum()
            fig_saldo = px.area(
                df_tl, x="data_convertida", y="saldo_acumulado",
                title="Saldo acumulado ao longo do tempo",
                labels={"data_convertida": "Data", "saldo_acumulado": "Saldo"},
            )
            fig_saldo.update_traces(
                line=dict(color="#1c1f26", width=3),
                fillcolor="rgba(126,217,176,0.28)",
                hovertemplate="<b>%{x|%d/%m/%Y}</b><br>Saldo: R$ %{y:,.2f}<extra></extra>",
            )
            fig_saldo.update_yaxes(tickprefix="R$ ")
            fig_saldo.update_xaxes(tickformat="%d/%m/%Y")
            st.plotly_chart(style_plot(fig_saldo, height=320, show_legend=False), use_container_width=True, config=PLOTLY_CONFIG)

        cat_total = (
            df_chart.groupby("categoria", as_index=False)["valor_abs"]
            .sum()
            .sort_values("valor_abs", ascending=False)
        )
        fluxo_tipo = df_chart.groupby(["categoria", "tipo"], as_index=False)["valor_abs"].sum()

        # --- Donut + Barras (lado a lado) ---
        c1, c2 = st.columns(2)
        with c1:
            # Top 5 + "Outras" (evita labels cortados no mobile)
            if len(cat_total) > 5:
                top8 = cat_total.head(5).copy()
                outros_valor = float(cat_total.iloc[5:]["valor_abs"].sum())
                if outros_valor > 0:
                    top8 = pd.concat([
                        top8,
                        pd.DataFrame([{"categoria": "Outras", "valor_abs": outros_valor}]),
                    ], ignore_index=True)
            else:
                top8 = cat_total.copy()

            fig = px.pie(
                top8,
                names="categoria",
                values="valor_abs",
                hole=0.58,
                title="Distribuição por categoria",
                color_discrete_sequence=[
                    "#d9ff00", "#1c1f26", "#7ed9b0", "#a8c4ff",
                    "#b8f0d8", "#c5c9d4", "#f3ff9a",
                ],
            )
            fig.update_traces(
                textposition="inside",
                textinfo="percent",
                textfont=dict(size=11, color="#1c1f26"),
                insidetextorientation="horizontal",
                marker=dict(line=dict(color="rgba(255,255,255,0.95)", width=2)),
                pull=[0.01] * len(top8),
                hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>",
                rotation=20,
            )
            fig = style_plot(fig, height=320)
            fig.update_layout(
                margin=dict(l=10, r=10, t=44, b=72),
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.12,
                    xanchor="center",
                    x=0.5,
                    font=dict(size=11),
                    itemwidth=40,
                ),
            )
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

        with c2:
            fig2 = px.bar(
                fluxo_tipo,
                x="categoria",
                y="valor_abs",
                color="tipo",
                title="Entradas x Saídas",
                color_discrete_map={"Entrada": "#d9ff00", "Saída": "#1c1f26"},
                labels={"categoria": "Categoria", "valor_abs": "Valor", "tipo": "Tipo"},
                barmode="group",
            )
            fig2.update_traces(
                marker_line_color="rgba(255,255,255,0.9)",
                marker_line_width=1,
                hovertemplate="<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>",
            )
            fig2.update_yaxes(tickprefix="R$ ")
            fig2.update_xaxes(tickangle=-35)
            fig2 = style_plot(fig2, height=340)
            fig2.update_layout(
                margin=dict(l=28, r=12, t=44, b=80),
                legend=dict(orientation="h", y=-0.22, x=0.5, xanchor="center"),
            )
            st.plotly_chart(fig2, use_container_width=True, config=PLOTLY_CONFIG)

        # --- Barras horizontais + Formas de pagamento ---
        c3, c4 = st.columns(2)
        with c3:
            top = cat_total.head(8).sort_values("valor_abs", ascending=True)
            fig3 = px.bar(
                top,
                x="valor_abs",
                y="categoria",
                orientation="h",
                title="Categorias que mais movimentam dinheiro",
                color="valor_abs",
                color_continuous_scale=["#e8faf3", "#7ed9b0", "#1c1f26"],
                labels={"valor_abs": "Valor", "categoria": ""},
                text="valor_abs",
            )
            fig3.update_traces(
                texttemplate="R$ %{x:,.0f}",
                textposition="outside",
                textfont=dict(size=11, color="#1c1f26"),
                hovertemplate="<b>%{y}</b><br>R$ %{x:,.2f}<extra></extra>",
                marker_line_color="rgba(255,255,255,0.9)",
                marker_line_width=1,
                cliponaxis=False,
            )
            fig3.update_layout(coloraxis_showscale=False)
            max_v = float(top["valor_abs"].max()) if len(top) else 1.0
            fig3.update_xaxes(tickprefix="R$ ", range=[0, max_v * 1.25 if max_v > 0 else 1])
            fig3 = style_plot(fig3, height=320, show_legend=False)
            fig3.update_layout(margin=dict(l=8, r=40, t=44, b=36))
            st.plotly_chart(fig3, use_container_width=True, config=PLOTLY_CONFIG)

        with c4:
            pag = (
                df_chart.groupby("cartao", as_index=False)["valor_abs"]
                .sum()
                .sort_values("valor_abs", ascending=False)
            )
            if len(pag) > 5:
                top_pag = pag.head(4).copy()
                rest_pag = float(pag.iloc[4:]["valor_abs"].sum())
                if rest_pag > 0:
                    top_pag = pd.concat([
                        top_pag,
                        pd.DataFrame([{"cartao": "Outras", "valor_abs": rest_pag}]),
                    ], ignore_index=True)
                pag = top_pag
            fig4 = px.pie(
                pag,
                names="cartao",
                values="valor_abs",
                hole=0.58,
                title="Formas de pagamento",
                color_discrete_sequence=["#1c1f26", "#d9ff00", "#7ed9b0", "#a8c4ff", "#b8f0d8", "#c5c9d4"],
            )
            fig4.update_traces(
                textposition="inside",
                textinfo="percent",
                textfont=dict(size=11, color="#1c1f26"),
                insidetextorientation="horizontal",
                marker=dict(line=dict(color="rgba(255,255,255,0.95)", width=2)),
                pull=[0.01] * len(pag),
                hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>",
            )
            fig4 = style_plot(fig4, height=320)
            fig4.update_layout(
                margin=dict(l=10, r=10, t=44, b=72),
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.12,
                    xanchor="center",
                    x=0.5,
                    font=dict(size=11),
                ),
            )
            st.plotly_chart(fig4, use_container_width=True, config=PLOTLY_CONFIG)

        if len(fluxo):
            fig_m = px.bar(
                fluxo,
                x="mes",
                y=["entradas", "saidas"],
                barmode="group",
                title="Evolução mensal consolidada",
                color_discrete_map={"entradas": "#d9ff00", "saidas": "#1c1f26"},
                labels={"mes": "Mês", "value": "Valor", "variable": "Tipo"},
            )
            fig_m.for_each_trace(
                lambda t: t.update(
                    marker_line_color="rgba(255,255,255,0.9)",
                    marker_line_width=1,
                    hovertemplate="<b>%{x|%m/%Y}</b><br>R$ %{y:,.2f}<extra></extra>",
                )
            )
            fig_m.update_xaxes(tickformat="%m/%Y")
            fig_m.update_yaxes(tickprefix="R$ ")
            fig_m = style_plot(fig_m, height=340)
            fig_m.update_layout(
                legend=dict(
                    orientation="h",
                    y=-0.18,
                    x=0.5,
                    xanchor="center",
                    title_text="",
                ),
                margin=dict(l=28, r=12, t=44, b=70),
            )
            st.plotly_chart(fig_m, use_container_width=True, config=PLOTLY_CONFIG)
    else:
        st.info("Cadastre ou importe movimentações para ver os gráficos.")

# ---------- ABA 3: Metas ----------
with aba[2]:
    st.subheader("Metas")
    st.markdown(
        """<div class="chart-intro"><strong>Minhas metas</strong><br>
        Defina objetivos simples, acompanhe quanto falta e veja em quanto tempo chega lá.</div>""",
        unsafe_allow_html=True,
    )

    with st.form("form_meta"):
        st.markdown("#### Adicionar meta")
        c1, c2, c3 = st.columns(3)
        with c1:
            data_m = st.date_input("Data da meta", value=date.today(), format="DD/MM/YYYY")
            nome_m = st.text_input("Nome da meta", placeholder="Ex.: Reserva, reforma, viagem")
            status_m = st.selectbox("Status", ["Em andamento", "Planejada", "Concluída"])
        with c2:
            valor_m = st.number_input("Valor necessário (R$)", value=0.0, min_value=0.0, step=0.01)
            atual_m = st.number_input("Quanto já tenho (R$)", value=0.0, min_value=0.0, step=0.01)
            aporte_m = st.number_input("Aporte mensal (R$)", value=0.0, min_value=0.0, step=0.01)
        with c3:
            prazo_m = st.text_input("Prazo desejado", placeholder="Ex.: Dezembro/2026")
            anot_m = st.text_area("Anotações")

        if st.form_submit_button("Salvar meta"):
            if not nome_m.strip():
                st.error("Informe o nome da meta.")
            elif valor_m <= 0:
                st.error("Informe o valor necessário.")
            else:
                try:
                    salvar_meta(data_m, nome_m.strip(), valor_m, atual_m, aporte_m, prazo_m.strip(), status_m, anot_m.strip())
                    st.success("Meta salva!")
                    st.rerun()
                except Exception as e:
                    st.error(mensagem_erro_usuario(e))

    if len(metas):
        st.markdown("#### Progresso das metas")
        for _, m in metas.iterrows():
            vm = max(float(m["valor_meta"]), 0)
            va = max(float(m["valor_atual"]), 0)
            am = max(float(m["aporte_mensal"]), 0)
            perc = min((va / vm) * 100, 100) if vm > 0 else 0
            falta = max(vm - va, 0)
            meses = math.ceil(falta / am) if falta > 0 and am > 0 else 0
            prev = f"Chega em {texto_meses(meses)}" if falta > 0 and am > 0 else ("Meta concluída" if falta <= 0 else "Defina um aporte mensal")

            st.markdown(
                f"""
<div class="goal-card">
    <div style="font-weight:850; font-size:1.05rem;">{escape(str(m['nome'] or 'Meta'))}</div>
    <div class="goal-meta">Status: {escape(str(m['status']))} • Prazo: {escape(str(m['prazo'] or '—'))}</div>
    <div class="goal-progress"><span style="width:{perc:.1f}%"></span></div>
    <div class="goal-meta">
        <strong>{perc:.0f}% concluída</strong><br>
        Tenho {brl(va)} de {brl(vm)}. Falta {brl(falta)}.<br>
        {prev}. Aporte: {brl(am)}.
    </div>
</div>
""",
                unsafe_allow_html=True,
            )
            if st.button("🗑️ Apagar meta", key=f"del_meta_{m['id']}"):
                excluir_meta(m["id"])
                st.rerun()
    else:
        st.info("Nenhuma meta cadastrada ainda.")

# ---------- ABA 4: Dívidas ----------
with aba[3]:
    st.subheader("Dívidas e Negociação")
    st.markdown(
        """<div class="chart-intro"><strong>Controle de dívidas</strong><br>
        Registre credor, saldo, parcela, prioridade e próxima ação.</div>""",
        unsafe_allow_html=True,
    )

    with st.form("form_divida"):
        st.markdown("#### Adicionar dívida ou acordo")
        c1, c2, c3 = st.columns(3)
        with c1:
            data_d = st.date_input("Data do registro", value=date.today(), format="DD/MM/YYYY")
            credor = st.text_input("Credor", placeholder="Ex.: Banco, cartão, loja")
            tipo_d = st.selectbox("Tipo", ["Cartão de crédito", "Empréstimo", "Conta atrasada", "Financiamento", "Acordo", "Outro"])
            prioridade = st.selectbox("Prioridade", ["Alta", "Média", "Baixa"])
        with c2:
            saldo_orig = st.number_input("Saldo original (R$)", value=0.0, min_value=0.0, step=0.01)
            desconto = st.number_input("Desconto (R$)", value=0.0, min_value=0.0, step=0.01)
            saldo_neg = st.number_input("Saldo negociado (R$)", value=0.0, min_value=0.0, step=0.01)
            parcela = st.number_input("Parcela possível (R$)", value=0.0, min_value=0.0, step=0.01)
        with c3:
            venc = st.date_input("Vencimento / próximo prazo", value=date.today(), format="DD/MM/YYYY")
            status_d = st.selectbox("Status", ["Mapear", "Negociar", "Acordada", "Em pagamento", "Quitada"])
            proxima = st.text_input("Próxima ação", placeholder="Ex.: ligar, pedir desconto")
            anot_d = st.text_area("Anotações")

        if st.form_submit_button("Salvar dívida"):
            saldo_final = saldo_neg if saldo_neg > 0 else max(saldo_orig - desconto, 0)
            if not credor.strip():
                st.error("Informe o credor.")
            elif saldo_final <= 0:
                st.error("Informe o saldo.")
            else:
                try:
                    salvar_divida(
                        data_d, credor.strip(), tipo_d, saldo_orig, desconto, saldo_final,
                        parcela, venc, prioridade, "", status_d, proxima.strip(), anot_d.strip(),
                    )
                    st.success("Dívida salva!")
                    st.rerun()
                except Exception as e:
                    st.error(mensagem_erro_usuario(e))

    if len(dividas):
        d_view = dividas.copy()
        d_view["saldo_base"] = d_view["saldo_negociado"].where(d_view["saldo_negociado"] > 0, d_view["saldo_original"])
        abertas = d_view[d_view["status"] != "Quitada"]
        total_aberto = abertas["saldo_base"].sum()
        parcelas_acord = abertas["parcela_possivel"].sum()
        economia = (d_view["saldo_original"] - d_view["saldo_base"]).clip(lower=0).sum()
        prioritarias = len(d_view[d_view["prioridade"] == "Alta"])

        st.markdown(
            f"""
<div class="debt-grid">
    <div class="metric-card"><div class="metric-label">Total aberto</div><div class="metric-value">{brl(total_aberto)}</div><div class="metric-foot">Saldo para negociar ou pagar</div></div>
    <div class="metric-card"><div class="metric-label">Parcelas acordadas</div><div class="metric-value">{brl(parcelas_acord)}</div><div class="metric-foot">Compromisso mensal possível</div></div>
    <div class="metric-card"><div class="metric-label">Economia prevista</div><div class="metric-value">{brl(economia)}</div><div class="metric-foot">Descontos registrados</div></div>
    <div class="metric-card"><div class="metric-label">Prioridade alta</div><div class="metric-value">{prioritarias}</div><div class="metric-foot">Dívidas que exigem atenção</div></div>
</div>
""",
            unsafe_allow_html=True,
        )

        with st.expander("⚡ E se eu quiser pagar mais rápido?", expanded=False):
            base = parcelas_acord if parcelas_acord > 0 else 0
            extra_d = st.slider("Quero pagar a mais por mês", 0, 5000, 0, 50, format="R$ %d")
            total_pag = base + extra_d
            if total_aberto > 0 and total_pag > 0:
                meses_atuais = math.ceil(total_aberto / base) if base > 0 else 0
                meses_novos = math.ceil(total_aberto / total_pag)
                ganho = max(meses_atuais - meses_novos, 0) if meses_atuais else 0
                st.markdown(
                    f"""
<div class="answer-grid">
    <div class="answer-card answer-care">
        <div class="answer-question">Prazo atual</div>
        <div class="answer-value">{texto_meses(meses_atuais) if meses_atuais else "Sem parcela"}</div>
        <div class="answer-action">No ritmo atual.</div>
    </div>
    <div class="answer-card answer-good">
        <div class="answer-question">Novo prazo</div>
        <div class="answer-value">{texto_meses(meses_novos)}</div>
        <div class="answer-action">Com {brl(total_pag)} por mês.</div>
    </div>
    <div class="answer-card answer-good">
        <div class="answer-question">Tempo ganho</div>
        <div class="answer-value">{texto_meses(ganho)}</div>
        <div class="answer-action">Estimativa simples (sem juros futuros).</div>
    </div>
</div>
""",
                    unsafe_allow_html=True,
                )
            else:
                st.info("Cadastre saldo e parcela para simular.")

        for _, d in d_view.iterrows():
            sb = d["saldo_base"]
            st.markdown(
                f"""
<div class="debt-item">
    <div>
        <div style="font-weight:850; font-size:1.05rem;">{escape(str(d['credor'] or 'Credor'))}</div>
        <div class="debt-meta">{escape(str(d['tipo']))} • Venc: {data_br(d['vencimento'])} • {escape(str(d['status']))} • Prioridade: {escape(str(d['prioridade']))}</div>
        <div class="debt-meta"><strong>Próxima ação:</strong> {escape(str(d['proxima_acao'] or '—'))}<br><strong>Anotações:</strong> {escape(str(d['anotacoes'] or '—'))}</div>
    </div>
    <div style="text-align:right;">
        <div class="negative" style="font-size:1.1rem; font-weight:900;">{brl(sb)}</div>
        <div class="debt-meta">Parcela: {brl(d['parcela_possivel'])}</div>
    </div>
</div>
""",
                unsafe_allow_html=True,
            )
            if st.button("🗑️ Apagar dívida", key=f"del_div_{d['id']}"):
                excluir_divida(d["id"])
                st.rerun()
    else:
        st.info("Nenhuma dívida cadastrada ainda.")

# ---------- ABA 5: Investimentos ----------
with aba[4]:
    st.subheader("Investimentos")
    with st.form("form_inv"):
        st.markdown("#### Adicionar investimento")
        c1, c2, c3 = st.columns(3)
        with c1:
            data_i = st.date_input("Data", value=date.today(), format="DD/MM/YYYY", key="di")
            tipo_i = st.selectbox(
                "Tipo",
                ["Reserva de emergência", "Tesouro Direto", "CDB", "LCI / LCA", "Fundo", "Ações", "FII", "Previdência", "Cripto", "Outro"],
            )
        with c2:
            valor_i = st.number_input("Valor investido (R$)", value=0.0, min_value=0.0, step=0.01)
            rent_i = st.text_input("Rentabilidade", placeholder="Ex.: 12% a.a. ou 105% CDI")
        with c3:
            desc_i = st.text_input("Descrição", placeholder="Ex.: CDB Banco X")
            status_i = st.selectbox("Status", ["Ativo", "Planejado", "Resgatado"])

        if st.form_submit_button("Salvar investimento"):
            if valor_i <= 0:
                st.error("Informe um valor maior que zero.")
            else:
                try:
                    salvar_investimento(data_i, tipo_i, valor_i, rent_i, desc_i, status_i)
                    st.success("Investimento salvo!")
                    st.rerun()
                except Exception as e:
                    st.error(mensagem_erro_usuario(e))

    if len(investimentos):
        total_ativos = investimentos[investimentos["status"] == "Ativo"]["valor"].sum()
        total_plan = investimentos[investimentos["status"] == "Planejado"]["valor"].sum()
        st.markdown(
            f"""
<div class="investment-grid">
    <div class="metric-card"><div class="metric-label">Patrimônio registrado</div><div class="metric-value">{brl(total_investido)}</div></div>
    <div class="metric-card"><div class="metric-label">Ativos</div><div class="metric-value">{brl(total_ativos)}</div></div>
    <div class="metric-card"><div class="metric-label">Planejados</div><div class="metric-value">{brl(total_plan)}</div></div>
</div>
""",
            unsafe_allow_html=True,
        )

        tipo_inv = investimentos.groupby("tipo", as_index=False)["valor"].sum().sort_values("valor", ascending=False)
        c1, c2 = st.columns(2)
        with c1:
            fig = px.pie(
                tipo_inv.head(5) if len(tipo_inv) > 5 else tipo_inv,
                names="tipo",
                values="valor",
                hole=0.58,
                title="Distribuição dos investimentos",
                color_discrete_sequence=["#1c1f26", "#d9ff00", "#7ed9b0", "#a8c4ff", "#b8f0d8", "#c5c9d4"],
            )
            if len(tipo_inv) > 5:
                # rebuild with Outras
                top_inv = tipo_inv.head(5).copy()
                rest = float(tipo_inv.iloc[5:]["valor"].sum())
                if rest > 0:
                    top_inv = pd.concat([top_inv, pd.DataFrame([{"tipo": "Outras", "valor": rest}])], ignore_index=True)
                fig = px.pie(
                    top_inv,
                    names="tipo",
                    values="valor",
                    hole=0.58,
                    title="Distribuição dos investimentos",
                    color_discrete_sequence=["#1c1f26", "#d9ff00", "#7ed9b0", "#a8c4ff", "#b8f0d8", "#c5c9d4"],
                )
            fig.update_traces(
                textposition="inside",
                textinfo="percent",
                textfont=dict(size=11, color="#1c1f26"),
                insidetextorientation="horizontal",
                marker=dict(line=dict(color="rgba(255,255,255,0.95)", width=2)),
                hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>",
            )
            fig = style_plot(fig, height=320)
            fig.update_layout(
                margin=dict(l=10, r=10, t=44, b=72),
                legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center", font=dict(size=11)),
            )
            st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)
        with c2:
            fig2 = px.bar(
                investimentos.groupby("status", as_index=False)["valor"].sum(),
                x="status", y="valor", color="status",
                title="Valores por status",
                color_discrete_map={"Ativo": "#0d9f6e", "Planejado": "#7ed9b0", "Resgatado": "#a8c4ff"},
                labels={"status": "Status", "valor": "Valor"},
            )
            fig2.update_traces(
                texttemplate="R$ %{y:,.0f}",
                textposition="outside",
                marker_line_color="rgba(255,255,255,0.9)",
                marker_line_width=1,
                hovertemplate="<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>",
            )
            fig2.update_yaxes(tickprefix="R$ ")
            fig2 = style_plot(fig2, height=320, show_legend=False)
            fig2.update_layout(margin=dict(l=40, r=20, t=60, b=40))
            st.plotly_chart(fig2, use_container_width=True, config=PLOTLY_CONFIG)

        for _, inv in investimentos.iterrows():
            st.markdown(
                f"""
<div class="investment-item">
    <div>
        <div style="font-weight:850;">{escape(str(inv['descricao'] or 'Investimento'))}</div>
        <div class="investment-meta">{escape(str(inv['tipo']))} • {data_br(inv['data'])} • {escape(str(inv['status']))}</div>
    </div>
    <div style="text-align:right;">
        <div class="positive" style="font-weight:900;">{brl(inv['valor'])}</div>
        <div class="investment-meta">Rentab.: {escape(str(inv['rentabilidade'] or '—'))}</div>
    </div>
</div>
""",
                unsafe_allow_html=True,
            )
            if st.button("🗑️ Apagar investimento", key=f"del_inv_{inv['id']}"):
                excluir_investimento(inv["id"])
                st.rerun()
    else:
        st.info("Nenhum investimento cadastrado ainda.")

# ---------- ABA 6: Histórico ----------
with aba[5]:
    st.subheader("Histórico")
    st.markdown(
        f"""<div class="history-summary"><strong>Relatório detalhado</strong><br>
        {len(df)} movimentações • {len(investimentos)} investimentos • {len(dividas)} dívidas • {len(metas)} metas</div>""",
        unsafe_allow_html=True,
    )

    st.download_button(
        "⬇️ Baixar relatório em PDF",
        data=gerar_pdf(df, investimentos, dividas, metas),
        file_name=f"relatorio-financeiro-{date.today().strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
    )

    if len(df):
        with st.expander("🧹 Limpar histórico completo", expanded=False):
            st.caption("Apaga todas as movimentações. Investimentos, dívidas e metas não são alterados.")
            conf = st.checkbox("Confirmo que quero apagar todas as movimentações")
            if st.button("Apagar todo o histórico", disabled=not conf):
                limpar_historico()
                st.success("Histórico limpo.")
                st.rerun()

        c1, c2, c3 = st.columns([2, 1, 1])
        busca = c1.text_input("Buscar", placeholder="Descrição, categoria ou pagamento")
        tipo_f = c2.selectbox("Tipo", ["Todos", "Entrada", "Saída"])
        cats_raw = df["categoria"].fillna("Sem categoria").replace("", "Sem categoria")
        cats = sorted({str(c) for c in cats_raw.unique() if str(c).lower() not in ("nan", "none", "")})
        cat_f = c3.selectbox("Categoria", ["Todas", *cats])

        df_h = df.copy()
        if busca.strip():
            termo = normalizar(busca)
            def _bate(r):
                texto = " ".join([
                    limpar_texto(r.get("descricao") if hasattr(r, "get") else r["descricao"]),
                    limpar_texto(r.get("categoria") if hasattr(r, "get") else r["categoria"]),
                    limpar_texto(r.get("cartao") if hasattr(r, "get") else r["cartao"]),
                ])
                return termo in normalizar(texto)
            try:
                df_h = df_h[df_h.apply(_bate, axis=1)]
            except Exception:
                pass
        if tipo_f != "Todos":
            df_h = df_h[df_h["tipo"] == tipo_f]
        if cat_f != "Todas":
            df_h = df_h[df_h["categoria"].fillna("Sem categoria").replace("", "Sem categoria") == cat_f]

        cols_show = ["data", "descricao", "categoria", "tipo", "cartao", "valor"]
        tabela = df_h[cols_show].copy()
        tabela = tabela.rename(columns={
            "data": "Data",
            "descricao": "Descricao",
            "categoria": "Categoria",
            "tipo": "Tipo",
            "cartao": "Pagamento",
            "valor": "Valor",
        })
        tabela["Data"] = tabela["Data"].astype(str).map(lambda v: data_br(v))
        tabela["Valor"] = tabela["Valor"].apply(lambda v: brl(v))
        st.dataframe(tabela, use_container_width=True, hide_index=True)
        st.caption(f"{len(df_h)} de {len(df)} movimentacoes exibidas.")

        records = df_h.to_dict(orient="records")
        for rec in records:
            try:
                valor_row = float(rec.get("valor") or 0)
            except Exception:
                valor_row = 0.0
            if valor_row != valor_row:
                valor_row = 0.0
            classe = "positive" if valor_row >= 0 else "negative"
            desc = limpar_texto(rec.get("descricao"), "Sem descricao")
            cat = limpar_texto(rec.get("categoria"), "Sem categoria")
            cart = limpar_texto(rec.get("cartao"), "-")
            data_txt = data_br(rec.get("data"))
            valor_txt = brl(valor_row)
            html_item = (
                '<div class="history-item">'
                "<div>"
                f'<div style="font-weight:850;">{escape(desc)}</div>'
                f'<div style="color:#8a90a0; font-size:0.88rem;">{escape(cat)} | {data_txt} | {escape(cart)}</div>'
                "</div>"
                f'<div class="{classe}">{valor_txt}</div>'
                "</div>"
            )
            st.markdown(html_item, unsafe_allow_html=True)
            try:
                rid = int(rec.get("id") or 0)
            except Exception:
                rid = 0
            if rid and st.button("Apagar", key=f"del_t_{rid}"):
                excluir_transacao(rid)
                st.rerun()

    else:
        st.info("Nenhum registro ainda.")

st.caption("Dashboard Financeiro • Visão clara e simples • Uso pessoal")
