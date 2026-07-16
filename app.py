import base64
import math
import re
import sqlite3
import textwrap
import zipfile
from datetime import date
from html import escape
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


st.set_page_config(
    page_title="Dashboard Financeiro",
    layout="wide",
    page_icon="📊",
    page_icon="DF",
)

# ====================== IMAGEM DE FUNDO ======================
DB_FILE = "financeiro.db"
BACKGROUND_IMAGE = Path(__file__).parent / "assets" / "background-blue-dunes.jpg"


def image_to_base64(path):
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode()


bg_base64 = image_to_base64(BACKGROUND_IMAGE)
bg_css = (
    f"url('data:image/jpeg;base64,{bg_base64}') center/cover fixed no-repeat"
    if bg_base64
    else "linear-gradient(160deg, #071b51, #183f73, #e09d9b)"
)

# ====================== ESTILO ======================
st.markdown(
    f"""
<style>
    @keyframes fadeIn {{
        from {{
            opacity: 0;
            transform: translateY(18px);
        }}
        to {{
            opacity: 1;
            transform: translateY(0);
        }}
    }}

    @keyframes float {{
        0%, 100% {{
            transform: translateY(0);
        }}
        50% {{
            transform: translateY(-6px);
        }}
    }}

    :root {{
        --black: #101d39;
        --dark: #15294b;
        --brown: #315d7b;
        --gold: #e1a09d;
        --sand: #dbe7ef;
        --cream: #f8fbff;
        --glass: rgba(238, 245, 250, 0.76);
        --glass-strong: rgba(247, 251, 255, 0.88);
        --border: rgba(255, 255, 255, 0.44);
        --shadow: 0 28px 80px rgba(2, 13, 42, 0.34);
    }}

    html, body, [class*="css"] {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}

    [data-testid="stAppViewContainer"] {{
        background: 
            linear-gradient(
                180deg,
                rgba(4, 18, 59, 0.22) 0%,
                rgba(23, 68, 112, 0.16) 46%,
                rgba(1, 10, 37, 0.52) 100%
            ),
            {bg_css};
        color: var(--black);
    }}

    [data-testid="stAppViewContainer"]::before {{
        content: "";
        position: fixed;
        inset: 0;
        pointer-events: none;
        background:
            radial-gradient(circle at 86% 17%, rgba(225, 160, 157, 0.22), transparent 30rem),
            radial-gradient(circle at 15% 72%, rgba(50, 151, 187, 0.20), transparent 30rem);
        z-index: 0;
    }}

    [data-testid="stHeader"],
    [data-testid="stToolbar"] {{
        background: transparent;
    }}

    .main .block-container {{
        position: relative;
        z-index: 1;
        max-width: 1180px;
        padding-top: 3rem;
        padding-bottom: 3rem;
        animation: fadeIn 0.9s ease-out;
    }}

    h1 {{
        color: var(--cream) !important;
        font-size: clamp(2.7rem, 6vw, 5.8rem);
        line-height: 0.94;
        font-weight: 850;
        letter-spacing: 0;
        text-shadow: 0 18px 42px rgba(0, 0, 0, 0.34);
        margin-bottom: 0.4rem;
    }}

    .hero-card h1 {{
        color: #f8fbff !important;
        -webkit-text-fill-color: #f8fbff !important;
    }}

    h2, h3 {{
        color: #f8fbff !important;
        font-weight: 760;
        letter-spacing: 0;
        text-shadow: 0 8px 22px rgba(1, 10, 37, 0.36);
    }}

    .hero {{
        min-height: 21rem;
        display: grid;
        grid-template-columns: minmax(0, 1.08fr) minmax(19rem, 0.72fr);
        gap: 1rem;
        align-items: stretch;
        margin-bottom: 1.2rem;
    }}

    .hero-card {{
        padding: clamp(1.35rem, 4vw, 2.35rem);
        border-radius: 30px;
        background: rgba(5, 25, 65, 0.62);
        border: 1px solid rgba(255, 255, 255, 0.30);
        box-shadow: var(--shadow);
        backdrop-filter: blur(16px) saturate(145%);
        -webkit-backdrop-filter: blur(16px) saturate(145%);
    }}

    .hero-top {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 1rem;
        margin-bottom: 2.9rem;
        color: rgba(248, 251, 255, 0.78);
        font-weight: 600;
    }}

    .hero-subtitle {{
        max-width: 35rem;
        color: rgba(248, 251, 255, 0.86);
        font-size: 1.04rem;
        line-height: 1.7;
        text-shadow: 0 8px 28px rgba(0, 0, 0, 0.28);
    }}

    .pill {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 2.4rem;
        padding: 0.35rem 1.15rem;
        border-radius: 999px;
        color: #122645;
        background: rgba(237, 174, 166, 0.96);
        font-weight: 750;
        box-shadow: 0 14px 34px rgba(0, 0, 0, 0.24);
    }}

    .hero-quote {{
        max-width: 37rem;
        margin: 1.65rem 0 0;
        padding-left: 1rem;
        border-left: 3px solid rgba(237, 174, 166, 0.92);
        color: rgba(248, 251, 255, 0.94);
        font-size: 1rem;
        font-weight: 650;
        line-height: 1.55;
    }}

    .hero-quote cite {{
        display: block;
        margin-top: 0.42rem;
        color: rgba(224, 233, 245, 0.74);
        font-size: 0.82rem;
        font-style: normal;
        font-weight: 750;
    }}

    .utility-card {{
        min-height: 21rem;
        padding: 1.25rem;
        border-radius: 30px;
        background: rgba(238, 245, 250, 0.84);
        border: 1px solid var(--border);
        box-shadow: var(--shadow);
        backdrop-filter: blur(18px) saturate(145%);
        -webkit-backdrop-filter: blur(18px) saturate(145%);
    }}

    .utility-head {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 0.75rem;
        margin-bottom: 0.9rem;
        color: rgba(16, 29, 57, 0.62);
        font-size: 0.78rem;
        font-weight: 850;
        text-transform: uppercase;
    }}

    .utility-head strong {{
        color: #15294b;
        font-size: 1.1rem;
        text-transform: none;
    }}

    .utility-grid {{
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 0.72rem;
    }}

    .utility-item {{
        min-height: 7.45rem;
        padding: 0.9rem;
        border-radius: 18px;
        background: rgba(255, 255, 255, 0.70);
        border: 1px solid rgba(255, 255, 255, 0.68);
        box-shadow: 0 12px 30px rgba(6, 29, 66, 0.10);
    }}

    .utility-label {{
        color: rgba(16, 29, 57, 0.58);
        font-size: 0.72rem;
        font-weight: 850;
        text-transform: uppercase;
    }}

    .utility-value {{
        margin-top: 1.38rem;
        color: #101d39;
        font-size: clamp(1.05rem, 1.9vw, 1.42rem);
        line-height: 1.05;
        font-weight: 850;
    }}

    .section-card {{
        margin-top: 0.7rem;
        padding: 1.25rem;
        border-radius: 26px;
        background: var(--glass);
        border: 1px solid var(--border);
        box-shadow: var(--shadow);
        backdrop-filter: blur(20px) saturate(155%);
        -webkit-backdrop-filter: blur(20px) saturate(155%);
    }}

    .metric-grid {{
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.9rem;
        margin: 1rem 0 1.2rem;
    }}

    .metric-card {{
        min-height: 8.6rem;
        padding: 1.15rem;
        border-radius: 22px;
        background: rgba(255, 250, 243, 0.82);
        border: 1px solid rgba(255, 255, 255, 0.68);
        box-shadow: 0 18px 44px rgba(64, 46, 26, 0.16);
        animation: float 4s ease-in-out infinite;
    }}

    .metric-label {{
        color: rgba(31, 29, 26, 0.62);
        font-size: 0.82rem;
        font-weight: 800;
        text-transform: uppercase;
    }}

    .metric-value {{
        margin-top: 1.35rem;
        color: var(--black);
        font-size: clamp(1.25rem, 2.4vw, 2rem);
        line-height: 1.05;
        font-weight: 850;
    }}

    .metric-foot {{
        margin-top: 0.55rem;
        color: rgba(31, 29, 26, 0.55);
        font-size: 0.84rem;
    }}

    div[data-testid="stTabs"] button {{
        border-radius: 999px;
        color: var(--dark);
        font-weight: 750;
        background: rgba(255, 250, 243, 0.62);
        border: 1px solid rgba(255, 255, 255, 0.58);
        padding: 0.35rem 1.05rem;
    }}

    div[data-testid="stTabs"] button[aria-selected="true"] {{
        color: var(--cream);
        background: var(--black);
    }}

    div[data-testid="stTabs"] [data-baseweb="tab-highlight"] {{
        background: transparent;
    }}

    div[data-testid="stTabs"] [data-baseweb="tab-list"],
    div[data-testid="stTabs"] [data-baseweb="tab-border"] {{
        background: transparent !important;
    }}

    div[data-testid="stForm"],
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        border-radius: 24px;
        border-color: rgba(255, 255, 255, 0.56);
        background: rgba(255, 250, 243, 0.52);
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
    }}

    div[data-testid="stForm"] label p,
    div[data-testid="stTextInput"] label p,
    div[data-testid="stNumberInput"] label p,
    div[data-testid="stDateInput"] label p,
    div[data-testid="stSelectbox"] label p,
    div[data-testid="stTextArea"] label p {{
        color: #332b23 !important;
        font-weight: 750 !important;
    }}

    div[data-baseweb="input"],
    div[data-baseweb="select"] > div,
    textarea {{
        overflow: hidden;
        min-height: 3rem;
        border: 1px solid rgba(95, 68, 38, 0.18) !important;
        border-radius: 15px !important;
        background: rgba(255, 252, 247, 0.94) !important;
        box-shadow: 0 10px 22px rgba(74, 49, 25, 0.09) !important;
        transition: border-color 0.2s ease, box-shadow 0.2s ease, transform 0.2s ease;
    }}

    div[data-baseweb="input"]:focus-within,
    div[data-baseweb="select"] > div:focus-within,
    textarea:focus {{
        border-color: rgba(178, 117, 48, 0.72) !important;
        box-shadow: 0 0 0 3px rgba(217, 155, 69, 0.18), 0 12px 24px rgba(74, 49, 25, 0.12) !important;
    }}

    div[data-baseweb="base-input"],
    div[data-baseweb="input"] > div {{
        border: 0 !important;
        background: transparent !important;
    }}

    input,
    textarea,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div {{
        color: #2b251f !important;
        font-weight: 620 !important;
    }}

    div[data-baseweb="select"] svg {{
        fill: #6f4d2d !important;
    }}

    input::placeholder,
    textarea::placeholder {{
        color: rgba(43, 37, 31, 0.46) !important;
    }}

    [data-testid="stNumberInput"] button {{
        border: 0 !important;
        border-left: 1px solid rgba(95, 68, 38, 0.14) !important;
        color: #4b3a29 !important;
        background: rgba(247, 232, 213, 0.88) !important;
    }}

    div[role="radiogroup"] label {{
        padding: 0.28rem 0.25rem;
    }}

    div[role="radiogroup"] label p {{
        color: rgba(248, 251, 255, 0.94) !important;
        font-weight: 750 !important;
    }}

    div[data-testid="stRadio"] > label p {{
        color: rgba(248, 251, 255, 0.94) !important;
        font-weight: 750 !important;
    }}

    div[data-testid="stForm"] h4 {{
        color: #101d39 !important;
        text-shadow: none;
    }}

    .stButton > button,
    [data-testid="stFormSubmitButton"] button {{
        min-height: 2.8rem;
        border: 0;
        border-radius: 999px;
        color: var(--cream);
        background: var(--black);
        font-weight: 800;
        box-shadow: 0 14px 30px rgba(31, 29, 26, 0.22);
        transition: 0.25s ease;
    }}

    .stButton > button:hover,
    [data-testid="stFormSubmitButton"] button:hover {{
        color: var(--cream);
        background: #342f27;
        transform: translateY(-2px);
        border: 0;
    }}

    [data-testid="stAlert"] {{
        border-radius: 20px;
        background: rgba(255, 250, 243, 0.74);
        border: 1px solid rgba(255, 255, 255, 0.58);
    }}

    div[data-testid="stExpander"] {{
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, 0.62);
        border-radius: 22px;
        background: rgba(238, 245, 250, 0.82);
        box-shadow: 0 18px 44px rgba(1, 10, 37, 0.18);
        backdrop-filter: blur(18px) saturate(140%);
        -webkit-backdrop-filter: blur(18px) saturate(140%);
    }}

    div[data-testid="stExpander"] summary p,
    div[data-testid="stExpander"] [data-testid="stCaptionContainer"] p,
    div[data-testid="stExpander"] p {{
        color: #101d39 !important;
    }}

    div[data-testid="stPlotlyChart"] {{
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, 0.74);
        border-radius: 22px;
        background: rgba(248, 251, 255, 0.94);
        box-shadow: 0 20px 48px rgba(1, 10, 37, 0.24);
    }}

    div[data-testid="stDataFrame"] {{
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, 0.74);
        border-radius: 18px;
        background: rgba(248, 251, 255, 0.96);
        box-shadow: 0 16px 36px rgba(1, 10, 37, 0.18);
    }}

    .chart-intro {{
        margin: 1.25rem 0 0.8rem;
        padding: 0.95rem 1.05rem;
        border-radius: 18px;
        color: rgba(248, 251, 255, 0.92);
        background: rgba(5, 25, 65, 0.72);
        border: 1px solid rgba(255, 255, 255, 0.34);
        box-shadow: 0 16px 38px rgba(1, 10, 37, 0.20);
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
    }}

    .chart-intro strong {{
        color: #ffffff;
    }}

    .history-item {{
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        gap: 0.85rem;
        align-items: center;
        margin-bottom: 0.78rem;
        padding: 1rem 1.08rem;
        border: 1px solid rgba(255, 255, 255, 0.72);
        border-radius: 18px;
        background: rgba(248, 251, 255, 0.92);
        box-shadow: 0 14px 34px rgba(1, 10, 37, 0.20);
        backdrop-filter: blur(14px);
        -webkit-backdrop-filter: blur(14px);
    }}

    .history-title {{
        color: #101d39;
        font-size: 1.02rem;
        font-weight: 800;
    }}

    .history-meta {{
        color: rgba(16, 29, 57, 0.68);
        font-size: 0.88rem;
        margin-top: 0.18rem;
    }}

    .positive {{
        color: #26704b;
        font-weight: 850;
    }}

    .negative {{
        color: #a3434c;
        font-weight: 850;
    }}

    .history-summary {{
        margin: 0.7rem 0 1.1rem;
        padding: 1rem 1.08rem;
        border: 1px solid rgba(255, 255, 255, 0.62);
        border-radius: 18px;
        color: rgba(248, 251, 255, 0.92);
        background: rgba(5, 25, 65, 0.68);
        box-shadow: 0 16px 38px rgba(1, 10, 37, 0.22);
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
    }}

    .history-summary strong {{
        color: #ffffff;
    }}

    div[data-testid="stDownloadButton"] button {{
        min-height: 2.8rem;
        border: 1px solid rgba(255, 255, 255, 0.62);
        border-radius: 999px;
        color: #101d39;
        background: rgba(248, 251, 255, 0.94);
        font-weight: 800;
        box-shadow: 0 14px 30px rgba(1, 10, 37, 0.20);
        transition: 0.25s ease;
    }}

    div[data-testid="stDownloadButton"] button:hover {{
        border-color: rgba(225, 160, 157, 0.88);
        color: #101d39;
        background: #ffffff;
        transform: translateY(-2px);
    }}

    .investment-grid {{
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.9rem;
        margin: 1rem 0 1.2rem;
    }}

    .investment-list {{
        display: grid;
        gap: 0.75rem;
        margin-top: 1rem;
    }}

    .investment-item {{
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        gap: 1rem;
        align-items: center;
        padding: 1rem;
        border: 1px solid rgba(255, 255, 255, 0.68);
        border-radius: 18px;
        background: rgba(255, 252, 247, 0.76);
        box-shadow: 0 12px 28px rgba(64, 46, 26, 0.11);
    }}

    .investment-title {{
        color: var(--black);
        font-size: 1.04rem;
        font-weight: 850;
    }}

    .investment-meta {{
        margin-top: 0.3rem;
        color: rgba(31, 29, 26, 0.65);
        font-size: 0.88rem;
        line-height: 1.45;
    }}

    .investment-value {{
        color: #3d6d45;
        font-size: 1.03rem;
        font-weight: 850;
        text-align: right;
    }}

    .investment-return {{
        margin-top: 0.3rem;
        color: rgba(31, 29, 26, 0.58);
        font-size: 0.83rem;
        text-align: right;
    }}

    .debt-grid {{
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.9rem;
        margin: 1rem 0 1.2rem;
    }}

    .debt-item {{
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        gap: 1rem;
        align-items: center;
        margin-bottom: 0.78rem;
        padding: 1rem;
        border: 1px solid rgba(255, 255, 255, 0.68);
        border-radius: 18px;
        background: rgba(255, 252, 247, 0.78);
        box-shadow: 0 12px 28px rgba(64, 46, 26, 0.11);
    }}

    .debt-title {{
        color: var(--black);
        font-size: 1.04rem;
        font-weight: 850;
    }}

    .debt-meta {{
        margin-top: 0.28rem;
        color: rgba(31, 29, 26, 0.65);
        font-size: 0.88rem;
        line-height: 1.45;
    }}

    .debt-note {{
        margin-top: 0.65rem;
        color: rgba(31, 29, 26, 0.74);
        font-size: 0.9rem;
        line-height: 1.55;
    }}

    .debt-value {{
        color: #cc4a5b;
        font-size: 1.03rem;
        font-weight: 850;
        text-align: right;
    }}

    .debt-return {{
        margin-top: 0.3rem;
        color: rgba(31, 29, 26, 0.58);
        font-size: 0.83rem;
        text-align: right;
    }}

    .answer-grid {{
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.9rem;
        margin: 1rem 0 1.2rem;
    }}

    .answer-card {{
        min-height: 9.2rem;
        padding: 1.05rem;
        border: 1px solid rgba(255, 255, 255, 0.68);
        border-radius: 20px;
        background: rgba(255, 252, 247, 0.82);
        box-shadow: 0 12px 28px rgba(64, 46, 26, 0.11);
    }}

    .answer-question {{
        color: rgba(31, 29, 26, 0.58);
        font-size: 0.78rem;
        font-weight: 850;
        text-transform: uppercase;
    }}

    .answer-value {{
        margin-top: 0.85rem;
        color: var(--black);
        font-size: clamp(1.1rem, 2.1vw, 1.55rem);
        line-height: 1.08;
        font-weight: 850;
    }}

    .answer-action {{
        margin-top: 0.65rem;
        color: rgba(31, 29, 26, 0.68);
        font-size: 0.88rem;
        line-height: 1.45;
    }}

    .answer-good {{
        border-color: rgba(13, 144, 111, 0.26);
        background: rgba(226, 248, 242, 0.86);
    }}

    .answer-care {{
        border-color: rgba(204, 138, 47, 0.28);
        background: rgba(255, 244, 222, 0.88);
    }}

    .answer-risk {{
        border-color: rgba(204, 74, 91, 0.28);
        background: rgba(255, 232, 235, 0.88);
    }}

    .goal-grid {{
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.9rem;
        margin: 1rem 0 1.2rem;
    }}

    .goal-card {{
        padding: 1.05rem;
        border: 1px solid rgba(255, 255, 255, 0.68);
        border-radius: 20px;
        background: rgba(255, 252, 247, 0.82);
        box-shadow: 0 12px 28px rgba(64, 46, 26, 0.11);
    }}

    .goal-title {{
        color: var(--black);
        font-size: 1.04rem;
        font-weight: 850;
    }}

    .goal-meta {{
        margin-top: 0.35rem;
        color: rgba(31, 29, 26, 0.66);
        font-size: 0.88rem;
        line-height: 1.45;
    }}

    .goal-progress {{
        overflow: hidden;
        height: 0.7rem;
        margin: 0.85rem 0 0.6rem;
        border-radius: 999px;
        background: rgba(8, 27, 51, 0.10);
    }}

    .goal-progress span {{
        display: block;
        height: 100%;
        border-radius: inherit;
        background: linear-gradient(90deg, #28c7b7, #0d906f);
    }}

    @media (max-width: 900px) {{
        .hero,
        .metric-grid,
        .answer-grid,
        .investment-grid,
        .debt-grid,
        .goal-grid,
        .investment-item,
        .debt-item,
        .history-item {{
            grid-template-columns: 1fr;
        }}

        .hero-top {{
            margin-bottom: 2.4rem;
        }}

        div[data-testid="stTabs"] [data-baseweb="tab-list"] {{
            flex-wrap: wrap;
            gap: 0.45rem;
        }}

        div[data-testid="stTabs"] button {{
            flex: 1 1 calc(50% - 0.45rem);
            justify-content: center;
        }}

    }}

    /* Aparencia inspirada em site institucional premium: limpo, azul profundo e verde agua. */
    :root {{
        --ink: #071426;
        --navy: #081b33;
        --deep: #0d2744;
        --blue: #17486f;
        --aqua: #28c7b7;
        --aqua-soft: #dff8f5;
        --mint: #8be6d4;
        --line: rgba(8, 27, 51, 0.10);
        --muted: #6b7687;
        --cream: #f7fbfa;
        --white: #ffffff;
        --card: rgba(255, 255, 255, 0.88);
        --glass: rgba(255, 255, 255, 0.76);
        --shadow: 0 24px 70px rgba(8, 27, 51, 0.16);
        --shadow-soft: 0 14px 38px rgba(8, 27, 51, 0.10);
    }}

    html, body, [class*="css"] {{
        font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}

    [data-testid="stAppViewContainer"] {{
        background:
            radial-gradient(circle at 8% 8%, rgba(40, 199, 183, 0.22), transparent 28rem),
            radial-gradient(circle at 88% 16%, rgba(23, 72, 111, 0.16), transparent 24rem),
            linear-gradient(180deg, #f8fcfb 0%, #edf7f6 44%, #f9fbfd 100%) !important;
        color: var(--ink);
    }}

    [data-testid="stAppViewContainer"]::before {{
        background:
            linear-gradient(115deg, rgba(8, 27, 51, 0.035) 0 1px, transparent 1px 100%),
            linear-gradient(245deg, rgba(40, 199, 183, 0.06), transparent 55%) !important;
        background-size: 46px 46px, auto !important;
    }}

    .main .block-container {{
        max-width: 1210px;
        padding-top: 2rem;
    }}

    .site-nav {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
        margin-bottom: 1rem;
        padding: 0.62rem 0.72rem;
        border: 1px solid rgba(8, 27, 51, 0.08);
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.82);
        box-shadow: var(--shadow-soft);
        backdrop-filter: blur(18px) saturate(150%);
        -webkit-backdrop-filter: blur(18px) saturate(150%);
    }}

    .brand-lockup {{
        display: flex;
        align-items: center;
        gap: 0.7rem;
        color: var(--navy);
        font-weight: 850;
    }}

    .brand-mark {{
        width: 2.55rem;
        height: 2.55rem;
        border-radius: 50%;
        display: grid;
        place-items: center;
        color: var(--white);
        background: linear-gradient(135deg, var(--navy), var(--aqua));
        box-shadow: 0 12px 30px rgba(40, 199, 183, 0.28);
    }}

    .nav-links {{
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 1.15rem;
        color: #516174;
        font-size: 0.88rem;
        font-weight: 720;
    }}

    .nav-cta {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 2.45rem;
        padding: 0.35rem 1rem;
        border-radius: 999px;
        color: var(--white);
        background: var(--navy);
        font-size: 0.88rem;
        font-weight: 800;
        box-shadow: 0 14px 32px rgba(8, 27, 51, 0.18);
    }}

    h1 {{
        color: var(--white) !important;
        font-size: clamp(2.55rem, 5.2vw, 5.25rem);
        line-height: 0.98;
        font-weight: 900;
        text-shadow: none;
    }}

    .hero-card h1 {{
        color: var(--white) !important;
        -webkit-text-fill-color: var(--white) !important;
    }}

    h2, h3 {{
        color: var(--navy) !important;
        font-weight: 820;
        text-shadow: none;
    }}

    .hero {{
        min-height: 24rem;
        grid-template-columns: minmax(0, 1.12fr) minmax(20rem, 0.72fr);
        gap: 1.05rem;
        margin-bottom: 1rem;
    }}

    .hero-card {{
        position: relative;
        overflow: hidden;
        padding: clamp(1.45rem, 4vw, 2.75rem);
        border-radius: 34px;
        background:
            radial-gradient(circle at 88% 18%, rgba(40, 199, 183, 0.42), transparent 16rem),
            linear-gradient(135deg, #071426 0%, #0b2543 52%, #123c5d 100%);
        border: 1px solid rgba(255, 255, 255, 0.14);
        box-shadow: var(--shadow);
        backdrop-filter: none;
        -webkit-backdrop-filter: none;
    }}

    .hero-card::after {{
        content: "";
        position: absolute;
        width: 18rem;
        height: 18rem;
        right: -6rem;
        bottom: -8rem;
        border-radius: 50%;
        border: 1px solid rgba(139, 230, 212, 0.24);
        background: rgba(139, 230, 212, 0.06);
        pointer-events: none;
    }}

    .hero-top,
    .hero-card > h1,
    .hero-card > p,
    .hero-card > blockquote {{
        position: relative;
        z-index: 1;
    }}

    .hero-top {{
        margin-bottom: 3rem;
        color: rgba(255, 255, 255, 0.72);
        font-weight: 750;
    }}

    .eyebrow {{
        display: inline-flex;
        align-items: center;
        gap: 0.42rem;
        color: rgba(255, 255, 255, 0.72);
        font-size: 0.9rem;
    }}

    .eyebrow::before {{
        content: "";
        width: 0.52rem;
        height: 0.52rem;
        border-radius: 50%;
        background: var(--aqua);
        box-shadow: 0 0 0 7px rgba(40, 199, 183, 0.12);
    }}

    .hero-subtitle {{
        max-width: 42rem;
        color: rgba(255, 255, 255, 0.82);
        line-height: 1.72;
        text-shadow: none;
    }}

    .pill {{
        color: #06231f;
        background: linear-gradient(135deg, #dff8f5, #8be6d4);
        font-weight: 820;
        box-shadow: 0 16px 34px rgba(40, 199, 183, 0.20);
    }}

    .hero-quote {{
        border-left-color: var(--aqua);
        color: rgba(255, 255, 255, 0.92);
    }}

    .hero-quote cite {{
        color: rgba(223, 248, 245, 0.78);
    }}

    .utility-card {{
        min-height: 24rem;
        padding: 1.35rem;
        border-radius: 34px;
        background: rgba(255, 255, 255, 0.90);
        border: 1px solid rgba(8, 27, 51, 0.08);
        box-shadow: var(--shadow);
    }}

    .utility-head {{
        color: var(--muted);
    }}

    .utility-head strong {{
        color: var(--navy);
        font-size: 1.15rem;
    }}

    .utility-item {{
        min-height: 8.2rem;
        border-radius: 22px;
        background: linear-gradient(180deg, #f8fbfd, #eef6f5);
        border: 1px solid rgba(8, 27, 51, 0.07);
        box-shadow: 0 14px 34px rgba(8, 27, 51, 0.08);
    }}

    .utility-label,
    .metric-label,
    .metric-foot,
    .history-meta,
    .investment-meta,
    .investment-return,
    .debt-meta,
    .debt-note,
    .debt-return,
    .answer-question,
    .answer-action,
    .goal-meta {{
        color: var(--muted);
    }}

    .utility-value,
    .metric-value {{
        color: var(--ink);
    }}

    .metric-card {{
        border-radius: 24px;
        background: rgba(255, 255, 255, 0.88);
        border: 1px solid rgba(8, 27, 51, 0.08);
        box-shadow: var(--shadow-soft);
        animation: none;
    }}

    div[data-testid="stTabs"] {{
        margin-top: 1.1rem;
    }}

    div[data-testid="stTabs"] button {{
        color: var(--navy);
        font-weight: 800;
        background: rgba(255, 255, 255, 0.86);
        border: 1px solid rgba(8, 27, 51, 0.10);
        padding: 0.42rem 1.08rem;
        box-shadow: 0 10px 26px rgba(8, 27, 51, 0.08);
    }}

    div[data-testid="stTabs"] button[aria-selected="true"] {{
        color: var(--white);
        background: var(--navy);
        border-color: var(--navy);
    }}

    div[data-testid="stForm"],
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        border-radius: 26px;
        border-color: rgba(8, 27, 51, 0.08);
        background: rgba(255, 255, 255, 0.72);
        box-shadow: var(--shadow-soft);
    }}

    div[data-testid="stForm"] label p,
    div[data-testid="stTextInput"] label p,
    div[data-testid="stNumberInput"] label p,
    div[data-testid="stDateInput"] label p,
    div[data-testid="stSelectbox"] label p,
    div[data-testid="stTextArea"] label p,
    div[data-testid="stRadio"] > label p,
    div[role="radiogroup"] label p {{
        color: var(--navy) !important;
        font-weight: 800 !important;
    }}

    div[data-baseweb="input"],
    div[data-baseweb="select"] > div,
    textarea {{
        border: 1px solid rgba(8, 27, 51, 0.10) !important;
        border-radius: 16px !important;
        background: rgba(255, 255, 255, 0.96) !important;
        box-shadow: 0 10px 24px rgba(8, 27, 51, 0.07) !important;
    }}

    div[data-baseweb="input"]:focus-within,
    div[data-baseweb="select"] > div:focus-within,
    textarea:focus {{
        border-color: rgba(40, 199, 183, 0.78) !important;
        box-shadow: 0 0 0 4px rgba(40, 199, 183, 0.14), 0 12px 24px rgba(8, 27, 51, 0.10) !important;
    }}

    input,
    textarea,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div {{
        color: var(--ink) !important;
    }}

    div[data-baseweb="select"] svg {{
        fill: var(--blue) !important;
    }}

    input::placeholder,
    textarea::placeholder {{
        color: rgba(7, 20, 38, 0.42) !important;
    }}

    [data-testid="stNumberInput"] button {{
        border-left: 1px solid rgba(8, 27, 51, 0.08) !important;
        color: var(--navy) !important;
        background: rgba(223, 248, 245, 0.88) !important;
    }}

    .stButton > button,
    [data-testid="stFormSubmitButton"] button {{
        color: var(--white);
        background: var(--navy);
        box-shadow: 0 14px 30px rgba(8, 27, 51, 0.18);
    }}

    .stButton > button:hover,
    [data-testid="stFormSubmitButton"] button:hover {{
        color: var(--white);
        background: #0f395f;
    }}

    [data-testid="stAlert"],
    div[data-testid="stExpander"],
    div[data-testid="stPlotlyChart"],
    div[data-testid="stDataFrame"] {{
        border: 1px solid rgba(8, 27, 51, 0.08);
        background: rgba(255, 255, 255, 0.88);
        box-shadow: var(--shadow-soft);
    }}

    div[data-testid="stExpander"] summary p,
    div[data-testid="stExpander"] [data-testid="stCaptionContainer"] p,
    div[data-testid="stExpander"] p {{
        color: var(--ink) !important;
    }}

    .chart-intro {{
        color: var(--ink);
        background: rgba(223, 248, 245, 0.72);
        border: 1px solid rgba(40, 199, 183, 0.22);
        box-shadow: var(--shadow-soft);
    }}

    .chart-intro strong,
    .history-summary strong,
    .history-title,
    .investment-title,
    .debt-title,
    .answer-value,
    .goal-title {{
        color: var(--navy);
    }}

    .history-item,
    .investment-item,
    .debt-item,
    .answer-card,
    .goal-card {{
        border: 1px solid rgba(8, 27, 51, 0.08);
        background: rgba(255, 255, 255, 0.92);
        box-shadow: var(--shadow-soft);
    }}

    .history-summary {{
        color: var(--ink);
        background: rgba(255, 255, 255, 0.80);
        border: 1px solid rgba(40, 199, 183, 0.20);
        box-shadow: var(--shadow-soft);
    }}

    .positive,
    .investment-value {{
        color: #0d906f;
    }}

    .negative {{
        color: #cc4a5b;
    }}

    .debt-value {{
        color: #cc4a5b;
    }}

    div[data-testid="stDownloadButton"] button {{
        color: var(--navy);
        background: rgba(255, 255, 255, 0.94);
        border: 1px solid rgba(8, 27, 51, 0.08);
        box-shadow: var(--shadow-soft);
    }}

    div[data-testid="stDownloadButton"] button:hover {{
        color: var(--navy);
        border-color: rgba(40, 199, 183, 0.60);
        background: #ffffff;
    }}

    @media (max-width: 900px) {{
        .site-nav {{
            align-items: flex-start;
            border-radius: 24px;
        }}

        .nav-links {{
            display: none;
        }}
    }}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<style>
    :root {
        --dash-ink: #111318;
        --dash-muted: #747985;
        --dash-line: rgba(17, 19, 24, 0.08);
        --dash-panel: rgba(246, 247, 249, 0.88);
        --dash-card: rgba(255, 255, 255, 0.92);
        --dash-lime: #d9ff00;
        --dash-lime-soft: #efffb4;
        --dash-blue: #8fb1ff;
        --dash-blue-soft: #e9efff;
        --dash-shadow: 0 26px 80px rgba(17, 19, 24, 0.12);
        --dash-shadow-soft: 0 14px 38px rgba(17, 19, 24, 0.08);
    }

    [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(circle at 12% 8%, rgba(255, 255, 255, 0.98), transparent 23rem),
            radial-gradient(circle at 86% 12%, rgba(217, 255, 0, 0.18), transparent 22rem),
            radial-gradient(circle at 74% 82%, rgba(143, 177, 255, 0.20), transparent 26rem),
            linear-gradient(145deg, #eef1f5 0%, #f7f8fa 46%, #e9edf2 100%) !important;
        color: var(--dash-ink);
    }

    [data-testid="stAppViewContainer"]::before {
        background:
            linear-gradient(120deg, rgba(255, 255, 255, 0.56), transparent 44%),
            repeating-linear-gradient(135deg, rgba(17, 19, 24, 0.018) 0 1px, transparent 1px 34px) !important;
    }

    .main .block-container {
        max-width: 1220px;
        padding-top: 1.6rem;
    }

    .site-nav {
        display: none;
    }

    .dashboard-shell {
        display: grid;
        grid-template-columns: 14.4rem minmax(0, 1fr);
        gap: 1rem;
        min-height: 31rem;
        margin-bottom: 1.2rem;
        padding: 0.55rem;
        border: 1px solid rgba(255, 255, 255, 0.72);
        border-radius: 34px;
        background: rgba(240, 242, 246, 0.72);
        box-shadow: var(--dash-shadow);
        backdrop-filter: blur(20px) saturate(140%);
        -webkit-backdrop-filter: blur(20px) saturate(140%);
    }

    .side-rail {
        display: flex;
        flex-direction: column;
        min-height: 30rem;
        padding: 1.15rem;
        border-radius: 28px;
        background: rgba(255, 255, 255, 0.78);
        border: 1px solid rgba(255, 255, 255, 0.88);
        box-shadow: var(--dash-shadow-soft);
    }

    .side-logo {
        display: flex;
        align-items: center;
        gap: 0.72rem;
        margin-bottom: 1.9rem;
        color: var(--dash-ink);
        font-size: 1.02rem;
        font-weight: 900;
    }

    .side-logo span:first-child {
        width: 2rem;
        height: 2rem;
        display: grid;
        place-items: center;
        border-radius: 50%;
        color: var(--dash-ink);
        background: var(--dash-lime);
        box-shadow: 0 12px 28px rgba(217, 255, 0, 0.26);
    }

    .side-menu {
        display: grid;
        gap: 0.48rem;
    }

    .side-menu span {
        display: flex;
        justify-content: space-between;
        align-items: center;
        min-height: 2.55rem;
        padding: 0 0.82rem;
        border-radius: 999px;
        color: #737783;
        font-size: 0.86rem;
        font-weight: 760;
    }

    .side-menu span.active {
        color: var(--dash-ink);
        background: #ffffff;
        box-shadow: 0 12px 28px rgba(17, 19, 24, 0.08);
    }

    .side-menu b {
        min-width: 1.45rem;
        height: 1.45rem;
        display: inline-grid;
        place-items: center;
        border-radius: 999px;
        background: var(--dash-blue);
        color: var(--dash-ink);
        font-size: 0.74rem;
    }

    .upgrade-card {
        margin-top: auto;
        padding: 1rem;
        border-radius: 24px;
        background: #ffffff;
        box-shadow: var(--dash-shadow-soft);
    }

    .upgrade-icon {
        width: 2.9rem;
        height: 2.9rem;
        display: grid;
        place-items: center;
        border-radius: 50%;
        color: var(--dash-ink);
        background: var(--dash-lime);
        font-size: 1.1rem;
        font-weight: 900;
    }

    .upgrade-title {
        margin-top: 1rem;
        color: var(--dash-ink);
        font-size: 1.08rem;
        font-weight: 900;
    }

    .upgrade-copy {
        margin-top: 0.3rem;
        color: var(--dash-muted);
        font-size: 0.8rem;
        line-height: 1.45;
    }

    .upgrade-button {
        margin-top: 0.95rem;
        min-height: 2.45rem;
        display: grid;
        place-items: center;
        border-radius: 999px;
        color: #ffffff;
        background: #050609;
        font-size: 0.82rem;
        font-weight: 850;
    }

    .hero {
        min-height: unset;
        display: block;
        margin: 0;
    }

    .hero-card,
    .utility-card {
        box-shadow: none;
    }

    .dashboard-stage {
        padding: clamp(1.05rem, 2vw, 1.45rem);
        border-radius: 30px;
        background: rgba(245, 246, 249, 0.72);
        border: 1px solid rgba(255, 255, 255, 0.76);
    }

    .stage-top {
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        align-items: flex-start;
        margin-bottom: 1.05rem;
    }

    .stage-title h1 {
        margin: 0.45rem 0 0.25rem;
        color: var(--dash-ink) !important;
        -webkit-text-fill-color: var(--dash-ink) !important;
        font-size: clamp(2.25rem, 5vw, 4.15rem);
        line-height: 0.96;
        letter-spacing: 0;
        text-shadow: none;
    }

    .stage-title p {
        max-width: 40rem;
        margin: 0.45rem 0 0;
        color: var(--dash-muted);
        font-size: 0.98rem;
        line-height: 1.65;
        font-weight: 620;
    }

    .stage-actions {
        display: flex;
        gap: 0.55rem;
        align-items: center;
        flex-shrink: 0;
    }

    .round-action {
        width: 2.65rem;
        height: 2.65rem;
        display: grid;
        place-items: center;
        border-radius: 50%;
        background: #ffffff;
        box-shadow: var(--dash-shadow-soft);
        font-weight: 900;
    }

    .user-chip {
        min-height: 2.65rem;
        display: flex;
        align-items: center;
        gap: 0.7rem;
        padding: 0.36rem 0.85rem 0.36rem 0.45rem;
        border-radius: 999px;
        background: #ffffff;
        box-shadow: var(--dash-shadow-soft);
        color: var(--dash-ink);
        font-size: 0.82rem;
        font-weight: 820;
        white-space: nowrap;
    }

    .avatar-dot {
        width: 2rem;
        height: 2rem;
        display: grid;
        place-items: center;
        border-radius: 50%;
        color: #ffffff;
        background: linear-gradient(135deg, #111318, #566070);
    }

    .hero-metrics {
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.85rem;
        margin: 1.15rem 0;
    }

    .metric-card {
        position: relative;
        overflow: hidden;
        min-height: 9.6rem;
        border-radius: 26px;
        background: rgba(255, 255, 255, 0.88);
        border: 1px solid rgba(255, 255, 255, 0.92);
        box-shadow: var(--dash-shadow-soft);
    }

    .metric-card::after {
        content: "↗";
        position: absolute;
        top: 1rem;
        right: 1rem;
        width: 2.35rem;
        height: 2.35rem;
        display: grid;
        place-items: center;
        border: 1px solid rgba(17, 19, 24, 0.22);
        border-radius: 50%;
        color: var(--dash-ink);
        font-weight: 900;
    }

    .metric-card.accent-card {
        background: linear-gradient(135deg, var(--dash-lime) 0%, #caff00 100%);
    }

    .metric-card.blue-card {
        background: linear-gradient(135deg, #ffffff 0%, var(--dash-blue-soft) 100%);
    }

    .metric-card.dark-card {
        background: #111318;
    }

    .metric-card.dark-card .metric-label,
    .metric-card.dark-card .metric-foot {
        color: rgba(255, 255, 255, 0.62);
    }

    .metric-card.dark-card .metric-value {
        color: #ffffff;
    }

    .metric-label {
        color: #7a7f89;
        font-size: 0.74rem;
        letter-spacing: 0;
    }

    .metric-value {
        color: var(--dash-ink);
        font-size: clamp(1.55rem, 3vw, 2.35rem);
    }

    .metric-foot {
        color: #787d88;
        font-size: 0.83rem;
    }

    .finance-quote {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        gap: 1rem;
        align-items: center;
        margin-top: 0.8rem;
        padding: 1rem 1.1rem;
        border-radius: 24px;
        background: #ffffff;
        box-shadow: var(--dash-shadow-soft);
    }

    .finance-quote strong {
        display: block;
        color: var(--dash-ink);
        font-size: 1rem;
    }

    .finance-quote span {
        color: var(--dash-muted);
        font-size: 0.84rem;
        font-weight: 700;
    }

    .finance-quote b {
        display: grid;
        place-items: center;
        min-width: 4.8rem;
        min-height: 2.55rem;
        border-radius: 999px;
        color: var(--dash-ink);
        background: var(--dash-lime);
        font-size: 1.1rem;
    }

    .chart-intro,
    .history-summary {
        color: var(--dash-ink);
        background: rgba(255, 255, 255, 0.86);
        border: 1px solid rgba(255, 255, 255, 0.92);
        border-radius: 24px;
        box-shadow: var(--dash-shadow-soft);
    }

    .chart-intro strong,
    .history-summary strong {
        color: var(--dash-ink);
    }

    div[data-testid="stPlotlyChart"],
    div[data-testid="stDataFrame"],
    div[data-testid="stExpander"],
    [data-testid="stAlert"],
    div[data-testid="stForm"],
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 26px;
        background: rgba(255, 255, 255, 0.88);
        border: 1px solid rgba(255, 255, 255, 0.92);
        box-shadow: var(--dash-shadow-soft);
    }

    div[data-testid="stTabs"] button {
        color: var(--dash-ink);
        background: rgba(255, 255, 255, 0.74);
        border: 1px solid rgba(255, 255, 255, 0.88);
        box-shadow: var(--dash-shadow-soft);
    }

    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: #ffffff;
        background: #111318;
        border-color: #111318;
    }

    h2, h3 {
        color: var(--dash-ink) !important;
    }

    .answer-grid,
    .indicator-grid {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.85rem;
        margin: 1rem 0 1.15rem;
    }

    .answer-card,
    .indicator-card,
    .goal-card,
    .investment-item,
    .debt-item,
    .history-item {
        border-radius: 24px;
        background: rgba(255, 255, 255, 0.90);
        border: 1px solid rgba(255, 255, 255, 0.92);
        box-shadow: var(--dash-shadow-soft);
    }

    .indicator-card {
        min-height: 8.7rem;
        padding: 1rem;
    }

    .indicator-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 0.7rem;
        color: var(--dash-muted);
        font-size: 0.76rem;
        font-weight: 850;
        text-transform: uppercase;
    }

    .indicator-icon {
        width: 2.2rem;
        height: 2.2rem;
        display: grid;
        place-items: center;
        border-radius: 50%;
        color: var(--dash-ink);
        background: var(--dash-lime);
    }

    .indicator-value {
        margin-top: 1rem;
        color: var(--dash-ink);
        font-size: clamp(1.25rem, 2.4vw, 1.85rem);
        font-weight: 900;
        line-height: 1;
    }

    .indicator-note {
        margin-top: 0.45rem;
        color: var(--dash-muted);
        font-size: 0.84rem;
        line-height: 1.4;
    }

    .progress-track {
        overflow: hidden;
        height: 0.55rem;
        margin-top: 0.8rem;
        border-radius: 999px;
        background: rgba(17, 19, 24, 0.08);
    }

    .progress-track span {
        display: block;
        height: 100%;
        border-radius: inherit;
        background: linear-gradient(90deg, var(--dash-lime), var(--dash-blue));
    }

    .visitor-panel {
        min-height: 24rem;
        padding: 1.15rem;
        border-radius: 26px;
        background: rgba(255, 255, 255, 0.90);
        border: 1px solid rgba(255, 255, 255, 0.92);
        box-shadow: var(--dash-shadow-soft);
    }

    .visitor-head {
        display: flex;
        justify-content: space-between;
        align-items: center;
        color: var(--dash-ink);
        font-weight: 900;
    }

    .visitor-head span {
        color: var(--dash-muted);
        font-size: 0.82rem;
        font-weight: 800;
    }

    .bubble-wrap {
        position: relative;
        min-height: 14rem;
        margin-top: 1rem;
    }

    .bubble {
        position: absolute;
        display: grid;
        place-items: center;
        border-radius: 50%;
        color: var(--dash-ink);
        text-align: center;
        font-weight: 900;
        box-shadow: 0 18px 42px rgba(17, 19, 24, 0.10);
    }

    .bubble small {
        display: block;
        margin-top: 0.2rem;
        color: rgba(17, 19, 24, 0.62);
        font-size: 0.72rem;
        font-weight: 760;
    }

    .bubble.income {
        width: 9.2rem;
        height: 9.2rem;
        left: 0.5rem;
        top: 0.8rem;
        background: var(--dash-lime);
    }

    .bubble.expense {
        width: 7.6rem;
        height: 7.6rem;
        right: 1.5rem;
        top: 2.3rem;
        background: var(--dash-blue);
    }

    .bubble.balance {
        width: 5.3rem;
        height: 5.3rem;
        left: 45%;
        bottom: 0.2rem;
        background: #ffffff;
        border: 1px solid var(--dash-line);
    }

    .target-list {
        display: grid;
        gap: 0.65rem;
        margin-top: 1rem;
    }

    .target-row {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        gap: 0.8rem;
        align-items: center;
        color: var(--dash-ink);
        font-size: 0.84rem;
        font-weight: 780;
    }

    .target-line {
        grid-column: 1 / -1;
        overflow: hidden;
        height: 0.48rem;
        border-radius: 999px;
        background: rgba(17, 19, 24, 0.08);
    }

    .target-line span {
        display: block;
        height: 100%;
        border-radius: inherit;
        background: var(--dash-lime);
    }

    .target-row:nth-child(2) .target-line span {
        background: var(--dash-blue);
    }

    .target-row:nth-child(3) .target-line span {
        background: #d7dbe2;
    }

    @media (max-width: 980px) {
        .dashboard-shell,
        .stage-top,
        .hero-metrics,
        .metric-grid,
        .answer-grid,
        .indicator-grid {
            grid-template-columns: 1fr;
        }

        .stage-top {
            display: grid;
        }

        .stage-actions {
            flex-wrap: wrap;
        }

        .side-rail {
            min-height: auto;
        }

        .upgrade-card {
            display: none;
        }
    }
</style>
""",
    unsafe_allow_html=True,
)

# ====================== BANCO ======================
DB_FILE = "financeiro.db"

# Versao sem login: o app abre direto no dashboard e usa banco local.


def mensagem_erro_usuario(erro):
    texto = str(erro or "").strip()
    if not texto:
        return "Nao conseguimos concluir agora. Tente novamente."
    if len(texto) > 180:
        return "Nao conseguimos concluir agora. Tente novamente em alguns minutos."
    return texto


def brl(valor):
    texto = f"R$ {valor:,.2f}"
    texto = f"R$ {float(valor or 0):,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def brl_compacto(valor):
def brl_curto(valor):
    numero = float(valor or 0)
    absoluto = abs(numero)
    if absoluto >= 1_000_000:
        texto = f"R$ {numero / 1_000_000:.1f} mi"
    elif absoluto >= 1_000:
        texto = f"R$ {numero / 1_000:.1f} mil"
    else:
        texto = f"R$ {numero:.0f}"
    return texto.replace(".", ",")


def pct(valor):
    return f"{float(valor or 0):.0f}%".replace(".", ",")


def data_br(valor):
    data_convertida = pd.to_datetime(valor, errors="coerce")
    data_convertida = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.isna(data_convertida):
        return "Sem data"
    return data_convertida.strftime("%d/%m/%Y")


def preparar_fluxo_mensal(df_transacoes):
    if df_transacoes.empty:
        return pd.DataFrame(columns=["mes", "entradas", "saidas", "saldo"])
def limpar_texto(valor, padrao=""):
    if pd.isna(valor):
        return padrao
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return padrao
    return texto

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
        .assign(saldo=lambda dados: dados["entradas"] - dados["saidas"])
        .sort_values("mes")
    )
def normalizar(valor):
    texto = limpar_texto(valor).lower()
    trocas = {
        "á": "a",
        "à": "a",
        "ã": "a",
        "â": "a",
        "é": "e",
        "ê": "e",
        "í": "i",
        "ó": "o",
        "ô": "o",
        "õ": "o",
        "ú": "u",
        "ç": "c",
    }
    for origem, destino in trocas.items():
        texto = texto.replace(origem, destino)
    return " ".join(texto.replace("_", " ").replace("-", " ").split())


def resumo_mes_recente(df_transacoes):
    df_fluxo = df_transacoes.copy()
    df_fluxo["data_convertida"] = pd.to_datetime(df_fluxo["data"], errors="coerce")
    df_fluxo = df_fluxo.dropna(subset=["data_convertida"])
    if df_fluxo.empty:
        return pd.DataFrame(), "Sem mês"

    mes_recente = df_fluxo["data_convertida"].max().to_period("M")
    df_mes = df_fluxo[df_fluxo["data_convertida"].dt.to_period("M") == mes_recente]
    return df_mes, mes_recente.strftime("%m/%Y")


def texto_meses(meses):
    if meses <= 0:
        return "agora"
    anos = meses // 12
    meses_restantes = meses % 12
    partes = []
    if anos:
        partes.append(f"{anos} ano" if anos == 1 else f"{anos} anos")
    if meses_restantes:
        partes.append(
            f"{meses_restantes} mês" if meses_restantes == 1 else f"{meses_restantes} meses"
        )
    return " e ".join(partes) if partes else "menos de 1 mês"


def pct(valor, casas=0):
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        numero = 0
    return f"{numero:.{casas}f}%".replace(".", ",")


def limitar_percentual(valor):
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        numero = 0
    return max(0, min(numero, 100))


def calcular_score_financeiro(
    entradas,
    saidas,
    saldo,
    investimentos,
    total_dividas_abertas,
    parcelas_dividas,
    progresso_metas,
):
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
    if total_dividas_abertas > 0:
        score -= min((total_dividas_abertas / max(entradas, 1)) * 8, 16)
    if saldo >= 0:
        score += 5
    return int(max(0, min(round(score), 100)))


def preparar_dados_dashboard(df_transacoes):
    if df_transacoes.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_base = df_transacoes.copy()
    df_base["data_convertida"] = pd.to_datetime(df_base["data"], errors="coerce")
    df_base["valor_abs"] = df_base["valor"].abs()
    df_base = df_base.dropna(subset=["data_convertida"]).sort_values("data_convertida")
    if df_base.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_base["mes"] = df_base["data_convertida"].dt.to_period("M").dt.to_timestamp()
    df_base["saldo_acumulado"] = df_base["valor"].cumsum()

    fluxo = (
        df_base.assign(
            Entradas=lambda dados: dados["valor"].clip(lower=0),
            Saídas=lambda dados: dados["valor"].clip(upper=0).abs(),
        )
        .groupby("mes", as_index=False)[["Entradas", "Saídas"]]
        .sum()
    )
    fluxo["Saldo"] = fluxo["Entradas"] - fluxo["Saídas"]

    categorias = (
        df_base.groupby("categoria", as_index=False)["valor_abs"]
        .sum()
        .sort_values("valor_abs", ascending=False)
    )

    pagamentos = (
        df_base.groupby("cartao", as_index=False)["valor_abs"]
        .sum()
        .sort_values("valor_abs", ascending=False)
    )
    pagamentos["cartao"] = pagamentos["cartao"].replace("", "Não informado")

    return fluxo, categorias, pagamentos


def gerar_modelo_excel():
    arquivo = BytesIO()
    workbook = Workbook()
    movimentacoes = workbook.active
    movimentacoes.title = "Movimentações"

    cabecalhos = [
        "Data",
        "Descrição",
        "Categoria",
        "Valor",
        "Tipo",
        "Forma de Pagamento",
    ]
    exemplos = [
        [date.today(), "Salário mensal", "Salário", 3500.00, "Entrada", "Conta corrente"],
        [date.today(), "Compras do mês", "Mercado", 250.00, "Saída", "Cartão"],
        [date.today(), "Conta de internet", "Contas", 99.90, "Saída", "Pix"],
    ]

    movimentacoes.append(cabecalhos)
    for exemplo in exemplos:
        movimentacoes.append(exemplo)

    tabela = Table(displayName="TabelaMovimentacoes", ref=f"A1:F{len(exemplos) + 1}")
    tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    movimentacoes.add_table(tabela)
    movimentacoes.freeze_panes = "A2"
    movimentacoes.auto_filter.ref = f"A1:F{len(exemplos) + 1}"

    larguras = {"A": 14, "B": 28, "C": 20, "D": 18, "E": 14, "F": 24}
    for coluna, largura in larguras.items():
        movimentacoes.column_dimensions[coluna].width = largura

    for celula in movimentacoes["A"][1:]:
        celula.number_format = "dd/mm/yyyy"
    for celula in movimentacoes["D"][1:]:
        celula.number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'

    resumo = workbook.create_sheet("Resumo de Valores")
    resumo.append(["Tabela de Valores", "Valor"])
    resumo.append(["Total de Entradas", '=SUMIF(Movimentações!E:E,"Entrada",Movimentações!D:D)'])
    resumo.append(["Total de Saídas", '=SUMIF(Movimentações!E:E,"Saída",Movimentações!D:D)'])
    resumo.append(["Saldo Previsto", "=B2-B3"])
    resumo.append(["Quantidade de Registros", "=COUNTA(Movimentações!A:A)-1"])

    resumo_tabela = Table(displayName="TabelaResumoValores", ref="A1:B5")
    resumo_tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    resumo.add_table(resumo_tabela)
    resumo.column_dimensions["A"].width = 28
    resumo.column_dimensions["B"].width = 22

    for celula in resumo["B"][1:4]:
        celula.number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'

    for planilha in [movimentacoes, resumo]:
        for celula in planilha[1]:
            celula.font = Font(bold=True, color="FFFFFF")
            celula.fill = PatternFill("solid", fgColor="15294B")
            celula.alignment = Alignment(horizontal="center")

    workbook.save(arquivo)
    arquivo.seek(0)
    return arquivo.getvalue()


def normalizar_coluna(nome):
    texto = str(nome).strip().lower()
def normalizar_coluna(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().lower()
    trocas = {
        "á": "a",
        "à": "a",
        "ã": "a",
        "â": "a",
        "é": "e",
        "ê": "e",
        "í": "i",
        "ó": "o",
        "ô": "o",
        "õ": "o",
        "ú": "u",
        "ç": "c",
    }
    for origem, destino in trocas.items():
        texto = texto.replace(origem, destino)
    return " ".join(texto.replace("_", " ").replace("-", " ").split())


def ler_planilha_movimentacoes(arquivo):
    nome = arquivo.name.lower()
    if nome.endswith(".csv"):
        try:
            return pd.read_csv(arquivo, sep=None, engine="python")
        except UnicodeDecodeError:
            arquivo.seek(0)
            return pd.read_csv(arquivo, sep=None, engine="python", encoding="latin-1")
def mensagem_erro_usuario(_erro):
    return "Nao conseguimos concluir agora. Tente novamente."

    if nome.endswith(".xlsx"):
        dados = arquivo.read()
        try:
            return pd.read_excel(BytesIO(dados), sheet_name=None, header=None, engine="openpyxl")
        except ValueError as erro:
            if "could not convert string to float" not in str(erro):
                raise
            arquivo_corrigido = corrigir_xlsx_malformado(dados)
            return pd.read_excel(arquivo_corrigido, sheet_name=None, header=None, engine="openpyxl")

    if nome.endswith(".xls"):
        return pd.read_excel(arquivo, sheet_name=None, header=None, engine="xlrd")

    raise ValueError("Envie uma planilha nos formatos CSV, XLSX ou XLS.")


def corrigir_xlsx_malformado(dados):
    origem = zipfile.ZipFile(BytesIO(dados), "r")
    destino = BytesIO()

    def corrigir_celula(correspondencia):
        celula = correspondencia.group(0)
        valor = re.search(r"<v>(.*?)</v>", celula, flags=re.DOTALL)
        if valor is None:
            return celula
        try:
            float(valor.group(1))
            return celula
        except ValueError:
            if 't="n"' in celula:
                return celula.replace('t="n"', 't="str"', 1)
            if re.search(r"<c\b[^>]*\bt=", celula):
                return celula
            return celula.replace(">", ' t="str">', 1)

    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as arquivo_saida:
        for item in origem.infolist():
            conteudo = origem.read(item.filename)
            if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                texto = conteudo.decode("utf-8")
                texto = re.sub(
                    r"<c\b[^>]*>.*?</c>",
                    corrigir_celula,
                    texto,
                    flags=re.DOTALL,
                )
                conteudo = texto.encode("utf-8")
            arquivo_saida.writestr(item, conteudo)

    origem.close()
    destino.seek(0)
    return destino


def converter_valor(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    if not texto:
        return None

    negativo = texto.startswith("(") and texto.endswith(")")
    negativo = texto.startswith("(") and texto.endswith(")") or texto.startswith("-")
    texto = (
        texto.replace("R$", "")
        .replace(" ", "")
        .replace("\u00a0", "")
        .replace("(", "")
        .replace(")", "")
    )
    texto = "".join(caractere for caractere in texto if caractere.isdigit() or caractere in ",.-")

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        numero = float(texto)
    except ValueError:
        return None

    return -abs(numero) if negativo else numero


def corrigir_xlsx_malformado(arquivo):
    if hasattr(arquivo, "seek"):
        arquivo.seek(0)
    return arquivo


def celula_planilha(df_aba, linha, coluna, padrao=""):
    try:
        valor = df_aba.iat[linha, coluna]
    except Exception:
        return padrao
    return valor if not pd.isna(valor) else padrao


def texto_planilha(valor, padrao=""):
    if pd.isna(valor):
        return padrao
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return padrao
    return texto


MESES_PLANILHA = {
    "janeiro": 1,
    "fevereiro": 2,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}

PALAVRAS_ENTRADA_PLANILHA = {
    "entrada",
    "receita",
    "recebimento",
    "salario",
    "renda",
    "freelance",
    "reembolso",
    "venda",
}
PALAVRAS_ENTRADA_PLANILHA = ["entrada", "receita", "salario", "renda", "freelance", "reembolso"]
PALAVRAS_SAIDA_PLANILHA = ["saida", "despesa", "debito", "gasto", "conta", "cartao"]

PALAVRAS_SAIDA_PLANILHA = {
    "saida",
    "despesa",
    "debito",
    "credito",
    "cartao",
    "gasto",
    "fixo",
    "parcelado",
}


def celula_planilha(df_planilha, linha, coluna):
    if linha < 0 or coluna < 0 or linha >= len(df_planilha.index) or coluna >= len(df_planilha.columns):
        return None
    return df_planilha.iat[linha, coluna]


def texto_planilha(valor, padrao=""):
    if pd.isna(valor):
        return padrao
    texto = str(valor).strip()
    return texto if texto and texto.lower() != "nan" else padrao


def data_planilha_mensal(valor, numero_mes):
    texto_data = texto_planilha(valor)
    data_sem_ano = re.fullmatch(r"(\d{1,2})/(\d{1,2})", texto_data)
    if data_sem_ano:
        dia = int(data_sem_ano.group(1))
        ano = date.today().year
        ultimo_dia = pd.Period(year=ano, month=numero_mes, freq="M").days_in_month
        return date(ano, numero_mes, min(dia, ultimo_dia)).isoformat()

    if isinstance(valor, (int, float)) and valor > 30000:
        data_convertida = pd.to_datetime(valor, unit="D", origin="1899-12-30", errors="coerce")
    else:
        data_convertida = pd.to_datetime(valor, errors="coerce", dayfirst=True)

def data_planilha_mensal(nome_aba, valor_data=None):
    mes = MESES_PLANILHA.get(normalizar_coluna(nome_aba), date.today().month)
    data_convertida = pd.to_datetime(valor_data, errors="coerce", dayfirst=True)
    if pd.isna(data_convertida):
        hoje = date.today()
        ano = hoje.year
        return date(ano, numero_mes, 1).isoformat()
        return date(date.today().year, mes, 1).isoformat()
    return data_convertida.date().isoformat()

    ano = data_convertida.year if data_convertida.year > 1970 else date.today().year
    ultimo_dia = pd.Period(year=ano, month=numero_mes, freq="M").days_in_month
    dia = min(data_convertida.day, ultimo_dia)
    return date(ano, numero_mes, dia).isoformat()

def texto_contexto_planilha(df_aba, linha, col_inicio, col_fim):
    textos = []
    for idx_linha in range(max(0, linha - 4), linha + 1):
        for idx_coluna in range(max(0, col_inicio), min(len(df_aba.columns), col_fim + 1)):
            textos.append(normalizar_coluna(celula_planilha(df_aba, idx_linha, idx_coluna)))
    return " ".join(textos)

def texto_contexto_planilha(df_planilha, linha, coluna_inicio, coluna_fim):
    pedacos = []
    linha_inicial = max(0, linha - 4)
    linha_final = min(len(df_planilha.index), linha + 1)
    coluna_inicial = max(0, coluna_inicio - 2)
    coluna_final = min(len(df_planilha.columns), coluna_fim + 3)

    for numero_linha in range(linha_inicial, linha_final):
        for numero_coluna in range(coluna_inicial, coluna_final):
            texto = texto_planilha(celula_planilha(df_planilha, numero_linha, numero_coluna))
            if texto:
                pedacos.append(normalizar_coluna(texto))

    return " ".join(pedacos)


def inferir_tipo_mensal(tipo_original, descricao, categoria, contexto):
    texto_linha = normalizar_coluna(
        " ".join(
            [
                texto_planilha(tipo_original),
                texto_planilha(descricao),
                texto_planilha(categoria),
            ]
        )
    )
    texto_contexto = normalizar_coluna(contexto)

    if any(palavra in texto_linha for palavra in PALAVRAS_ENTRADA_PLANILHA):
        return "Entrada"
    if any(palavra in texto_linha for palavra in PALAVRAS_SAIDA_PLANILHA):
def inferir_tipo_mensal(descricao="", categoria="", tipo="", contexto=""):
    texto = normalizar_coluna(" ".join([str(descricao), str(categoria), str(tipo), str(contexto)]))
    texto_direto = normalizar_coluna(" ".join([str(descricao), str(categoria), str(tipo)]))
    if "saida" in texto or any(palavra in texto_direto for palavra in ["debito", "credito", "despesa", "gasto"]):
        return "Saída"
    if "entradas" in texto_contexto or "receitas" in texto_contexto:
    if any(palavra in texto for palavra in PALAVRAS_ENTRADA_PLANILHA):
        return "Entrada"
    if any(palavra in texto for palavra in PALAVRAS_SAIDA_PLANILHA):
        return "Saída"
    return "Saída"


def adicionar_movimentacao_mensal(
    linhas,
    df_mes,
    linha,
    colunas,
    numero_mes,
    tipo,
    categoria_padrao,
    pagamento_padrao,
):
    valor_original = converter_valor(celula_planilha(df_mes, linha, colunas["valor"]))
    if valor_original is None or valor_original == 0:
        return False

    descricao = texto_planilha(
        celula_planilha(df_mes, linha, colunas.get("descricao", -1)),
        "Movimentação importada",
    )
    if normalizar_coluna(descricao) in {
        "total",
        "total de gastos",
        "total de fixos",
        "total de cartao de credito",
    }:
        return False

    categoria = texto_planilha(
        celula_planilha(df_mes, linha, colunas.get("categoria", -1)),
        categoria_padrao,
    )
    pagamento = texto_planilha(
        celula_planilha(df_mes, linha, colunas.get("pagamento", -1)),
        pagamento_padrao,
    )
    data_original = celula_planilha(df_mes, linha, colunas.get("data", -1))
    data_final = data_planilha_mensal(data_original, numero_mes)
    valor_final = abs(valor_original) if tipo == "Entrada" else -abs(valor_original)

def adicionar_movimentacao_mensal(linhas, data_mov, descricao, categoria, valor, tipo, cartao):
    if valor is None or valor == 0:
        return
    tipo_final = "Entrada" if tipo == "Entrada" else "Saída"
    valor_final = abs(valor) if tipo_final == "Entrada" else -abs(valor)
    linhas.append(
        {
            "data": data_final,
            "descricao": descricao,
            "categoria": categoria,
            "data": data_mov,
            "descricao": texto_planilha(descricao, "Importado da planilha"),
            "categoria": texto_planilha(categoria, "Importado"),
            "valor": valor_final,
            "tipo": tipo,
            "cartao": pagamento,
            "tipo": tipo_final,
            "cartao": texto_planilha(cartao, "Planilha"),
        }
    )
    return True


def adicionar_blocos_de_entrada(linhas, df_mes, numero_mes):
    adicionadas = 0
    for linha_bloco in range(len(df_mes.index)):
        for coluna_bloco in range(len(df_mes.columns)):
            if normalizar_coluna(celula_planilha(df_mes, linha_bloco, coluna_bloco)) != "entradas":
                continue
def adicionar_blocos_de_entrada(linhas, nome_aba, df_aba, linha_idx, col_descricao, col_valor):
    contexto = texto_contexto_planilha(df_aba, linha_idx, col_descricao, col_valor + 2)
    for idx in range(linha_idx + 1, len(df_aba.index)):
        descricao = texto_planilha(celula_planilha(df_aba, idx, col_descricao))
        valor = converter_valor(celula_planilha(df_aba, idx, col_valor))
        desc_norm = normalizar_coluna(descricao)
        if not descricao and valor is None:
            break
        if desc_norm.startswith("total"):
            break
        if valor is None or valor == 0:
            continue
        tipo = inferir_tipo_mensal(descricao, "Receita", "Entrada", contexto)
        adicionar_movimentacao_mensal(
            linhas,
            data_planilha_mensal(nome_aba),
            descricao,
            "Receita",
            valor,
            tipo,
            "Planilha",
        )

            coluna_descricao = coluna_bloco
            coluna_valor = coluna_bloco + 1
            linha_inicio = linha_bloco + 1

            for linha_cabecalho in range(linha_bloco + 1, min(linha_bloco + 4, len(df_mes.index))):
                cabecalho_nome = normalizar_coluna(celula_planilha(df_mes, linha_cabecalho, coluna_bloco))
                if cabecalho_nome in {"nome", "descricao"}:
                    coluna_descricao = coluna_bloco
                    coluna_valor_cabecalho = next(
                        (
                            coluna
                            for coluna in range(coluna_bloco + 1, min(coluna_bloco + 5, len(df_mes.columns)))
                            if normalizar_coluna(celula_planilha(df_mes, linha_cabecalho, coluna)) == "valor"
                        ),
                        None,
                    )
                    if coluna_valor_cabecalho is not None:
                        coluna_valor = coluna_valor_cabecalho
                    linha_inicio = linha_cabecalho + 1
                    break

            if coluna_valor >= len(df_mes.columns):
                continue

            linhas_vazias = 0
            for numero_linha in range(linha_inicio, len(df_mes.index)):
                descricao = texto_planilha(celula_planilha(df_mes, numero_linha, coluna_descricao))
                descricao_normalizada = normalizar_coluna(descricao)
                valor = converter_valor(celula_planilha(df_mes, numero_linha, coluna_valor))

                if descricao_normalizada in {"saidas", "saida", "investimentos", "reserva"}:
                    break
                if descricao_normalizada.startswith("total"):
                    break

                if not descricao and (valor is None or valor == 0):
                    linhas_vazias += 1
                    if linhas_vazias >= 2:
                        break
                    continue

                linhas_vazias = 0
                if adicionar_movimentacao_mensal(
                    linhas,
                    df_mes,
                    numero_linha,
                    {"descricao": coluna_descricao, "valor": coluna_valor},
                    numero_mes,
                    "Entrada",
                    "Receita",
                    "Planilha mensal",
                ):
                    adicionadas += 1

    return adicionadas


def preparar_modelo_organizacao_financeira(planilhas):
    abas_mensais = {
        nome: (df_mes, MESES_PLANILHA[normalizar_coluna(nome)])
        for nome, df_mes in planilhas.items()
        if normalizar_coluna(nome) in MESES_PLANILHA
    }
    if not abas_mensais:
        return None

    linhas = []
    for nome_mes, (df_mes, numero_mes) in abas_mensais.items():
        tabelas_encontradas = []
        for numero_linha in range(len(df_mes.index)):
            colunas_nome = [
                coluna
                for coluna in range(len(df_mes.columns))
                if normalizar_coluna(celula_planilha(df_mes, numero_linha, coluna))
                in {"nome", "descricao"}
            ]
            colunas_valor = [
                coluna
                for coluna in range(len(df_mes.columns))
                if normalizar_coluna(celula_planilha(df_mes, numero_linha, coluna)) == "valor"
            ]

            for coluna_nome in colunas_nome:
                coluna_valor = next(
                    (
                        coluna
                        for coluna in colunas_valor
                        if coluna_nome < coluna <= coluna_nome + 8
                    ),
                    None,
                )
                if coluna_valor is None:
    for nome_aba, df_aba in planilhas.items():
        if MESES_PLANILHA.get(normalizar_coluna(nome_aba)) is None:
            continue
        for linha_idx in range(len(df_aba.index)):
            for col_idx in range(len(df_aba.columns) - 1):
                if normalizar_coluna(celula_planilha(df_aba, linha_idx, col_idx)) != "entradas":
                    continue

                colunas = {"descricao": coluna_nome, "valor": coluna_valor}
                for chave in ("data", "tipo", "categoria"):
                    coluna = next(
                        (
                            coluna
                            for coluna in range(coluna_nome + 1, coluna_valor)
                            if normalizar_coluna(
                                celula_planilha(df_mes, numero_linha, coluna)
                            )
                            == chave
                        ),
                        None,
                ultima_descricao = ""
                for idx in range(linha_idx + 1, min(len(df_aba.index), linha_idx + 14)):
                    descricao = texto_planilha(celula_planilha(df_aba, idx, col_idx))
                    valor_original = converter_valor(celula_planilha(df_aba, idx, col_idx + 1))
                    desc_norm = normalizar_coluna(descricao)
                    if desc_norm.startswith("total") or desc_norm == "saidas":
                        break
                    if descricao:
                        ultima_descricao = descricao
                    if valor_original is None or valor_original <= 0:
                        continue
                    adicionar_movimentacao_mensal(
                        linhas,
                        data_planilha_mensal(nome_aba),
                        descricao or ultima_descricao or "Entrada da planilha",
                        "Receita",
                        valor_original,
                        "Entrada",
                        "Planilha",
                    )
                    if coluna is not None:
                        colunas[chave] = coluna
                contexto = texto_contexto_planilha(df_mes, numero_linha, coluna_nome, coluna_valor)
        for linha_idx in range(len(df_aba.index)):
            valores_linha = [normalizar_coluna(celula_planilha(df_aba, linha_idx, col)) for col in range(len(df_aba.columns))]
            if "valor" not in valores_linha:
                continue
            col_descricao = next((i for i, item in enumerate(valores_linha) if item in ["descricao", "nome"]), None)
            if col_descricao is None:
                continue
            col_valor = next((i for i, item in enumerate(valores_linha) if item == "valor" and i > col_descricao), None)
            if col_valor is None:
                continue
            col_tipo = next((i for i, item in enumerate(valores_linha) if item in ["tipo", "pagamento"] and i > col_descricao), None)
            col_categoria = next((i for i, item in enumerate(valores_linha) if item == "categoria" and i > col_descricao), None)
            col_data = next((i for i, item in enumerate(valores_linha) if item == "data" and i > col_descricao), None)
            contexto = texto_contexto_planilha(df_aba, linha_idx, col_descricao, col_valor + 3)

                # Blocos laterais de "Entradas" têm subtabelas de resumo logo abaixo
                # ("Saídas", "Investimentos e reserva"). Eles são lidos por uma
                # rotina própria para evitar que totais sejam importados como gastos.
                if (
                    normalizar_coluna(celula_planilha(df_mes, numero_linha, coluna_nome))
                    == "descricao"
                    and ("entradas" in contexto or "receitas" in contexto)
                ):
            for idx in range(linha_idx + 1, len(df_aba.index)):
                descricao = texto_planilha(celula_planilha(df_aba, idx, col_descricao))
                valor_original = converter_valor(celula_planilha(df_aba, idx, col_valor))
                desc_norm = normalizar_coluna(descricao)
                if not descricao and valor_original is None:
                    break
                if not descricao:
                    continue

                tabelas_encontradas.append((numero_linha, colunas, contexto))

        for linha_cabecalho, colunas, contexto in tabelas_encontradas:
            linhas_vazias = 0
            for numero_linha in range(linha_cabecalho + 1, len(df_mes.index)):
                descricao = texto_planilha(celula_planilha(df_mes, numero_linha, colunas["descricao"]))
                valor = converter_valor(celula_planilha(df_mes, numero_linha, colunas["valor"]))
                if not descricao and (valor is None or valor == 0):
                    linhas_vazias += 1
                    if linhas_vazias >= 2:
                        break
                if desc_norm.startswith("total") or desc_norm in ["saidas", "entradas", "investimentos", "reserva"]:
                    break
                if valor_original is None or valor_original == 0:
                    continue
                linhas_vazias = 0
                tipo_original = texto_planilha(celula_planilha(df_aba, idx, col_tipo), "Planilha") if col_tipo is not None else "Planilha"
                categoria = texto_planilha(celula_planilha(df_aba, idx, col_categoria), "Receita" if "entrada" in contexto else "Outros") if col_categoria is not None else ("Receita" if "entrada" in contexto else "Outros")
                tipo = inferir_tipo_mensal(descricao, categoria, tipo_original, contexto)
                data_mov = data_planilha_mensal(nome_aba, celula_planilha(df_aba, idx, col_data) if col_data is not None else None)
                adicionar_movimentacao_mensal(linhas, data_mov, descricao, categoria, valor_original, tipo, tipo_original)
    return pd.DataFrame(linhas).drop_duplicates() if linhas else pd.DataFrame()

                descricao_normalizada = normalizar_coluna(descricao)
                if descricao_normalizada.startswith("total"):
                    continue

                tipo_original = texto_planilha(
                    celula_planilha(df_mes, numero_linha, colunas.get("tipo", -1)),
                    "Débito",
                )
                categoria_original = texto_planilha(
                    celula_planilha(df_mes, numero_linha, colunas.get("categoria", -1)),
                    "",
                )
                tipo = inferir_tipo_mensal(tipo_original, descricao, categoria_original, contexto)
                adicionar_movimentacao_mensal(
                    linhas,
                    df_mes,
                    numero_linha,
                    {
                        **colunas,
                        "pagamento": colunas.get("tipo", -1),
                    },
                    numero_mes,
                    tipo,
                    "Receita" if tipo == "Entrada" else "Outros",
                    tipo_original,
                )
def ler_planilha_movimentacoes(arquivo):
    nome = getattr(arquivo, "name", "").lower()
    if nome.endswith(".csv"):
        try:
            return pd.read_csv(arquivo, sep=None, engine="python")
        except UnicodeDecodeError:
            arquivo.seek(0)
            return pd.read_csv(arquivo, sep=None, engine="python", encoding="latin-1")
    arquivo = corrigir_xlsx_malformado(arquivo)
    dados = arquivo.read()
    if nome.endswith(".xlsx"):
        return pd.read_excel(BytesIO(dados), sheet_name=None, header=None, engine="openpyxl")
    if nome.endswith(".xls"):
        return pd.read_excel(BytesIO(dados), sheet_name=None, header=None)
    raise ValueError("Envie uma planilha CSV, XLSX ou XLS.")

        adicionar_blocos_de_entrada(linhas, df_mes, numero_mes)

    if not linhas:
        return pd.DataFrame(linhas), 0

    df_linhas = pd.DataFrame(linhas)
    df_linhas = df_linhas.drop_duplicates(
        subset=["data", "descricao", "valor", "tipo"],
    ).reset_index(drop=True)
    return df_linhas, 0


def preparar_movimentacoes_importadas(df_planilha):
    if isinstance(df_planilha, dict):
        modelo_organizado = preparar_modelo_organizacao_financeira(df_planilha)
        if modelo_organizado is not None:
            return modelo_organizado

        primeira_aba = next(iter(df_planilha.values()), pd.DataFrame())
        if primeira_aba.empty:
def preparar_movimentacoes_importadas(dados):
    if isinstance(dados, dict):
        mensal = preparar_modelo_organizacao_financeira(dados)
        if not mensal.empty:
            return mensal, 0
        primeira = next(iter(dados.values()), pd.DataFrame())
        if primeira.empty:
            return pd.DataFrame(), 0
        primeira_aba = primeira_aba.copy()
        primeira_aba.columns = primeira_aba.iloc[0]
        df_planilha = primeira_aba.iloc[1:].reset_index(drop=True)
        primeira = primeira.copy()
        primeira.columns = primeira.iloc[0]
        dados = primeira.iloc[1:].reset_index(drop=True)

    aliases = {
        "data": ["data", "dt", "dia", "date"],
        "descricao": ["descricao", "descrição", "historico", "histórico", "lancamento", "lançamento", "detalhe"],
        "categoria": ["categoria", "grupo", "classificacao", "classificação"],
        "valor": ["valor", "valor r$", "valor rs", "amount", "preco", "preço"],
        "tipo": ["tipo", "natureza", "entrada saida", "receita despesa"],
        "cartao": ["cartao", "cartão", "forma de pagamento", "pagamento", "conta", "banco"],
        "data": ["data", "dt", "dia"],
        "descricao": ["descricao", "descrição", "historico", "histórico", "nome"],
        "categoria": ["categoria", "grupo"],
        "valor": ["valor", "valor r$", "valor rs"],
        "tipo": ["tipo", "natureza"],
        "cartao": ["forma de pagamento", "pagamento", "cartao", "cartão", "conta"],
    }

    colunas_normais = {normalizar_coluna(coluna): coluna for coluna in df_planilha.columns}
    mapa_colunas = {}

    colunas = {normalizar_coluna(coluna): coluna for coluna in dados.columns}
    mapa = {}
    for destino, opcoes in aliases.items():
        for opcao in opcoes:
            chave = normalizar_coluna(opcao)
            if chave in colunas_normais:
                mapa_colunas[destino] = colunas_normais[chave]
            if normalizar_coluna(opcao) in colunas:
                mapa[destino] = colunas[normalizar_coluna(opcao)]
                break
    if "valor" not in mapa:
        raise ValueError("A planilha precisa ter uma coluna de valor.")

    if "valor" not in mapa_colunas:
        raise ValueError(
            "Não encontrei uma coluna de valor nem abas mensais no formato da Organização Financeira."
        )

    linhas = []
    ignoradas = 0

    for _, linha in df_planilha.iterrows():
        valor_original = converter_valor(linha.get(mapa_colunas["valor"]))
    for _, row in dados.iterrows():
        valor_original = converter_valor(row.get(mapa["valor"]))
        if valor_original is None or valor_original == 0:
            ignoradas += 1
            continue

        data_original = linha.get(mapa_colunas.get("data"), date.today())
        data_convertida = pd.to_datetime(data_original, errors="coerce", dayfirst=True)
        data_convertida = pd.to_datetime(row.get(mapa.get("data"), date.today()), errors="coerce", dayfirst=True)
        if pd.isna(data_convertida):
            data_convertida = pd.Timestamp(date.today())

        tipo_original = str(linha.get(mapa_colunas.get("tipo"), "")).strip().lower()
        if any(palavra in tipo_original for palavra in ["saida", "saída", "despesa", "debito", "débito", "gasto"]):
            tipo = "Saída"
        elif any(palavra in tipo_original for palavra in ["entrada", "receita", "credito", "crédito", "salario", "salário"]):
        tipo_texto = normalizar_coluna(row.get(mapa.get("tipo"), ""))
        if any(palavra in tipo_texto for palavra in PALAVRAS_ENTRADA_PLANILHA):
            tipo = "Entrada"
        elif any(palavra in tipo_texto for palavra in PALAVRAS_SAIDA_PLANILHA):
            tipo = "Saída"
        else:
            tipo = "Entrada" if valor_original > 0 else "Saída"

        valor_final = abs(valor_original) if tipo == "Entrada" else -abs(valor_original)
        descricao = str(linha.get(mapa_colunas.get("descricao"), "")).strip()
        categoria = str(linha.get(mapa_colunas.get("categoria"), "")).strip()
        cartao = str(linha.get(mapa_colunas.get("cartao"), "")).strip()

        linhas.append(
            {
                "data": data_convertida.date().isoformat(),
                "descricao": descricao if descricao and descricao.lower() != "nan" else "Importado da planilha",
                "categoria": categoria if categoria and categoria.lower() != "nan" else "Importado",
                "valor": valor_final,
                "tipo": tipo,
                "cartao": cartao if cartao and cartao.lower() != "nan" else "Planilha",
            }
        adicionar_movimentacao_mensal(
            linhas,
            data_convertida.date().isoformat(),
            texto_planilha(row.get(mapa.get("descricao")), "Importado da planilha"),
            texto_planilha(row.get(mapa.get("categoria")), "Importado"),
            valor_original,
            tipo,
            texto_planilha(row.get(mapa.get("cartao")), "Planilha"),
        )

    return pd.DataFrame(linhas), ignoradas


def chave_movimentacao(registro):
    data_movimentacao = pd.to_datetime(registro.get("data"), errors="coerce", dayfirst=True)
    data_normalizada = (
        data_movimentacao.date().isoformat()
        if not pd.isna(data_movimentacao)
        else texto_planilha(registro.get("data"))
    )
    descricao = normalizar_coluna(texto_planilha(registro.get("descricao")))
    tipo = normalizar_coluna(texto_planilha(registro.get("tipo")))
    valor = converter_valor(registro.get("valor"))
    valor_normalizado = round(float(valor or 0), 2)
    return data_normalizada, descricao, valor_normalizado, tipo
def normalizar_dataframe_financeiro(df):
    return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()


def conciliar_movimentacoes(df_importado, df_existente):
    if df_importado.empty:
        return df_importado.copy(), 0
def normalizar_dataframe_dividas(df):
    return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    chaves_existentes = {
        chave_movimentacao(registro)
        for registro in df_existente.to_dict(orient="records")
    }
    chaves_novas = set()
    indices_novos = []
    duplicadas = 0

    for indice, registro in df_importado.iterrows():
        chave = chave_movimentacao(registro)
        if chave in chaves_existentes or chave in chaves_novas:
            duplicadas += 1
            continue
        chaves_novas.add(chave)
        indices_novos.append(indice)
def normalizar_dataframe_metas(df):
    return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    return df_importado.loc[indices_novos].reset_index(drop=True), duplicadas


def importar_movimentacoes(df_importado):
    df_novo, duplicadas = conciliar_movimentacoes(df_importado, carregar_dados())
    if df_novo.empty:
        return 0, duplicadas

    conn = sqlite3.connect(DB_FILE)
    conn.executemany(
        """
        INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        df_novo[["data", "descricao", "categoria", "valor", "tipo", "cartao"]].itertuples(
            index=False,
            name=None,
        ),
    )
    conn.commit()
    conn.close()
    return len(df_novo), duplicadas


def _pdf_texto(texto):
    return (
        str(texto)
        .replace("\\", "\\\\")
        .replace("(", "\\(")
        .replace(")", "\\)")
        .encode("latin-1", errors="replace")
        .decode("latin-1")
    )


def gerar_relatorio_pdf(df_transacoes, df_investimentos, df_dividas=None, df_metas=None):
    if df_dividas is None:
        df_dividas = pd.DataFrame()
    if df_metas is None:
        df_metas = pd.DataFrame()

    linhas = []

    def adicionar(texto="", estilo="corpo", largura=92):
        partes = textwrap.wrap(str(texto), width=largura) or [""]
        for parte in partes:
            linhas.append((parte, estilo))

    entradas = df_transacoes[df_transacoes["valor"] > 0]["valor"].sum() if len(df_transacoes) else 0
    saidas = abs(df_transacoes[df_transacoes["valor"] < 0]["valor"].sum()) if len(df_transacoes) else 0
    saldo = df_transacoes["valor"].sum() if len(df_transacoes) else 0
    investimentos = df_investimentos["valor"].sum() if len(df_investimentos) else 0
    total_dividas = (
        df_dividas["saldo_negociado"].where(df_dividas["saldo_negociado"] > 0, df_dividas["saldo_original"]).sum()
        if len(df_dividas)
        else 0
    )
    total_metas = df_metas["valor_meta"].sum() if len(df_metas) else 0

    adicionar("Dashboard Financeiro - Relatório Financeiro Detalhado", "titulo", 62)
    adicionar(f"Emitido em {date.today().strftime('%d/%m/%Y')}", "pequeno")
    adicionar()
    adicionar("Resumo financeiro", "secao")
    adicionar(f"Entradas totais: {brl(entradas)}")
    adicionar(f"Saídas totais: {brl(saidas)}")
    adicionar(f"Saldo atual: {brl(saldo)}")
    adicionar(f"Patrimônio investido: {brl(investimentos)}")
    adicionar(f"Dívidas monitoradas: {brl(total_dividas)}")
    adicionar(f"Metas planejadas: {brl(total_metas)}")
    adicionar(f"Movimentações registradas: {len(df_transacoes)}")
    adicionar()
    adicionar("Movimentações", "secao")

    if len(df_transacoes):
        for _, row in df_transacoes.iterrows():
            adicionar(
                f"{data_br(row['data'])} | {row['descricao'] or 'Sem descrição'} | "
                f"{row['categoria'] or 'Sem categoria'} | {brl(row['valor'])}",
                "corpo",
            )
            adicionar(
                f"Tipo: {row['tipo'] or 'Não informado'} | "
                f"Forma de pagamento: {row['cartao'] or 'Não informada'}",
                "pequeno",
            )
            adicionar("-" * 92, "pequeno")
    else:
        adicionar("Nenhuma movimentação cadastrada.")

    adicionar()
    adicionar("Investimentos", "secao")

    if len(df_investimentos):
        for _, investimento in df_investimentos.iterrows():
            adicionar(
                f"{data_br(investimento['data'])} | "
                f"{investimento['descricao'] or 'Sem descrição'} | "
                f"{investimento['tipo'] or 'Sem categoria'} | "
                f"{brl(investimento['valor'])}",
                "corpo",
            )
            adicionar(
                f"Status: {investimento['status'] or 'Não informado'} | "
                f"Rentabilidade: {investimento['rentabilidade'] or 'Não informada'}",
                "pequeno",
            )
            adicionar("-" * 92, "pequeno")
    else:
        adicionar("Nenhum investimento cadastrado.")

    adicionar()
    adicionar("Dívidas e negociações", "secao")

    if len(df_dividas):
        for _, divida in df_dividas.iterrows():
            saldo_base = divida["saldo_negociado"] if divida["saldo_negociado"] > 0 else divida["saldo_original"]
            adicionar(
                f"{data_br(divida['data'])} | {divida['credor'] or 'Credor não informado'} | "
                f"{divida['tipo'] or 'Tipo não informado'} | {brl(saldo_base)}",
                "corpo",
            )
            adicionar(
                f"Status: {divida['status'] or 'Não informado'} | "
                f"Prioridade: {divida['prioridade'] or 'Não informada'} | "
                f"Próxima ação: {divida['proxima_acao'] or 'Não informada'}",
                "pequeno",
            )
            if divida.get("anotacoes"):
                adicionar(f"Anotações: {divida['anotacoes']}", "pequeno")
            adicionar("-" * 92, "pequeno")
    else:
        adicionar("Nenhuma dívida cadastrada.")

    adicionar()
    adicionar("Metas", "secao")

    if len(df_metas):
        for _, meta in df_metas.iterrows():
            valor_meta = max(float(meta["valor_meta"]), 0)
            valor_atual = max(float(meta["valor_atual"]), 0)
            aporte_mensal = max(float(meta["aporte_mensal"]), 0)
            falta = max(valor_meta - valor_atual, 0)
            meses = math.ceil(falta / aporte_mensal) if falta > 0 and aporte_mensal > 0 else 0
            previsao = (
                f"Chega em {texto_meses(meses)}"
                if falta > 0 and aporte_mensal > 0
                else "Meta concluída" if falta <= 0 else "Sem aporte mensal definido"
            )
            adicionar(
                f"{data_br(meta['data'])} | {meta['nome'] or 'Meta sem nome'} | "
                f"Objetivo: {brl(valor_meta)} | Atual: {brl(valor_atual)}",
                "corpo",
            )
            adicionar(
                f"Status: {meta['status'] or 'Não informado'} | Prazo: {meta['prazo'] or 'Não informado'} | "
                f"{previsao}",
                "pequeno",
            )
            if meta.get("anotacoes"):
                adicionar(f"Anotações: {meta['anotacoes']}", "pequeno")
            adicionar("-" * 92, "pequeno")
    else:
        adicionar("Nenhuma meta cadastrada.")

    paginas = []
    pagina_atual = []
    altura_usada = 0
    alturas = {"titulo": 28, "secao": 22, "corpo": 16, "pequeno": 13}

    for linha in linhas:
        altura = alturas[linha[1]]
        if altura_usada + altura > 720:
            paginas.append(pagina_atual)
            pagina_atual = []
            altura_usada = 0
        pagina_atual.append(linha)
        altura_usada += altura

    if pagina_atual:
        paginas.append(pagina_atual)

    objetos = {
        1: b"<< /Type /Catalog /Pages 2 0 R >>",
        3: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
        4: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>",
    }
    referencias_paginas = []

    for indice, pagina in enumerate(paginas):
        numero_pagina = 5 + indice * 2
        numero_conteudo = numero_pagina + 1
        referencias_paginas.append(f"{numero_pagina} 0 R")
        comandos = []
        y = 795

        for texto, estilo in pagina:
            fonte = "F2" if estilo in {"titulo", "secao"} else "F1"
            tamanho = {"titulo": 17, "secao": 13, "corpo": 10, "pequeno": 8}[estilo]
            comandos.append(
                f"BT /{fonte} {tamanho} Tf 50 {y} Td ({_pdf_texto(texto)}) Tj ET"
            )
            y -= alturas[estilo]

        comandos.append(
            f"BT /F1 8 Tf 50 30 Td (Pagina {indice + 1} de {len(paginas)}) Tj ET"
        )
        fluxo = "\n".join(comandos).encode("latin-1")
        objetos[numero_pagina] = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
            f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
            f"/Contents {numero_conteudo} 0 R >>"
        ).encode("latin-1")
        objetos[numero_conteudo] = (
            f"<< /Length {len(fluxo)} >>\nstream\n".encode("latin-1")
            + fluxo
            + b"\nendstream"
        )

    objetos[2] = (
        f"<< /Type /Pages /Kids [{' '.join(referencias_paginas)}] "
        f"/Count {len(paginas)} >>"
    ).encode("latin-1")

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
    pdf.extend(
        f"trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\n"
        f"startxref\n{xref}\n%%EOF".encode("latin-1")
    )
    return bytes(pdf)


def init_db():
    conn = sqlite3.connect(DB_FILE)
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS transacoes (
            id INTEGER PRIMARY KEY,
            data DATE,
            descricao TEXT,
            categoria TEXT,
            valor REAL,
            tipo TEXT,
            cartao TEXT
        );

        CREATE TABLE IF NOT EXISTS investimentos (
            id INTEGER PRIMARY KEY,
            data DATE,
            tipo TEXT,
            valor REAL,
            rentabilidade TEXT,
            descricao TEXT,
            status TEXT
        );

        CREATE TABLE IF NOT EXISTS dividas (
            id INTEGER PRIMARY KEY,
            data DATE,
            credor TEXT,
            tipo TEXT,
            saldo_original REAL,
            desconto REAL,
            saldo_negociado REAL,
            parcela_possivel REAL,
            vencimento DATE,
            prioridade TEXT,
            consequencia TEXT,
            status TEXT,
            proxima_acao TEXT,
            anotacoes TEXT
        );

        CREATE TABLE IF NOT EXISTS metas (
            id INTEGER PRIMARY KEY,
            data DATE,
            nome TEXT,
            valor_meta REAL,
            valor_atual REAL,
            aporte_mensal REAL,
            prazo TEXT,
            status TEXT,
            anotacoes TEXT
        );
        """
    )
    conn.commit()
    conn.close()


def normalizar_dataframe_financeiro(df_dados, colunas_texto):
    df_dados = df_dados.copy()
    if "valor" in df_dados.columns:
        df_dados["valor"] = pd.to_numeric(df_dados["valor"], errors="coerce").fillna(0.0)
    if "data" in df_dados.columns:
        df_dados["data"] = df_dados["data"].fillna("").astype(str)
    for coluna in colunas_texto:
        if coluna in df_dados.columns:
            df_dados[coluna] = df_dados[coluna].fillna("").astype(str)
    return df_dados


def normalizar_dataframe_dividas(df_dados):
    df_dados = df_dados.copy()
    for coluna in ["saldo_original", "desconto", "saldo_negociado", "parcela_possivel"]:
        if coluna in df_dados.columns:
            df_dados[coluna] = pd.to_numeric(df_dados[coluna], errors="coerce").fillna(0.0)
    for coluna in ["data", "vencimento"]:
        if coluna in df_dados.columns:
            df_dados[coluna] = df_dados[coluna].fillna("").astype(str)
    for coluna in [
        "credor",
        "tipo",
        "prioridade",
        "consequencia",
        "status",
        "proxima_acao",
        "anotacoes",
    ]:
        if coluna in df_dados.columns:
            df_dados[coluna] = df_dados[coluna].fillna("").astype(str)
    return df_dados


def normalizar_dataframe_metas(df_dados):
    df_dados = df_dados.copy()
    for coluna in ["valor_meta", "valor_atual", "aporte_mensal"]:
        if coluna in df_dados.columns:
            df_dados[coluna] = pd.to_numeric(df_dados[coluna], errors="coerce").fillna(0.0)
    if "data" in df_dados.columns:
        df_dados["data"] = df_dados["data"].fillna("").astype(str)
    for coluna in ["nome", "prazo", "status", "anotacoes"]:
        if coluna in df_dados.columns:
            df_dados[coluna] = df_dados[coluna].fillna("").astype(str)
    return df_dados


def carregar_dados():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query(
        "SELECT * FROM transacoes ORDER BY data DESC, id DESC",
        conn,
    )
    df = pd.read_sql_query("SELECT * FROM transacoes ORDER BY data DESC, id DESC", conn)
    conn.close()
    return normalizar_dataframe_financeiro(
        df,
        ["descricao", "categoria", "tipo", "cartao"],
    )
    if df.empty:
        return pd.DataFrame(columns=["id", "data", "descricao", "categoria", "valor", "tipo", "cartao"])
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)
    for coluna in ["data", "descricao", "categoria", "tipo", "cartao"]:
        df[coluna] = df[coluna].fillna("").astype(str)
    return df


def carregar_investimentos():
    conn = sqlite3.connect(DB_FILE)
    df_investimentos = pd.read_sql_query(
        "SELECT * FROM investimentos ORDER BY data DESC, id DESC",
        conn,
    )
    df = pd.read_sql_query("SELECT * FROM investimentos ORDER BY data DESC, id DESC", conn)
    conn.close()
    return normalizar_dataframe_financeiro(
        df_investimentos,
        ["tipo", "rentabilidade", "descricao", "status"],
    )
    if df.empty:
        return pd.DataFrame(columns=["id", "data", "tipo", "valor", "rentabilidade", "descricao", "status"])
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)
    for coluna in ["data", "tipo", "rentabilidade", "descricao", "status"]:
        df[coluna] = df[coluna].fillna("").astype(str)
    return df


def carregar_dividas():
    conn = sqlite3.connect(DB_FILE)
    df_dividas = pd.read_sql_query(
        "SELECT * FROM dividas ORDER BY data DESC, id DESC",
        conn,
    )
    df = pd.read_sql_query("SELECT * FROM dividas ORDER BY data DESC, id DESC", conn)
    conn.close()
    return normalizar_dataframe_dividas(df_dividas)
    if df.empty:
        return pd.DataFrame(
            columns=[
                "id",
                "data",
                "credor",
                "tipo",
                "saldo_original",
                "desconto",
                "saldo_negociado",
                "parcela_possivel",
                "vencimento",
                "prioridade",
                "consequencia",
                "status",
                "proxima_acao",
                "anotacoes",
            ]
        )
    for coluna in ["saldo_original", "desconto", "saldo_negociado", "parcela_possivel"]:
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0.0)
    for coluna in ["data", "credor", "tipo", "vencimento", "prioridade", "consequencia", "status", "proxima_acao", "anotacoes"]:
        df[coluna] = df[coluna].fillna("").astype(str)
    return df


def carregar_metas():
    conn = sqlite3.connect(DB_FILE)
    df_metas = pd.read_sql_query(
        "SELECT * FROM metas ORDER BY data DESC, id DESC",
        conn,
    df = pd.read_sql_query("SELECT * FROM metas ORDER BY data DESC, id DESC", conn)
    conn.close()
    if df.empty:
        return pd.DataFrame(columns=["id", "data", "nome", "valor_meta", "valor_atual", "aporte_mensal", "prazo", "status", "anotacoes"])
    for coluna in ["valor_meta", "valor_atual", "aporte_mensal"]:
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0.0)
    for coluna in ["data", "nome", "prazo", "status", "anotacoes"]:
        df[coluna] = df[coluna].fillna("").astype(str)
    return df


def salvar_transacao(data_mov, descricao, categoria, valor, tipo, cartao):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        "INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao) VALUES (?, ?, ?, ?, ?, ?)",
        (str(data_mov), descricao, categoria, valor, tipo, cartao),
    )
    conn.commit()
    conn.close()
    return normalizar_dataframe_metas(df_metas)


def excluir_transacao(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def limpar_historico_movimentacoes():
def excluir_investimento(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes")
    conn.execute("DELETE FROM investimentos WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def excluir_investimento(iid):
def excluir_divida(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM investimentos WHERE id = ?", (iid,))
    conn.execute("DELETE FROM dividas WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def excluir_divida(did):
def excluir_meta(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM dividas WHERE id = ?", (did,))
    conn.execute("DELETE FROM metas WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def excluir_meta(mid):
def limpar_historico():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM metas WHERE id = ?", (mid,))
    conn.execute("DELETE FROM transacoes")
    conn.commit()
    conn.close()


def salvar_transacao(data_movimentacao, descricao, categoria, valor, tipo, cartao):
def limpar_historico_movimentacoes():
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        """
        INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (data_movimentacao, descricao, categoria, valor, tipo, cartao),
    )
    conn.execute("DELETE FROM transacoes")
    conn.commit()
    conn.close()


def salvar_investimento(data_investimento, tipo, valor, rentabilidade, descricao, status):
def salvar_investimento(data_inv, tipo, valor, rentabilidade, descricao, status):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        """
        INSERT INTO investimentos
            (data, tipo, valor, rentabilidade, descricao, status)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (data_investimento, tipo, valor, rentabilidade, descricao, status),
        "INSERT INTO investimentos (data, tipo, valor, rentabilidade, descricao, status) VALUES (?, ?, ?, ?, ?, ?)",
        (str(data_inv), tipo, valor, rentabilidade, descricao, status),
    )
    conn.commit()
    conn.close()


def salvar_divida(
    data_divida,
    credor,
    tipo,
    saldo_original,
    desconto,
    saldo_negociado,
    parcela_possivel,
    vencimento,
    prioridade,
    consequencia,
    status,
    proxima_acao,
    anotacoes,
):
def salvar_divida(data_divida, credor, tipo, saldo_original, desconto, saldo_negociado, parcela_possivel, vencimento, prioridade, consequencia, status, proxima_acao, anotacoes):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        """
        INSERT INTO dividas
            (
                data, credor, tipo, saldo_original, desconto, saldo_negociado,
                parcela_possivel, vencimento, prioridade, consequencia,
                status, proxima_acao, anotacoes
            )
            (data, credor, tipo, saldo_original, desconto, saldo_negociado, parcela_possivel, vencimento, prioridade, consequencia, status, proxima_acao, anotacoes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            data_divida,
            str(data_divida),
            credor,
            tipo,
            saldo_original,
            desconto,
            saldo_negociado,
            parcela_possivel,
            vencimento,
            str(vencimento),
            prioridade,
            consequencia,
            status,
            proxima_acao,
            anotacoes,
        ),
    )
    conn.commit()
    conn.close()


def salvar_meta(data_meta, nome, valor_meta, valor_atual, aporte_mensal, prazo, status, anotacoes):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        """
        INSERT INTO metas
            (data, nome, valor_meta, valor_atual, aporte_mensal, prazo, status, anotacoes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            data_meta,
            nome,
            valor_meta,
            valor_atual,
            aporte_mensal,
            prazo,
            status,
            anotacoes,
        ),
        (str(data_meta), nome, valor_meta, valor_atual, aporte_mensal, prazo, status, anotacoes),
    )
    conn.commit()
    conn.close()


def style_plot(fig):
    fig.update_layout(
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(color="#111318", size=12),
        title=dict(font=dict(size=20, color="#111318"), x=0.04, xanchor="left"),
        legend=dict(
            bgcolor="rgba(255,255,255,0)",
            font=dict(color="#111318"),
            orientation="h",
            yanchor="bottom",
            y=1.01,
            xanchor="right",
            x=1,
        ),
        hoverlabel=dict(bgcolor="#111318", font_color="#ffffff"),
        margin=dict(l=34, r=26, t=80, b=44),
        height=410,
        separators=",.",
def chave_movimentacao(registro):
    data_mov = pd.to_datetime(registro.get("data"), errors="coerce", dayfirst=True)
    data_norm = data_mov.date().isoformat() if not pd.isna(data_mov) else limpar_texto(registro.get("data"))
    return (
        data_norm,
        normalizar_coluna(registro.get("descricao")),
        round(float(registro.get("valor") or 0), 2),
        normalizar_coluna(registro.get("tipo")),
    )
    fig.update_xaxes(
        gridcolor="rgba(17,19,24,0.08)",
        linecolor="rgba(17,19,24,0.10)",
        zerolinecolor="rgba(17,19,24,0.10)",
        title_font=dict(color="#747985"),
        tickfont=dict(color="#747985"),
    )
    fig.update_yaxes(
        gridcolor="rgba(17,19,24,0.08)",
        linecolor="rgba(17,19,24,0.10)",
        zerolinecolor="rgba(17,19,24,0.10)",
        title_font=dict(color="#747985"),
        tickfont=dict(color="#747985"),
    )
    return fig


# Acesso direto: o usuário entra no dashboard sem etapa de login.
init_db()
try:
    df = carregar_dados()
    df_investimentos = carregar_investimentos()
    df_dividas = carregar_dividas()
    df_metas = carregar_metas()
except Exception as erro:
    st.error(f"Não foi possível carregar seus dados: {mensagem_erro_usuario(erro)}")
    st.info("Atualize a página e tente novamente.")
    st.stop()
def conciliar_movimentacoes(df_importado, df_existente):
    if df_importado.empty:
        return df_importado.copy(), 0
    existentes = {chave_movimentacao(registro) for registro in df_existente.to_dict(orient="records")}
    vistos = set()
    indices = []
    duplicadas = 0
    for indice, registro in df_importado.iterrows():
        chave = chave_movimentacao(registro)
        if chave in existentes or chave in vistos:
            duplicadas += 1
            continue
        vistos.add(chave)
        indices.append(indice)
    return df_importado.loc[indices].reset_index(drop=True), duplicadas

hero_total_entradas = df[df["valor"] > 0]["valor"].sum() if len(df) > 0 else 0
hero_total_saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) > 0 else 0
hero_saldo = df["valor"].sum() if len(df) > 0 else 0
hero_total_investido = df_investimentos["valor"].sum() if len(df_investimentos) > 0 else 0
hero_total_metas = df_metas["valor_meta"].sum() if len(df_metas) > 0 else 0
hero_taxa_sobra = (hero_saldo / hero_total_entradas) * 100 if hero_total_entradas > 0 else 0
hero_comprometimento = (hero_total_saidas / hero_total_entradas) * 100 if hero_total_entradas > 0 else 0
if len(df_dividas):
    hero_dividas_view = df_dividas.copy()
    hero_dividas_view["saldo_base"] = hero_dividas_view["saldo_negociado"].where(
        hero_dividas_view["saldo_negociado"] > 0,
        hero_dividas_view["saldo_original"],

def importar_movimentacoes(df_importado):
    df_novo, duplicadas = conciliar_movimentacoes(df_importado, carregar_dados())
    if df_novo.empty:
        return 0, duplicadas
    conn = sqlite3.connect(DB_FILE)
    conn.executemany(
        "INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao) VALUES (?, ?, ?, ?, ?, ?)",
        df_novo[["data", "descricao", "categoria", "valor", "tipo", "cartao"]].itertuples(index=False, name=None),
    )
    hero_dividas_abertas = hero_dividas_view[hero_dividas_view["status"] != "Quitada"]
    hero_total_dividas_abertas = hero_dividas_abertas["saldo_base"].sum()
    hero_parcelas_dividas = hero_dividas_abertas["parcela_possivel"].sum()
else:
    hero_total_dividas_abertas = 0
    hero_parcelas_dividas = 0
if len(df_metas) and hero_total_metas > 0:
    hero_progresso_metas = limitar_percentual((df_metas["valor_atual"].sum() / hero_total_metas) * 100)
else:
    hero_progresso_metas = 0
hero_score = calcular_score_financeiro(
    hero_total_entradas,
    hero_total_saidas,
    hero_saldo,
    hero_total_investido,
    hero_total_dividas_abertas,
    hero_parcelas_dividas,
    hero_progresso_metas,
)
    conn.commit()
    conn.close()
    return len(df_novo), duplicadas

# ====================== HERO ======================
st.markdown(
    f"""
    <section class="dashboard-shell">
        <aside class="side-rail">
            <div class="side-logo"><span>DF</span><strong>Dashboard Financeiro</strong></div>
            <div class="side-menu">
                <span class="active">Dashboard <b>{len(df)}</b></span>
                <span>Metas <b>{len(df_metas)}</b></span>
                <span>Dívidas <b>{len(df_dividas)}</b></span>
                <span>Histórico <b>{len(df)}</b></span>
                <span>Relatório <b>PDF</b></span>
            </div>
            <div class="upgrade-card">
                <div class="upgrade-icon">↗</div>
                <div class="upgrade-title">Plano financeiro</div>
                <div class="upgrade-copy">Veja o que entra, o que sai e qual próxima decisão melhora seu mês.</div>
                <div class="upgrade-button">Organizar agora</div>
            </div>
        </aside>

        <section class="dashboard-stage">
            <div class="stage-top">
                <div class="stage-title">
                    <span class="eyebrow">Painel inteligente</span>
                    <h1>Dashboard<br>Financeiro</h1>
                    <p>Controle entradas, gastos, dívidas, metas e investimentos em uma visão clara para decidir melhor todos os meses.</p>
                </div>
                <div class="stage-actions">
                    <div class="round-action">⌕</div>
                    <div class="round-action">•</div>
                    <div class="user-chip"><span class="avatar-dot">DF</span> Visão premium</div>
                </div>
            </div>
def ler_planilha(arquivo):
    nome = arquivo.name.lower()
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

            <div class="metric-grid hero-metrics">
                <div class="metric-card accent-card">
                    <div class="metric-label">Saldo atual</div>
                    <div class="metric-value">{brl(hero_saldo)}</div>
                    <div class="metric-foot">Resultado de tudo que foi registrado</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Entradas</div>
                    <div class="metric-value">{brl(hero_total_entradas)}</div>
                    <div class="metric-foot">Receitas manuais e importadas</div>
                </div>
                <div class="metric-card blue-card">
                    <div class="metric-label">Saúde financeira</div>
                    <div class="metric-value">{hero_score}/100</div>
                    <div class="metric-foot">{pct(hero_taxa_sobra)} de sobra acumulada</div>
                </div>
                <div class="metric-card dark-card">
                    <div class="metric-label">Patrimônio</div>
                    <div class="metric-value">{brl(hero_total_investido)}</div>
                    <div class="metric-foot">Investimentos registrados</div>
                </div>
            </div>

            <div class="finance-quote">
                <div>
                    <strong>“Preço é o que você paga; valor é o que você recebe.”</strong>
                    <span>Benjamin Graham · Use seus números para proteger valor, não só pagar contas.</span>
                </div>
                <b>{pct(hero_comprometimento)}</b>
            </div>
        </section>
    </section>
    """,
    unsafe_allow_html=True,
)
MESES = {
    "janeiro": 1,
    "fevereiro": 2,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}

# ====================== NAVEGAÇÃO ======================
aba = st.tabs(["➕ Nova Movimentação", "📊 Dashboard", "🎯 Metas", "🤝 Dívidas", "📋 Histórico"])

# ====================== ABA 1 ======================
with aba[0]:
    st.subheader("Adicionar Nova Movimentação")
def preparar_planilha_mensal(planilhas):
    linhas = []
    for nome_aba, df_aba in planilhas.items():
        mes = MESES.get(normalizar(nome_aba))
        if not mes:
            continue
        for linha_idx in range(len(df_aba.index)):
            for col_idx in range(len(df_aba.columns) - 1):
                if normalizar_coluna(df_aba.iat[linha_idx, col_idx]) != "entradas":
                    continue
                ultima_descricao = ""
                for idx in range(linha_idx + 1, min(len(df_aba.index), linha_idx + 14)):
                    descricao = limpar_texto(df_aba.iat[idx, col_idx])
                    valor_original = converter_valor(df_aba.iat[idx, col_idx + 1])
                    desc_norm = normalizar_coluna(descricao)
                    if desc_norm.startswith("total") or desc_norm == "saidas":
                        break
                    if descricao:
                        ultima_descricao = descricao
                    if valor_original is None or valor_original <= 0:
                        continue
                    linhas.append(
                        {
                            "data": date(date.today().year, mes, 1).isoformat(),
                            "descricao": descricao or ultima_descricao or "Entrada da planilha",
                            "categoria": "Receita",
                            "valor": abs(valor_original),
                            "tipo": "Entrada",
                            "cartao": "Planilha",
                        }
                    )
        for linha_idx in range(len(df_aba.index)):
            valores_linha = [normalizar(df_aba.iat[linha_idx, col]) for col in range(len(df_aba.columns))]
            if not any(item in valores_linha for item in ["descricao", "nome"]):
                continue
            if "valor" not in valores_linha:
                continue

    tipo = st.radio("Tipo", ["💰 Entrada", "💸 Saída"], horizontal=True)
    tipo_limpo = "Entrada" if "Entrada" in tipo else "Saída"
            col_descricao = next((i for i, item in enumerate(valores_linha) if item in ["descricao", "nome"]), None)
            col_valor = next((i for i, item in enumerate(valores_linha) if item == "valor" and i > col_descricao), None)
            col_tipo = next((i for i, item in enumerate(valores_linha) if item in ["tipo", "pagamento"] and i > col_descricao), None)
            col_categoria = next((i for i, item in enumerate(valores_linha) if item == "categoria" and i > col_descricao), None)
            if col_descricao is None or col_valor is None:
                continue

    with st.form("nova"):
        col1, col2 = st.columns(2)
            contexto = " ".join(
                normalizar(df_aba.iat[i, j])
                for i in range(max(0, linha_idx - 4), linha_idx + 1)
                for j in range(max(0, col_descricao - 2), min(len(df_aba.columns), col_valor + 3))
            )

        with col1:
            data = st.date_input("Data", value=date.today(), format="DD/MM/YYYY")
            descricao = st.text_input("Descrição")
            for idx in range(linha_idx + 1, len(df_aba.index)):
                descricao = limpar_texto(df_aba.iat[idx, col_descricao])
                valor_original = converter_valor(df_aba.iat[idx, col_valor])
                desc_norm = normalizar(descricao)
                if not descricao and not valor_original:
                    break
                if not descricao:
                    continue
                if desc_norm.startswith("total") or desc_norm in ["saidas", "investimentos", "reserva"]:
                    break
                if valor_original is None or valor_original == 0:
                    continue

            if tipo_limpo == "Entrada":
                categoria = st.selectbox(
                    "Categoria",
                    ["Salário", "Rendas Extras", "Freelance", "Reembolso"],
                tipo_original = limpar_texto(df_aba.iat[idx, col_tipo], "Debito") if col_tipo is not None else "Debito"
                categoria = limpar_texto(df_aba.iat[idx, col_categoria], "Receita" if "entrada" in contexto else "Outros") if col_categoria is not None else ("Receita" if "entrada" in contexto else "Outros")
                texto_tipo = normalizar(" ".join([tipo_original, descricao, categoria, contexto]))
                texto_direto = normalizar(" ".join([tipo_original, descricao, categoria]))
                if "saida" in texto_tipo or any(p in texto_direto for p in ["debito", "credito", "despesa", "gasto"]):
                    tipo = "Saida"
                elif any(p in texto_tipo for p in ["entrada", "receita", "salario", "renda", "freelance", "reembolso"]):
                    tipo = "Entrada"
                else:
                    tipo = "Saida"
                valor = abs(valor_original) if tipo == "Entrada" else -abs(valor_original)
                ano = date.today().year
                linhas.append(
                    {
                        "data": date(ano, mes, 1).isoformat(),
                        "descricao": descricao or "Importado da planilha",
                        "categoria": categoria,
                        "valor": valor,
                        "tipo": tipo,
                        "cartao": tipo_original,
                    }
                )
            else:
                categoria = st.selectbox(
                    "Categoria",
                    [
                        "Mercado",
                        "Aluguel",
                        "Contas",
                        "Lazer",
                        "Roupa",
                        "Beleza",
                        "Transporte",
                        "Outro",
                    ],
                )
    return pd.DataFrame(linhas).drop_duplicates() if linhas else pd.DataFrame()

        with col2:
            valor = st.number_input("Valor R$", value=0.0, step=0.01, min_value=0.0)
            cartao = st.text_input("Forma de Pagamento")

        if st.form_submit_button("Salvar Movimentação"):
            if not descricao.strip():
                st.error("Informe uma descrição para a movimentação.")
            elif valor <= 0:
                st.error("Informe um valor maior que zero.")
            else:
                valor_final = valor if tipo_limpo == "Entrada" else -abs(valor)
                try:
                    salvar_transacao(data, descricao.strip(), categoria, valor_final, tipo_limpo, cartao)
                    st.success("Movimentação salva com sucesso!")
                    st.rerun()
                except Exception as erro:
                    st.error(f"Não foi possível salvar a movimentação: {mensagem_erro_usuario(erro)}")
def preparar_importacao(dados):
    if isinstance(dados, dict):
        mensal = preparar_planilha_mensal(dados)
        if not mensal.empty:
            return mensal, 0
        primeira = next(iter(dados.values()), pd.DataFrame())
        if primeira.empty:
            return pd.DataFrame(), 0
        primeira = primeira.copy()
        primeira.columns = primeira.iloc[0]
        dados = primeira.iloc[1:].reset_index(drop=True)

# ====================== ABA 2 ======================
with aba[1]:
    st.subheader("Dashboard em Tempo Real")
    if st.session_state.get("resultado_importacao"):
        st.success(st.session_state.pop("resultado_importacao"))
    aliases = {
        "data": ["data", "dt", "dia"],
        "descricao": ["descricao", "descrição", "historico", "histórico", "nome"],
        "categoria": ["categoria", "grupo"],
        "valor": ["valor", "valor r$", "valor rs"],
        "tipo": ["tipo", "natureza"],
        "cartao": ["forma de pagamento", "pagamento", "cartao", "cartão", "conta"],
    }
    colunas = {normalizar(coluna): coluna for coluna in dados.columns}
    mapa = {}
    for destino, opcoes in aliases.items():
        for opcao in opcoes:
            if normalizar(opcao) in colunas:
                mapa[destino] = colunas[normalizar(opcao)]
                break
    if "valor" not in mapa:
        raise ValueError("A planilha precisa ter uma coluna de valor.")

    total_entradas = df[df["valor"] > 0]["valor"].sum() if len(df) > 0 else 0
    total_saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) > 0 else 0
    saldo = df["valor"].sum() if len(df) > 0 else 0
    linhas = []
    ignoradas = 0
    for _, row in dados.iterrows():
        valor_original = converter_valor(row.get(mapa["valor"]))
        if valor_original is None or valor_original == 0:
            ignoradas += 1
            continue
        data_original = row.get(mapa.get("data"), date.today())
        data_convertida = pd.to_datetime(data_original, errors="coerce", dayfirst=True)
        if pd.isna(data_convertida):
            data_convertida = pd.Timestamp(date.today())
        tipo_original = normalizar(row.get(mapa.get("tipo"), ""))
        if any(p in tipo_original for p in ["saida", "despesa", "debito", "gasto"]):
            tipo = "Saida"
        elif any(p in tipo_original for p in ["entrada", "receita", "credito", "salario"]):
            tipo = "Entrada"
        else:
            tipo = "Entrada" if valor_original > 0 else "Saida"
        linhas.append(
            {
                "data": data_convertida.date().isoformat(),
                "descricao": limpar_texto(row.get(mapa.get("descricao")), "Importado da planilha"),
                "categoria": limpar_texto(row.get(mapa.get("categoria")), "Importado"),
                "valor": abs(valor_original) if tipo == "Entrada" else -abs(valor_original),
                "tipo": tipo,
                "cartao": limpar_texto(row.get(mapa.get("cartao")), "Planilha"),
            }
        )
    return pd.DataFrame(linhas), ignoradas

    st.markdown(
        f"""<div class="metric-grid"><div class="metric-card accent-card"><div class="metric-label">Entradas</div><div class="metric-value">{brl(total_entradas)}</div><div class="metric-foot">Receitas registradas</div></div><div class="metric-card"><div class="metric-label">Saídas</div><div class="metric-value">{brl(total_saidas)}</div><div class="metric-foot">Despesas acumuladas</div></div><div class="metric-card blue-card"><div class="metric-label">Saldo</div><div class="metric-value">{brl(saldo)}</div><div class="metric-foot">Resultado atual</div></div><div class="metric-card dark-card"><div class="metric-label">Registros</div><div class="metric-value">{len(df)}</div><div class="metric-foot">Movimentações salvas</div></div></div>""",
        unsafe_allow_html=True,
    )

    df_mes_recente, nome_mes_recente = resumo_mes_recente(df)
    fluxo_mensal_resumo = preparar_fluxo_mensal(df)
    entradas_mes = df_mes_recente[df_mes_recente["valor"] > 0]["valor"].sum() if len(df_mes_recente) else 0
    saidas_mes = abs(df_mes_recente[df_mes_recente["valor"] < 0]["valor"].sum()) if len(df_mes_recente) else 0
    saldo_mes = entradas_mes - saidas_mes
    media_sobra = fluxo_mensal_resumo["saldo"].mean() if len(fluxo_mensal_resumo) else saldo
    media_entradas = fluxo_mensal_resumo["entradas"].mean() if len(fluxo_mensal_resumo) else total_entradas
    media_saidas = fluxo_mensal_resumo["saidas"].mean() if len(fluxo_mensal_resumo) else total_saidas
def modelo_excel():
    arquivo = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Movimentacoes"
    linhas = [
        ["Data", "Descricao", "Categoria", "Valor", "Tipo", "Forma de Pagamento"],
        [date.today(), "Salario mensal", "Salario", 3500, "Entrada", "Conta corrente"],
        [date.today(), "Mercado", "Mercado", 420, "Saida", "Cartao"],
        [date.today(), "Internet", "Contas", 120, "Saida", "Pix"],
    ]
    for linha in linhas:
        ws.append(linha)
    tabela = Table(displayName="TabelaMovimentacoes", ref="A1:F4")
    tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True)
    ws.add_table(tabela)
    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="111318")
        celula.alignment = Alignment(horizontal="center")
    for col, largura in {"A": 14, "B": 28, "C": 18, "D": 16, "E": 12, "F": 24}.items():
        ws.column_dimensions[col].width = largura
    for celula in ws["A"][1:]:
        celula.number_format = "dd/mm/yyyy"
    for celula in ws["D"][1:]:
        celula.number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'
    wb.save(arquivo)
    arquivo.seek(0)
    return arquivo.getvalue()

    despesas_mes = df_mes_recente[df_mes_recente["valor"] < 0].copy() if len(df_mes_recente) else pd.DataFrame()
    if len(despesas_mes):
        despesas_mes["valor_abs"] = despesas_mes["valor"].abs()
        maior_gasto = despesas_mes.groupby("categoria", as_index=False)["valor_abs"].sum()
        maior_gasto = maior_gasto.sort_values("valor_abs", ascending=False).iloc[0]
        maior_gasto_texto = f"{maior_gasto['categoria']} · {brl(maior_gasto['valor_abs'])}"
        maior_gasto_acao = "Comece revisando essa categoria antes de cortar tudo ao mesmo tempo."
    else:
        maior_gasto_texto = "Sem gastos no mês"
        maior_gasto_acao = "Quando lançar gastos, o app mostra onde atacar primeiro."

    dividas_abertas = pd.DataFrame()
    if len(df_dividas):
        df_dividas_resumo = df_dividas.copy()
        df_dividas_resumo["saldo_base"] = df_dividas_resumo["saldo_negociado"].where(
            df_dividas_resumo["saldo_negociado"] > 0,
            df_dividas_resumo["saldo_original"],
        )
        dividas_abertas = df_dividas_resumo[df_dividas_resumo["status"] != "Quitada"]
def gerar_pdf(df, investimentos, dividas):
    linhas = []

    if saldo_mes >= 0:
        status_mes = "No azul"
        classe_mes = "answer-good"
        acao_mes = f"Em {nome_mes_recente}, sobrou {brl(saldo_mes)}. Separe uma parte antes de gastar."
    else:
        status_mes = "Atenção"
        classe_mes = "answer-risk"
        acao_mes = f"Em {nome_mes_recente}, faltou {brl(abs(saldo_mes))}. Reduza o maior gasto primeiro."
    def add(texto="", largura=92):
        for parte in textwrap.wrap(str(texto), width=largura) or [""]:
            linhas.append(parte)

    if media_sobra > 0:
        classe_sobra = "answer-good"
        acao_sobra = "Esse é o valor médio que pode virar reserva, meta ou quitação de dívida."
    elif media_sobra < 0:
        classe_sobra = "answer-risk"
        acao_sobra = "A média está negativa. Antes de investir, ajuste gastos e dívidas."
    else:
        classe_sobra = "answer-care"
        acao_sobra = "Você está empatando. Um corte pequeno já muda o jogo."
    entradas = df[df["valor"] > 0]["valor"].sum() if len(df) else 0
    saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) else 0
    saldo = df["valor"].sum() if len(df) else 0
    add("Dashboard Financeiro - Relatorio detalhado")
    add(f"Emitido em {date.today().strftime('%d/%m/%Y')}")
    add("")
    add(f"Entradas: {brl(entradas)}")
    add(f"Saidas: {brl(saidas)}")
    add(f"Saldo: {brl(saldo)}")
    add(f"Investimentos cadastrados: {len(investimentos)}")
    add(f"Dividas cadastradas: {len(dividas)}")
    add("")
    add("Movimentacoes")
    for _, row in df.iterrows():
        add(f"{data_br(row['data'])} | {row['descricao']} | {row['categoria']} | {row['tipo']} | {brl(row['valor'])}")

    if len(dividas_abertas):
        proxima_acao_valor = "Negociar dívida"
        proxima_acao_texto = "Priorize a dívida de maior risco ou maior parcela antes de assumir novos gastos."
        proxima_acao_classe = "answer-care"
    elif media_sobra > 0:
        proxima_acao_valor = "Guardar primeiro"
        proxima_acao_texto = "Reserve uma parte assim que receber, antes das despesas do mês."
        proxima_acao_classe = "answer-good"
    else:
        proxima_acao_valor = "Cortar vazamento"
        proxima_acao_texto = "Escolha uma categoria para reduzir este mês e acompanhe no dashboard."
        proxima_acao_classe = "answer-risk"
    paginas = [linhas[i : i + 42] for i in range(0, len(linhas), 42)] or [[]]
    objetos = {
        1: b"<< /Type /Catalog /Pages 2 0 R >>",
        3: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
        4: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>",
    }
    referencias = []
    for idx, pagina in enumerate(paginas):
        page_id = 5 + idx * 2
        content_id = page_id + 1
        referencias.append(f"{page_id} 0 R")
        comandos = []
        y = 795
        for linha in pagina:
            fonte = "F2" if y == 795 else "F1"
            tamanho = 15 if y == 795 else 9
            texto = str(linha).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)").encode("latin-1", errors="replace").decode("latin-1")
            comandos.append(f"BT /{fonte} {tamanho} Tf 50 {y} Td ({texto}) Tj ET")
            y -= 18
        fluxo = "\n".join(comandos).encode("latin-1")
        objetos[page_id] = f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {content_id} 0 R >>".encode("latin-1")
        objetos[content_id] = f"<< /Length {len(fluxo)} >>\nstream\n".encode("latin-1") + fluxo + b"\nendstream"
    objetos[2] = f"<< /Type /Pages /Kids [{' '.join(referencias)}] /Count {len(paginas)} >>".encode("latin-1")
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

    taxa_sobra_dashboard = (saldo / total_entradas) * 100 if total_entradas > 0 else 0
    comprometimento_dashboard = (
        ((total_saidas + hero_parcelas_dividas) / total_entradas) * 100
        if total_entradas > 0
        else 0
    )
    reserva_meses = hero_total_investido / media_saidas if media_saidas > 0 else 0
    reserva_meses_texto = f"{reserva_meses:.1f}".replace(".", ",")
    reserva_percentual = limitar_percentual((reserva_meses / 6) * 100)
    score_dashboard = calcular_score_financeiro(
        total_entradas,
        total_saidas,
        saldo,
        hero_total_investido,
        hero_total_dividas_abertas,
        hero_parcelas_dividas,
        hero_progresso_metas,
    )

    st.markdown(
        f"""<div class="indicator-grid">
            <div class="indicator-card"><div class="indicator-top">Score financeiro <span class="indicator-icon">★</span></div><div class="indicator-value">{score_dashboard}/100</div><div class="indicator-note">Combina saldo, dívidas, reserva e metas.</div><div class="progress-track"><span style="width:{limitar_percentual(score_dashboard)}%"></span></div></div>
            <div class="indicator-card"><div class="indicator-top">Taxa de sobra <span class="indicator-icon">↗</span></div><div class="indicator-value">{pct(taxa_sobra_dashboard)}</div><div class="indicator-note">Parte das entradas que virou saldo.</div><div class="progress-track"><span style="width:{limitar_percentual(taxa_sobra_dashboard)}%"></span></div></div>
            <div class="indicator-card"><div class="indicator-top">Comprometimento <span class="indicator-icon">!</span></div><div class="indicator-value">{pct(comprometimento_dashboard)}</div><div class="indicator-note">Saídas e parcelas comparadas às entradas.</div><div class="progress-track"><span style="width:{limitar_percentual(comprometimento_dashboard)}%"></span></div></div>
            <div class="indicator-card"><div class="indicator-top">Reserva estimada <span class="indicator-icon">◆</span></div><div class="indicator-value">{reserva_meses_texto} meses</div><div class="indicator-note">Investimentos divididos pela média de saídas.</div><div class="progress-track"><span style="width:{reserva_percentual}%"></span></div></div>
        </div>""",
        unsafe_allow_html=True,
def style_plot(fig):
    fig.update_layout(
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(color="#111318", size=12),
        title=dict(font=dict(size=19, color="#111318"), x=0.04),
        margin=dict(l=34, r=26, t=76, b=48),
        height=410,
        legend=dict(orientation="h", y=1.02, x=1, xanchor="right"),
        hoverlabel=dict(bgcolor="#111318", font_color="#ffffff"),
        separators=",.",
    )
    fig.update_xaxes(gridcolor="rgba(17,19,24,0.08)", tickfont=dict(color="#747985"))
    fig.update_yaxes(gridcolor="rgba(17,19,24,0.08)", tickfont=dict(color="#747985"))
    return fig

    st.markdown(
        f"""<div class="chart-intro"><strong>O que preciso saber agora?</strong><br>
        Respostas simples com base nas movimentações cadastradas e importadas.</div>
        <div class="answer-grid">
            <div class="answer-card {classe_mes}"><div class="answer-question">Como está meu mês?</div><div class="answer-value">{status_mes}</div><div class="answer-action">{acao_mes}</div></div>
            <div class="answer-card answer-care"><div class="answer-question">Maior gasto</div><div class="answer-value">{escape(str(maior_gasto_texto))}</div><div class="answer-action">{maior_gasto_acao}</div></div>
            <div class="answer-card {classe_sobra}"><div class="answer-question">Quanto sobra em média?</div><div class="answer-value">{brl(media_sobra)}</div><div class="answer-action">{acao_sobra}</div></div>
            <div class="answer-card {proxima_acao_classe}"><div class="answer-question">Próxima ação</div><div class="answer-value">{proxima_acao_valor}</div><div class="answer-action">{proxima_acao_texto}</div></div>
        </div>""",
        unsafe_allow_html=True,
    )

    with st.expander("🔮 Meu dinheiro futuro", expanded=False):
        st.caption(
            "Simule os próximos 6 meses com ajustes simples. A projeção usa sua média mensal atual."
        )
        col_cenario1, col_cenario2, col_cenario3 = st.columns(3)
        with col_cenario1:
            reduzir_gastos = st.slider(
                "E se eu gastar menos por mês?",
                min_value=0,
                max_value=3000,
                value=0,
                step=50,
                format="R$ %d",
            )
        with col_cenario2:
            renda_extra = st.slider(
                "E se eu receber um extra por mês?",
                min_value=0,
                max_value=5000,
                value=0,
                step=50,
                format="R$ %d",
            )
        with col_cenario3:
            separar_meta = st.slider(
                "Quero separar para metas por mês",
                min_value=0,
                max_value=5000,
                value=0,
                step=50,
                format="R$ %d",
            )
def score_financeiro(entradas, saidas, saldo, investimentos, dividas):
    if entradas <= 0 and saidas <= 0:
        return 50
    sobra = (saldo / entradas) * 100 if entradas > 0 else -20
    comprometimento = (saidas / entradas) * 100 if entradas > 0 else 100
    score = 55 + max(-25, min(sobra, 25)) - max(0, min(comprometimento - 75, 20))
    if investimentos > 0:
        score += 10
    if dividas > 0:
        score -= min((dividas / max(entradas, 1)) * 6, 15)
    return int(max(0, min(round(score), 100)))

        fluxo_normal = media_sobra + reduzir_gastos + renda_extra - separar_meta
        ajuste_bom = max(media_entradas * 0.05, 100) if media_entradas else 100
        ajuste_apertado = max(media_saidas * 0.08, 100) if media_saidas else 100
        inicio_mes = pd.Timestamp(date.today()).to_period("M").to_timestamp()
        saldo_normal = saldo
        saldo_bom = saldo
        saldo_apertado = saldo
        projecoes = []

        for indice_mes in range(1, 7):
            mes = inicio_mes + pd.DateOffset(months=indice_mes)
            saldo_normal += fluxo_normal
            saldo_bom += fluxo_normal + ajuste_bom
            saldo_apertado += fluxo_normal - ajuste_apertado
            projecoes.extend(
                [
                    {"Mês": mes, "Cenário": "Normal", "Saldo projetado": saldo_normal},
                    {"Mês": mes, "Cenário": "Bom", "Saldo projetado": saldo_bom},
                    {"Mês": mes, "Cenário": "Apertado", "Saldo projetado": saldo_apertado},
                ]
            )
bg_base64 = image_to_base64(BACKGROUND_IMAGE)
bg_css = f"url('data:image/jpeg;base64,{bg_base64}') center/cover fixed no-repeat" if bg_base64 else "linear-gradient(145deg, #eef1f5, #f8f9fb)"

        df_projecoes = pd.DataFrame(projecoes)
        normal_negativo = df_projecoes[
            (df_projecoes["Cenário"] == "Normal") & (df_projecoes["Saldo projetado"] < 0)
        ]
        if normal_negativo.empty:
            resposta_futuro = "Pelo cenário normal, seu saldo continua positivo nos próximos 6 meses."
        else:
            mes_alerta = normal_negativo.iloc[0]["Mês"].strftime("%m/%Y")
            resposta_futuro = f"No cenário normal, seu saldo pode ficar negativo em {mes_alerta}."
st.markdown(
    f"""
<style>
    :root {{
        --ink: #111318;
        --muted: #747985;
        --panel: rgba(246,247,249,.88);
        --card: rgba(255,255,255,.92);
        --lime: #d9ff00;
        --blue: #8fb1ff;
        --line: rgba(17,19,24,.08);
        --shadow: 0 26px 80px rgba(17,19,24,.12);
        --soft: 0 14px 38px rgba(17,19,24,.08);
    }}
    html, body, [class*="css"] {{
        font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}
    [data-testid="stAppViewContainer"] {{
        background:
            radial-gradient(circle at 12% 8%, rgba(255,255,255,.95), transparent 22rem),
            radial-gradient(circle at 88% 12%, rgba(217,255,0,.18), transparent 24rem),
            radial-gradient(circle at 72% 86%, rgba(143,177,255,.22), transparent 28rem),
            {bg_css};
        color: var(--ink);
    }}
    [data-testid="stHeader"], [data-testid="stToolbar"] {{
        background: transparent;
    }}
    .main .block-container {{
        max-width: 1220px;
        padding-top: 1.6rem;
        padding-bottom: 3rem;
    }}
    h1, h2, h3 {{
        color: var(--ink) !important;
        letter-spacing: 0;
    }}
    .classic-header {{
        display: grid;
        grid-template-columns: minmax(0, 1.45fr) minmax(18rem, .8fr);
        gap: 1rem;
        margin-bottom: 1.15rem;
    }}
    .hero-card, .overview-card {{
        border-radius: 30px;
        border: 1px solid rgba(255,255,255,.82);
        background: rgba(245,246,249,.82);
        box-shadow: var(--shadow);
        backdrop-filter: blur(20px);
    }}
    .hero-card {{
        min-height: 22rem;
        padding: clamp(1.4rem, 3vw, 2.25rem);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }}
    .hero-top {{
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        align-items: center;
        color: var(--muted);
        font-weight: 900;
    }}
    .brand-mark {{
        display: inline-flex;
        align-items: center;
        gap: .65rem;
    }}
    .brand-mark span {{
        width: 2.2rem;
        height: 2.2rem;
        display: inline-grid;
        place-items: center;
        border-radius: 50%;
        color: var(--ink);
        background: var(--lime);
        box-shadow: var(--soft);
    }}
    .header-pill {{
        display: inline-flex;
        align-items: center;
        min-height: 2.35rem;
        padding: 0 1rem;
        border-radius: 999px;
        color: #fff;
        background: #111318;
        box-shadow: var(--soft);
        font-size: .86rem;
    }}
    .hero-card h1 {{
        margin: 2.8rem 0 .85rem;
        font-size: clamp(2.65rem, 5.8vw, 5.2rem);
        line-height: .96;
        font-weight: 900;
    }}
    .hero-card p {{
        max-width: 43rem;
        margin: 0;
        color: var(--muted);
        line-height: 1.65;
        font-weight: 620;
    }}
    .header-quote {{
        margin-top: 1.6rem;
        padding-left: 1rem;
        border-left: 3px solid var(--lime);
        color: var(--ink);
        font-weight: 760;
        line-height: 1.45;
    }}
    .header-quote span {{
        display: block;
        margin-top: .35rem;
        color: var(--muted);
        font-size: .86rem;
    }}
    .overview-card {{
        padding: 1.15rem;
    }}
    .overview-head {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 1rem;
        margin-bottom: .9rem;
        font-weight: 900;
    }}
    .overview-head span {{
        color: var(--muted);
        font-size: .78rem;
        text-transform: uppercase;
    }}
    .overview-grid {{
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: .75rem;
    }}
    .overview-item {{
        min-height: 7.6rem;
        padding: .95rem;
        border-radius: 20px;
        background: rgba(255,255,255,.88);
        border: 1px solid rgba(255,255,255,.94);
        box-shadow: var(--soft);
    }}
    .overview-label {{
        color: var(--muted);
        font-size: .74rem;
        font-weight: 850;
        text-transform: uppercase;
    }}
    .overview-value {{
        margin-top: 1.1rem;
        color: var(--ink);
        font-size: clamp(1.15rem, 2vw, 1.65rem);
        line-height: 1.05;
        font-weight: 900;
    }}
    .overview-action {{
        margin-top: .9rem;
        min-height: 3rem;
        display: grid;
        place-items: center;
        border-radius: 999px;
        color: #fff;
        background: #111318;
        font-weight: 900;
    }}
    .metric-grid, .indicator-grid {{
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: .85rem;
        margin: 1rem 0 1.15rem;
    }}
    .metric-card, .indicator-card, .history-card {{
        position: relative;
        overflow: hidden;
        min-height: 9rem;
        padding: 1.1rem;
        border-radius: 26px;
        background: var(--card);
        border: 1px solid rgba(255,255,255,.92);
        box-shadow: var(--soft);
    }}
    .accent {{
        background: linear-gradient(135deg, var(--lime), #caff00);
    }}
    .blue {{
        background: linear-gradient(135deg, #fff, #e9efff);
    }}
    .dark {{
        background: #111318;
        color: white;
    }}
    .dark .metric-label, .dark .metric-foot, .dark .metric-value {{
        color: white;
    }}
    .metric-label, .indicator-top {{
        color: var(--muted);
        font-size: .74rem;
        font-weight: 850;
        text-transform: uppercase;
    }}
    .metric-value, .indicator-value {{
        margin-top: 1rem;
        color: var(--ink);
        font-size: clamp(1.45rem, 3vw, 2.25rem);
        line-height: 1;
        font-weight: 900;
    }}
    .metric-foot, .indicator-note {{
        margin-top: .5rem;
        color: var(--muted);
        font-size: .84rem;
        line-height: 1.4;
    }}
    .quote, .chart-intro, .history-summary {{
        margin: .9rem 0;
        padding: 1rem 1.1rem;
        border-radius: 24px;
        background: rgba(255,255,255,.88);
        border: 1px solid rgba(255,255,255,.92);
        box-shadow: var(--soft);
        color: var(--ink);
    }}
    .quote strong, .chart-intro strong, .history-summary strong {{
        color: var(--ink);
    }}
    div[data-testid="stTabs"] button {{
        border-radius: 999px;
        color: var(--ink);
        font-weight: 800;
        background: rgba(255,255,255,.76);
        border: 1px solid rgba(255,255,255,.9);
        box-shadow: var(--soft);
    }}
    div[data-testid="stTabs"] button[aria-selected="true"] {{
        color: white;
        background: #111318;
        border-color: #111318;
    }}
    div[data-testid="stPlotlyChart"], div[data-testid="stDataFrame"], div[data-testid="stExpander"], div[data-testid="stForm"], [data-testid="stAlert"] {{
        border-radius: 24px;
        background: rgba(255,255,255,.9);
        border: 1px solid rgba(255,255,255,.92);
        box-shadow: var(--soft);
    }}
    .positive {{
        color: #0d906f;
        font-weight: 900;
    }}
    .negative {{
        color: #cc4a5b;
        font-weight: 900;
    }}
    @media (max-width: 980px) {{
        .classic-header, .overview-grid, .metric-grid, .indicator-grid {{
            grid-template-columns: 1fr;
            display: grid;
        }}
    }}
</style>
""",
    unsafe_allow_html=True,
)

        st.markdown(
            f"""<div class="history-summary"><strong>Vou ficar sem dinheiro?</strong><br>{resposta_futuro}</div>""",
            unsafe_allow_html=True,
        )
        fig_futuro = px.line(
            df_projecoes,
            x="Mês",
            y="Saldo projetado",
            color="Cenário",
            title="Como deve ficar meu dinheiro nos próximos meses?",
            markers=True,
            color_discrete_map={
                "Bom": "#0d906f",
                "Normal": "#28c7b7",
                "Apertado": "#cc4a5b",
            },
        )
        fig_futuro.update_yaxes(tickprefix="R$ ")
        st.plotly_chart(style_plot(fig_futuro), use_container_width=True)

    with st.expander("📤 Subir planilha de movimentações", expanded=False):
        st.caption(
            "Os registros novos serão integrados ao mesmo histórico, saldo e gráficos dos "
            "cadastros manuais. Lançamentos já existentes não serão duplicados."
        )
init_db()
df = carregar_dados()
investimentos = carregar_investimentos()
dividas = carregar_dividas()

        st.download_button(
            "Baixar modelo Excel com tabela de valores",
            data=gerar_modelo_excel(),
            file_name="modelo-dashboard-financeiro.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
entradas = df[df["valor"] > 0]["valor"].sum() if len(df) else 0
saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) else 0
saldo = df["valor"].sum() if len(df) else 0
total_investido = investimentos["valor"].sum() if len(investimentos) else 0
if len(dividas):
    dividas_base = dividas.copy()
    dividas_base["saldo_base"] = dividas_base["saldo_negociado"].where(dividas_base["saldo_negociado"] > 0, dividas_base["saldo_original"])
    total_dividas = dividas_base[dividas_base["status"] != "Quitada"]["saldo_base"].sum()
else:
    total_dividas = 0
taxa_sobra = (saldo / entradas) * 100 if entradas > 0 else 0
comprometimento = (saidas / entradas) * 100 if entradas > 0 else 0
score = score_financeiro(entradas, saidas, saldo, total_investido, total_dividas)

        arquivo_planilha = st.file_uploader(
            "Escolha a planilha",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=False,
        )
st.markdown(
    f"""
<section class="classic-header">
    <div class="hero-card">
        <div class="hero-top">
            <div class="brand-mark"><span>DF</span><strong>Dashboard Financeiro</strong></div>
            <div class="header-pill">Finance 2025</div>
        </div>
        <div>
            <span class="metric-label">Painel inteligente</span>
            <h1>Dashboard<br>Financeiro</h1>
            <p>Controle entradas, gastos, dívidas, investimentos e histórico em uma visão clara para decidir melhor todos os meses.</p>
            <div class="header-quote">
                "Preço é o que você paga; valor é o que você recebe."
                <span>Benjamin Graham - use seus números para proteger valor, não apenas pagar contas.</span>
            </div>
        </div>
    </div>
    <aside class="overview-card">
        <div class="overview-head">
            <strong>Visão geral</strong>
            <span>Atualizado agora</span>
        </div>
        <div class="overview-grid">
            <div class="overview-item"><div class="overview-label">Saldo atual</div><div class="overview-value">{brl(saldo)}</div></div>
            <div class="overview-item"><div class="overview-label">Entradas</div><div class="overview-value">{brl(entradas)}</div></div>
            <div class="overview-item"><div class="overview-label">Saídas</div><div class="overview-value">{brl(saidas)}</div></div>
            <div class="overview-item"><div class="overview-label">Investimentos</div><div class="overview-value">{brl(total_investido)}</div></div>
        </div>
        <div class="overview-action">Saúde financeira: {score}/100</div>
    </aside>
</section>
""",
    unsafe_allow_html=True,
)

        if arquivo_planilha is not None:
            try:
                df_planilha = ler_planilha_movimentacoes(arquivo_planilha)
                df_importado, linhas_ignoradas = preparar_movimentacoes_importadas(df_planilha)
col_top1, col_top2, col_top3, col_top4 = st.columns([1, 1, 1.2, 1.2])
with col_top1:
    if st.button("Atualizar painel", use_container_width=True):
        st.rerun()
with col_top2:
    st.download_button(
        "Baixar PDF",
        data=gerar_pdf(df, investimentos, dividas),
        file_name=f"relatorio-financeiro-{date.today().strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
with col_top3:
    st.info("Para cadastrar, use a aba Nova Movimentacao.")
with col_top4:
    st.info("Para importar planilha, abra a aba Dashboard.")

                if df_importado.empty:
                    st.warning("Não encontrei movimentações válidas nessa planilha.")
                else:
                    df_novo, duplicadas_encontradas = conciliar_movimentacoes(df_importado, df)
                    entradas_importadas = df_novo[df_novo["valor"] > 0]["valor"].sum()
                    saidas_importadas = abs(df_novo[df_novo["valor"] < 0]["valor"].sum())
                    saldo_importado = df_novo["valor"].sum()
aba = st.tabs(["Nova Movimentacao", "Dashboard", "Investimentos", "Dividas", "Historico"])

                    col_valor1, col_valor2, col_valor3, col_valor4 = st.columns(4)
                    col_valor1.metric("Novas entradas", brl(entradas_importadas))
                    col_valor2.metric("Novas saídas", brl(saidas_importadas))
                    col_valor3.metric("Impacto no saldo", brl(saldo_importado))
                    col_valor4.metric("Novos registros", len(df_novo))
with aba[0]:
    st.subheader("Adicionar Nova Movimentacao")
    tipo_selecionado = st.radio("Tipo", ["Entrada", "Saida"], horizontal=True)
    with st.form("form_movimentacao"):
        col1, col2 = st.columns(2)
        with col1:
            data_mov = st.date_input("Data", value=date.today(), format="DD/MM/YYYY")
            descricao = st.text_input("Descricao")
            if tipo_selecionado == "Entrada":
                categoria = st.selectbox("Categoria", ["Salario", "Rendas Extras", "Freelance", "Reembolso", "Outro"])
            else:
                categoria = st.selectbox("Categoria", ["Mercado", "Aluguel", "Contas", "Lazer", "Roupa", "Beleza", "Transporte", "Dividas", "Outro"])
        with col2:
            valor = st.number_input("Valor R$", value=0.0, min_value=0.0, step=0.01)
            cartao = st.text_input("Forma de pagamento")
        if st.form_submit_button("Salvar movimentacao"):
            if not descricao.strip():
                st.error("Informe uma descricao.")
            elif valor <= 0:
                st.error("Informe um valor maior que zero.")
            else:
                salvar_transacao(data_mov, descricao.strip(), categoria, valor if tipo_selecionado == "Entrada" else -abs(valor), tipo_selecionado, cartao.strip())
                st.success("Movimentacao salva.")
                st.rerun()

                    previa = df_novo.rename(
                        columns={
                            "data": "Data",
                            "descricao": "Descrição",
                            "categoria": "Categoria",
                            "valor": "Valor (R$)",
                            "tipo": "Tipo",
                            "cartao": "Forma de Pagamento",
                        }
                    )
                    previa["Data"] = pd.to_datetime(previa["Data"], errors="coerce")
                    previa["Valor (R$)"] = previa["Valor (R$)"].map(brl)
with aba[1]:
    st.subheader("Dashboard")
    st.markdown(
        f"""
<div class="indicator-grid">
    <div class="indicator-card"><div class="indicator-top">Score financeiro</div><div class="indicator-value">{score}/100</div><div class="indicator-note">Combina saldo, dividas e investimentos.</div></div>
    <div class="indicator-card"><div class="indicator-top">Taxa de sobra</div><div class="indicator-value">{pct(taxa_sobra)}</div><div class="indicator-note">Parte das entradas que virou saldo.</div></div>
    <div class="indicator-card"><div class="indicator-top">Comprometimento</div><div class="indicator-value">{pct(comprometimento)}</div><div class="indicator-note">Saidas comparadas as entradas.</div></div>
    <div class="indicator-card"><div class="indicator-top">Dividas abertas</div><div class="indicator-value">{brl(total_dividas)}</div><div class="indicator-note">Saldo monitorado para negociar.</div></div>
</div>
""",
        unsafe_allow_html=True,
    )

                    st.markdown("**Novos registros que serão integrados**")
                    if linhas_ignoradas:
                        st.caption(f"{linhas_ignoradas} linhas foram ignoradas por não terem valor válido.")
                    if duplicadas_encontradas:
                        st.info(
                            f"{duplicadas_encontradas} lançamentos já cadastrados foram "
                            "reconhecidos e não serão duplicados."
                        )
                    if df_novo.empty:
                        st.info("Todos os lançamentos dessa planilha já estão integrados ao dashboard.")
                    else:
                        st.dataframe(
                            previa,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                                "Valor (R$)": st.column_config.TextColumn("Valor (R$)"),
                            },
                        )

                        if st.button("Integrar movimentações ao dashboard", type="primary"):
                            total_importado, total_duplicadas = importar_movimentacoes(df_importado)
                            mensagem = (
                                f"{total_importado} movimentações integradas. "
                                "Totais, gráficos e histórico foram atualizados."
                            )
                            if total_duplicadas:
                                mensagem += f" {total_duplicadas} duplicadas foram ignoradas."
                            st.session_state["resultado_importacao"] = mensagem
                            st.rerun()
    with st.expander("Subir planilha de movimentacoes", expanded=False):
        st.caption("Registros novos entram no mesmo historico e nos mesmos graficos. Duplicados sao ignorados.")
        st.download_button(
            "Baixar modelo Excel",
            data=modelo_excel(),
            file_name="modelo-dashboard-financeiro.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        arquivo = st.file_uploader("Escolha a planilha", type=["csv", "xlsx", "xls"])
        if arquivo is not None:
            try:
                dados = ler_planilha(arquivo)
                df_importado, ignoradas = preparar_importacao(dados)
                df_novo, duplicadas = conciliar_movimentacoes(df_importado, df)
                col_a, col_b, col_c, col_d = st.columns(4)
                col_a.metric("Novas entradas", brl(df_novo[df_novo["valor"] > 0]["valor"].sum() if len(df_novo) else 0))
                col_b.metric("Novas saidas", brl(abs(df_novo[df_novo["valor"] < 0]["valor"].sum()) if len(df_novo) else 0))
                col_c.metric("Impacto no saldo", brl(df_novo["valor"].sum() if len(df_novo) else 0))
                col_d.metric("Novos registros", len(df_novo))
                if ignoradas:
                    st.caption(f"{ignoradas} linhas foram ignoradas por nao terem valor valido.")
                if duplicadas:
                    st.info(f"{duplicadas} lancamentos duplicados foram reconhecidos e ignorados.")
                if len(df_novo):
                    previa = df_novo.rename(columns={"data": "Data", "descricao": "Descricao", "categoria": "Categoria", "valor": "Valor", "tipo": "Tipo", "cartao": "Pagamento"})
                    previa["Data"] = previa["Data"].map(data_br)
                    previa["Valor"] = previa["Valor"].map(brl)
                    st.dataframe(previa, use_container_width=True, hide_index=True)
                    if st.button("Integrar movimentacoes ao dashboard", type="primary"):
                        total, total_duplicadas = importar_movimentacoes(df_importado)
                        st.success(f"{total} movimentacoes integradas. {total_duplicadas} duplicadas ignoradas.")
                        st.rerun()
                else:
                    st.info("Todos os lancamentos dessa planilha ja estao no dashboard ou nao foram reconhecidos.")
            except Exception as erro:
                st.error(f"Não consegui importar essa planilha: {mensagem_erro_usuario(erro)}")
                st.error(f"Nao consegui importar essa planilha: {erro}")

    if len(df) > 0:
    if len(df):
        df_chart = df.copy()
        df_chart["categoria"] = df_chart["categoria"].replace("", "Sem categoria").fillna("Sem categoria")
        df_chart["cartao"] = df_chart["cartao"].replace("", "Não informado").fillna("Não informado")
        df_chart["valor_abs"] = df_chart["valor"].abs()
        df_chart["data_convertida"] = pd.to_datetime(df_chart["data"], errors="coerce")

        categoria_total = (
            df_chart.groupby("categoria", as_index=False)["valor_abs"]
            .sum()
            .sort_values("valor_abs", ascending=False)
        )

        fluxo_total = df_chart.groupby(["categoria", "tipo"], as_index=False)["valor_abs"].sum()
        fluxo_mensal, categorias_dashboard, pagamentos_dashboard = preparar_dados_dashboard(df_chart)
        df_chart["valor_abs"] = df_chart["valor"].abs()
        df_chart["categoria"] = df_chart["categoria"].replace("", "Sem categoria")
        df_chart["cartao"] = df_chart["cartao"].replace("", "Nao informado")
        df_timeline = df_chart.dropna(subset=["data_convertida"]).sort_values("data_convertida")
        if len(df_timeline):
            df_timeline["saldo_acumulado"] = df_timeline["valor"].cumsum()
            fig_saldo = px.area(df_timeline, x="data_convertida", y="saldo_acumulado", title="Saldo acumulado ao longo do tempo")
            fig_saldo.update_traces(line=dict(color="#111318", width=3), fillcolor="rgba(217,255,0,.34)")
            fig_saldo.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig_saldo), use_container_width=True)

        st.markdown(
            f"""<div class="chart-intro"><strong>Visão consolidada</strong><br>
            Os {len(df)} lançamentos cadastrados manualmente e integrados por planilha alimentam
            automaticamente os gráficos e o histórico abaixo.</div>""",
            unsafe_allow_html=True,
        )

        if len(df_timeline):
            col_trend, col_bubbles = st.columns([1.25, 0.75])
            with col_trend:
                fig_saldo = px.area(
                    df_timeline,
                    x="data_convertida",
                    y="saldo_acumulado",
                    title="Saldo acumulado ao longo do tempo",
                    labels={"data_convertida": "Data", "saldo_acumulado": "Saldo acumulado"},
                )
                fig_saldo.update_traces(
                    line=dict(color="#111318", width=3),
                    fillcolor="rgba(217,255,0,0.34)",
                    hovertemplate="<b>%{x|%d/%m/%Y}</b><br>R$ %{y:,.2f}<extra></extra>",
                )
                fig_saldo.update_xaxes(tickformat="%d/%m/%Y")
                fig_saldo.update_yaxes(tickprefix="R$ ")
                st.plotly_chart(style_plot(fig_saldo), use_container_width=True)

            with col_bubbles:
                entrada_bolha = brl_compacto(total_entradas)
                saida_bolha = brl_compacto(total_saidas)
                saldo_bolha = brl_compacto(saldo)
                st.markdown(
                    f"""<div class="visitor-panel">
                        <div class="visitor-head">Radar financeiro <span>Atual</span></div>
                        <div class="bubble-wrap">
                            <div class="bubble income">{entrada_bolha}<small>Entradas</small></div>
                            <div class="bubble expense">{saida_bolha}<small>Saídas</small></div>
                            <div class="bubble balance">{saldo_bolha}<small>Saldo</small></div>
                        </div>
                        <div class="target-list">
                            <div class="target-row"><span>Saúde financeira</span><strong>{score_dashboard}%</strong><div class="target-line"><span style="width:{limitar_percentual(score_dashboard)}%"></span></div></div>
                            <div class="target-row"><span>Metas</span><strong>{pct(hero_progresso_metas)}</strong><div class="target-line"><span style="width:{limitar_percentual(hero_progresso_metas)}%"></span></div></div>
                            <div class="target-row"><span>Reserva</span><strong>{pct(reserva_percentual)}</strong><div class="target-line"><span style="width:{reserva_percentual}%"></span></div></div>
                        </div>
                    </div>""",
                    unsafe_allow_html=True,
                )

        col_g1, col_g2 = st.columns(2)

        categoria_total = df_chart.groupby("categoria", as_index=False)["valor_abs"].sum().sort_values("valor_abs", ascending=False)
        fluxo = df_chart.groupby(["categoria", "tipo"], as_index=False)["valor_abs"].sum()
        with col_g1:
            fig = px.pie(
                categoria_total,
                names="categoria",
                values="valor_abs",
                title="Distribuição por Categoria",
                hole=0.58,
                color_discrete_sequence=[
                    "#d9ff00",
                    "#111318",
                    "#8fb1ff",
                    "#d7dbe2",
                    "#efffb4",
                    "#5c6573",
                    "#c2cff7",
                ],
            )
            fig.update_traces(
                textposition="outside",
                textinfo="percent+label",
                insidetextorientation="radial",
                marker=dict(line=dict(color="rgba(255,255,255,0.94)", width=2)),
                pull=[0.018] * len(categoria_total),
                hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>",
            )
            fig = style_plot(fig)
            fig.update_layout(
                height=520,
                showlegend=True,
                uniformtext_minsize=11,
                uniformtext_mode="hide",
                margin=dict(l=52, r=92, t=88, b=108),
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.12,
                    xanchor="center",
                    x=0.5,
                    bgcolor="rgba(255,255,255,0)",
                    font=dict(color="#111318", size=11),
                ),
            )
            st.plotly_chart(fig, use_container_width=True)

            fig = px.pie(categoria_total, names="categoria", values="valor_abs", hole=0.58, title="Distribuicao por categoria", color_discrete_sequence=["#d9ff00", "#111318", "#8fb1ff", "#d7dbe2", "#efffb4", "#5c6573"])
            fig.update_traces(textposition="outside", textinfo="percent+label")
            st.plotly_chart(style_plot(fig), use_container_width=True)
        with col_g2:
            fig2 = px.bar(
                fluxo_total,
                x="categoria",
                y="valor_abs",
                color="tipo",
                title="Entradas x Saídas",
                color_discrete_map={
                    "Entrada": "#d9ff00",
                    "Saída": "#111318",
                },
                labels={"categoria": "Categoria", "valor_abs": "Valor", "tipo": "Tipo"},
            )
            fig2.update_traces(
                marker_line_color="rgba(255,255,255,0.90)",
                marker_line_width=1,
                hovertemplate="<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>",
            )
            fig2 = px.bar(fluxo, x="categoria", y="valor_abs", color="tipo", title="Entradas x Saidas", color_discrete_map={"Entrada": "#d9ff00", "Saida": "#111318"})
            fig2.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig2), use_container_width=True)

        top_categorias = categorias_dashboard.head(8).sort_values("valor_abs", ascending=True)
        pagamentos_top = pagamentos_dashboard.head(7)
        col_top, col_pagamento = st.columns(2)

        with col_top:
            fig_top = px.bar(
                top_categorias,
                x="valor_abs",
                y="categoria",
                orientation="h",
                title="Categorias que mais movimentam dinheiro",
                labels={"valor_abs": "Valor", "categoria": "Categoria"},
                color="valor_abs",
                color_continuous_scale=["#e9edf2", "#8fb1ff", "#111318"],
            )
            fig_top.update_traces(
                hovertemplate="<b>%{y}</b><br>R$ %{x:,.2f}<extra></extra>",
                marker_line_color="rgba(255,255,255,0.90)",
                marker_line_width=1,
            )
            fig_top.update_layout(coloraxis_showscale=False)
            fig_top.update_xaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig_top), use_container_width=True)

        with col_pagamento:
            fig_pagamento = px.pie(
                pagamentos_top,
                names="cartao",
                values="valor_abs",
                hole=0.62,
                title="Formas de pagamento",
                color_discrete_sequence=[
                    "#111318",
                    "#d9ff00",
                    "#8fb1ff",
                    "#d7dbe2",
                    "#efffb4",
                    "#5c6573",
                ],
            )
            fig_pagamento.update_traces(
                textposition="inside",
                textinfo="percent",
                marker=dict(line=dict(color="rgba(255,255,255,0.94)", width=2)),
                hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>",
            )
            fig_pagamento = style_plot(fig_pagamento)
            fig_pagamento.update_layout(height=410, legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center"))
            st.plotly_chart(fig_pagamento, use_container_width=True)

        if len(fluxo_mensal) > 0:
            fig_mensal = px.bar(
                fluxo_mensal,
                x="mes",
                y=["Entradas", "Saídas"],
                barmode="group",
                title="Evolução mensal consolidada",
                color_discrete_map={"Entradas": "#d9ff00", "Saídas": "#111318"},
                labels={"mes": "Mês", "value": "Valor", "variable": "Movimentação"},
            )
            fig_mensal.for_each_trace(
                lambda trace: trace.update(
                    marker_line_color="rgba(255,255,255,0.90)",
                    marker_line_width=1,
                    hovertemplate="<b>%{x|%m/%Y}</b><br>R$ %{y:,.2f}<extra></extra>",
                )
            )
            fig_mensal.update_xaxes(tickformat="%m/%Y")
            fig_mensal.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig_mensal), use_container_width=True)
        col_g3, col_g4 = st.columns(2)
        with col_g3:
            top_categorias = categoria_total.head(8).sort_values("valor_abs", ascending=True)
            fig3 = px.bar(top_categorias, x="valor_abs", y="categoria", orientation="h", title="Categorias que mais movimentam dinheiro", color="valor_abs", color_continuous_scale=["#e9edf2", "#8fb1ff", "#111318"])
            fig3.update_layout(coloraxis_showscale=False)
            fig3.update_xaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig3), use_container_width=True)
        with col_g4:
            pagamentos = df_chart.groupby("cartao", as_index=False)["valor_abs"].sum().sort_values("valor_abs", ascending=False).head(7)
            fig4 = px.pie(pagamentos, names="cartao", values="valor_abs", hole=0.62, title="Formas de pagamento", color_discrete_sequence=["#111318", "#d9ff00", "#8fb1ff", "#d7dbe2", "#efffb4"])
            st.plotly_chart(style_plot(fig4), use_container_width=True)
    else:
        st.info("Adicione uma movimentação para visualizar os gráficos.")
        st.info("Cadastre ou importe movimentacoes para ver os graficos.")

# ====================== ABA 3 ======================
with aba[2]:
    st.subheader("Metas e Investimentos")
    st.markdown(
        """<div class="chart-intro"><strong>Minhas metas</strong><br>
        Defina objetivos simples, acompanhe quanto falta e veja em quanto tempo chega lá
        mantendo um aporte mensal.</div>""",
        unsafe_allow_html=True,
    )

    with st.form("nova_meta"):
        st.markdown("#### Adicionar meta")
        col_meta1, col_meta2, col_meta3 = st.columns(3)

        with col_meta1:
            data_meta = st.date_input("Data da meta", value=date.today(), format="DD/MM/YYYY")
            nome_meta = st.text_input(
                "Nome da meta",
                placeholder="Ex.: Reserva, reforma, viagem, quitar dívida",
            )
            status_meta = st.selectbox("Status da meta", ["Em andamento", "Planejada", "Concluída"])

        with col_meta2:
            valor_meta = st.number_input(
                "Valor necessário (R$)",
                value=0.0,
                step=0.01,
                min_value=0.0,
                key="valor_meta",
            )
            valor_atual_meta = st.number_input(
                "Quanto já tenho (R$)",
                value=0.0,
                step=0.01,
                min_value=0.0,
                key="valor_atual_meta",
            )
            aporte_meta = st.number_input(
                "Quanto posso guardar por mês (R$)",
                value=0.0,
                step=0.01,
                min_value=0.0,
                key="aporte_meta",
            )

        with col_meta3:
            prazo_meta = st.text_input("Prazo desejado", placeholder="Ex.: Dezembro/2026")
            anotacoes_meta = st.text_area(
                "Anotações",
                placeholder="Ex.: guardar após receber, usar 13º, reduzir delivery",
            )

        if st.form_submit_button("Salvar meta"):
            if not nome_meta.strip():
                st.error("Informe o nome da meta.")
            elif valor_meta <= 0:
                st.error("Informe o valor necessário para a meta.")
            else:
                try:
                    salvar_meta(
                        data_meta,
                        nome_meta.strip(),
                        valor_meta,
                        valor_atual_meta,
                        aporte_meta,
                        prazo_meta.strip(),
                        status_meta,
                        anotacoes_meta.strip(),
                    )
                    st.success("Meta salva com sucesso!")
                    st.rerun()
                except Exception as erro:
                    st.error(f"Não foi possível salvar a meta: {mensagem_erro_usuario(erro)}")

    if len(df_metas) > 0:
        st.markdown("#### Progresso das metas")
        st.markdown('<div class="goal-grid">', unsafe_allow_html=True)
        for _, meta in df_metas.iterrows():
            valor_meta = max(float(meta["valor_meta"]), 0)
            valor_atual = max(float(meta["valor_atual"]), 0)
            aporte_mensal = max(float(meta["aporte_mensal"]), 0)
            percentual = min((valor_atual / valor_meta) * 100, 100) if valor_meta > 0 else 0
            falta = max(valor_meta - valor_atual, 0)
            meses = math.ceil(falta / aporte_mensal) if falta > 0 and aporte_mensal > 0 else 0
            previsao = (
                f"Chega em {texto_meses(meses)}"
                if falta > 0 and aporte_mensal > 0
                else "Meta concluída" if falta <= 0 else "Defina um aporte mensal"
            )
            nome = escape(str(meta["nome"] or "Meta sem nome"))
            status = escape(str(meta["status"] or "Sem status"))
            prazo = escape(str(meta["prazo"] or "Sem prazo definido"))
            anotacoes = escape(str(meta["anotacoes"] or "Sem anotações"))

            st.markdown(
                f"""<div class="goal-card"><div class="goal-title">{nome}</div><div class="goal-meta">Status: {status} • Prazo: {prazo}</div><div class="goal-progress"><span style="width:{percentual:.1f}%"></span></div><div class="goal-meta"><strong>{percentual:.0f}% concluída</strong><br>Tenho {brl(valor_atual)} de {brl(valor_meta)}. Falta {brl(falta)}.<br>{previsao}. Aporte mensal: {brl(aporte_mensal)}.<br>Anotações: {anotacoes}</div></div>""",
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

        for _, meta in df_metas.iterrows():
            if st.button("🗑️ Apagar meta", key=f"del_meta{meta['id']}"):
                excluir_meta(meta["id"])
                st.rerun()
    else:
        st.info("Nenhuma meta cadastrada ainda.")

    st.divider()
    st.markdown("### Investimentos")

    with st.form("novo_investimento"):
        st.markdown("#### Adicionar investimento")
    st.subheader("Investimentos")
    with st.form("form_investimento"):
        col1, col2, col3 = st.columns(3)

        with col1:
            data_investimento = st.date_input(
                "Data do investimento",
                value=date.today(),
                format="DD/MM/YYYY",
            )
            tipo_investimento = st.selectbox(
                "Tipo de investimento",
                [
                    "Reserva de emergência",
                    "Tesouro Direto",
                    "CDB",
                    "LCI / LCA",
                    "Fundo de investimento",
                    "Ações",
                    "FII",
                    "Previdência privada",
                    "Criptomoedas",
                    "Outro",
                ],
            )

            data_inv = st.date_input("Data", value=date.today(), format="DD/MM/YYYY", key="data_inv")
            tipo_inv = st.selectbox("Tipo", ["Reserva de emergencia", "Tesouro Direto", "CDB", "LCI / LCA", "Acoes", "FII", "Outro"])
        with col2:
            valor_investimento = st.number_input(
                "Valor investido (R$)",
                value=0.0,
                step=0.01,
                min_value=0.0,
            )
            rentabilidade = st.text_input(
                "Rentabilidade",
                placeholder="Ex.: 12% a.a. ou 105% do CDI",
            )

            valor_inv = st.number_input("Valor investido", value=0.0, min_value=0.0, step=0.01)
            rentabilidade = st.text_input("Rentabilidade")
        with col3:
            descricao_investimento = st.text_input(
                "Descrição",
                placeholder="Ex.: CDB Banco X - reserva",
            )
            status_investimento = st.selectbox(
                "Status",
                ["Ativo", "Planejado", "Resgatado"],
            )

            descricao_inv = st.text_input("Descricao do investimento")
            status_inv = st.selectbox("Status", ["Ativo", "Planejado", "Resgatado"])
        if st.form_submit_button("Salvar investimento"):
            if valor_investimento <= 0:
                st.error("Informe um valor de investimento maior que zero.")
            if valor_inv <= 0:
                st.error("Informe um valor maior que zero.")
            else:
                try:
                    salvar_investimento(
                        data_investimento,
                        tipo_investimento,
                        valor_investimento,
                        rentabilidade,
                        descricao_investimento,
                        status_investimento,
                    )
                    st.success("Investimento salvo com sucesso!")
                    st.rerun()
                except Exception as erro:
                    st.error(f"Não foi possível salvar o investimento: {mensagem_erro_usuario(erro)}")

    total_investido = df_investimentos["valor"].sum() if len(df_investimentos) > 0 else 0
    total_ativos = (
        df_investimentos[df_investimentos["status"] == "Ativo"]["valor"].sum()
        if len(df_investimentos) > 0
        else 0
    )
    total_planejado = (
        df_investimentos[df_investimentos["status"] == "Planejado"]["valor"].sum()
        if len(df_investimentos) > 0
        else 0
    )

    st.markdown(
        f"""<div class="investment-grid"><div class="metric-card"><div class="metric-label">Patrimônio registrado</div><div class="metric-value">{brl(total_investido)}</div><div class="metric-foot">Todos os investimentos</div></div><div class="metric-card"><div class="metric-label">Investimentos ativos</div><div class="metric-value">{brl(total_ativos)}</div><div class="metric-foot">Valores em andamento</div></div><div class="metric-card"><div class="metric-label">Aportes planejados</div><div class="metric-value">{brl(total_planejado)}</div><div class="metric-foot">Próximos objetivos</div></div></div>""",
        unsafe_allow_html=True,
    )

    if len(df_investimentos) > 0:
        investimentos_tipo = (
            df_investimentos.groupby("tipo", as_index=False)["valor"]
            .sum()
            .sort_values("valor", ascending=False)
        )

        col_i1, col_i2 = st.columns([1, 1])

        with col_i1:
            fig_investimentos = px.pie(
                investimentos_tipo,
                names="tipo",
                values="valor",
                title="Distribuição dos investimentos",
                hole=0.58,
                color_discrete_sequence=[
                    "#081b33",
                    "#28c7b7",
                    "#17486f",
                    "#8be6d4",
                    "#0d906f",
                    "#78a8c8",
                ],
            )
            fig_investimentos.update_traces(textposition="inside", textinfo="percent")
            st.plotly_chart(style_plot(fig_investimentos), use_container_width=True)

        with col_i2:
            fig_status = px.bar(
                df_investimentos.groupby("status", as_index=False)["valor"].sum(),
                x="status",
                y="valor",
                color="status",
                title="Valores por status",
                color_discrete_map={
                    "Ativo": "#0d906f",
                    "Planejado": "#28c7b7",
                    "Resgatado": "#17486f",
                },
            )
            st.plotly_chart(style_plot(fig_status), use_container_width=True)

        st.markdown("#### Investimentos cadastrados")

        for _, investimento in df_investimentos.iterrows():
            titulo = escape(str(investimento["descricao"] or "Investimento sem descrição"))
            tipo_item = escape(str(investimento["tipo"] or "Sem categoria"))
            data_item = escape(data_br(investimento["data"]))
            status_item = escape(str(investimento["status"] or "Sem status"))
            retorno_item = escape(str(investimento["rentabilidade"] or "Não informada"))

            st.markdown(
                f"""<div class="investment-item"><div><div class="investment-title">{titulo}</div><div class="investment-meta">{tipo_item} • {data_item} • Status: {status_item}</div></div><div><div class="investment-value">{brl(investimento["valor"])}</div><div class="investment-return">Rentabilidade: {retorno_item}</div></div></div>""",
                unsafe_allow_html=True,
            )

            if st.button("🗑️ Apagar investimento", key=f"del_invest{investimento['id']}"):
                excluir_investimento(investimento["id"])
                salvar_investimento(data_inv, tipo_inv, valor_inv, rentabilidade, descricao_inv, status_inv)
                st.success("Investimento salvo.")
                st.rerun()
    if len(investimentos):
        st.dataframe(investimentos.assign(data=investimentos["data"].map(data_br), valor=investimentos["valor"].map(brl)), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum investimento cadastrado ainda.")
        st.info("Nenhum investimento cadastrado.")

# ====================== ABA 4 ======================
with aba[3]:
    st.subheader("Dívidas e Negociação")
    st.markdown(
        """<div class="chart-intro"><strong>Controle Nome Limpo</strong><br>
        Registre credor, saldo original, saldo negociado, parcela possível, prioridade,
        próxima ação e anotações antes de aceitar qualquer acordo.</div>""",
        unsafe_allow_html=True,
    )

    with st.form("nova_divida"):
        st.markdown("#### Adicionar dívida ou acordo")
    st.subheader("Dividas e negociacao")
    with st.form("form_divida"):
        col1, col2, col3 = st.columns(3)

        with col1:
            data_divida = st.date_input(
                "Data do registro",
                value=date.today(),
                format="DD/MM/YYYY",
            )
            credor = st.text_input("Credor", placeholder="Ex.: Banco, cartão, loja")
            tipo_divida = st.selectbox(
                "Tipo",
                [
                    "Cartão de crédito",
                    "Empréstimo",
                    "Conta atrasada",
                    "Financiamento",
                    "Acordo",
                    "Cheque especial",
                    "Outro",
                ],
            )
            prioridade = st.selectbox("Prioridade", ["Alta", "Média", "Baixa"])

            data_div = st.date_input("Data do registro", value=date.today(), format="DD/MM/YYYY")
            credor = st.text_input("Credor")
            tipo_divida = st.selectbox("Tipo", ["Cartao de credito", "Emprestimo", "Conta atrasada", "Financiamento", "Acordo", "Outro"])
        with col2:
            saldo_original = st.number_input(
                "Saldo original (R$)",
                value=0.0,
                step=0.01,
                min_value=0.0,
            )
            desconto = st.number_input(
                "Desconto / abatimento (R$)",
                value=0.0,
                step=0.01,
                min_value=0.0,
            )
            saldo_negociado = st.number_input(
                "Saldo negociado (R$)",
                value=0.0,
                step=0.01,
                min_value=0.0,
            )
            parcela_possivel = st.number_input(
                "Parcela possível (R$)",
                value=0.0,
                step=0.01,
                min_value=0.0,
            )

            saldo_original = st.number_input("Saldo original", value=0.0, min_value=0.0, step=0.01)
            desconto = st.number_input("Desconto", value=0.0, min_value=0.0, step=0.01)
            saldo_negociado = st.number_input("Saldo negociado", value=0.0, min_value=0.0, step=0.01)
        with col3:
            vencimento_divida = st.date_input(
                "Vencimento / próximo prazo",
                value=date.today(),
                format="DD/MM/YYYY",
            )
            status_divida = st.selectbox(
                "Status",
                ["Mapear", "Negociar", "Acordada", "Em pagamento", "Quitada"],
            )
            consequencia = st.text_input(
                "Consequência se atrasar",
                placeholder="Ex.: juros, bloqueio, negativação",
            )
            proxima_acao = st.text_input(
                "Próxima ação",
                placeholder="Ex.: ligar, pedir desconto, pagar 1ª parcela",
            )

        anotacoes = st.text_area(
            "Anotações da negociação",
            placeholder="Registre propostas, protocolos, datas de contato e condições do acordo.",
        )

        if st.form_submit_button("Salvar dívida"):
            saldo_final = saldo_negociado
            if saldo_final <= 0 and saldo_original > 0:
                saldo_final = max(saldo_original - desconto, 0)

            if not credor.strip():
                st.error("Informe o credor da dívida.")
            elif saldo_original <= 0 and saldo_final <= 0:
                st.error("Informe o saldo original ou o saldo negociado.")
            parcela = st.number_input("Parcela possivel", value=0.0, min_value=0.0, step=0.01)
            vencimento = st.date_input("Vencimento", value=date.today(), format="DD/MM/YYYY")
            prioridade = st.selectbox("Prioridade", ["Alta", "Media", "Baixa"])
        status = st.selectbox("Status", ["Mapear", "Negociar", "Acordada", "Em pagamento", "Quitada"])
        proxima_acao = st.text_input("Proxima acao")
        anotacoes = st.text_area("Anotacoes")
        if st.form_submit_button("Salvar divida"):
            saldo_final = saldo_negociado if saldo_negociado > 0 else max(saldo_original - desconto, 0)
            if not credor.strip() or saldo_final <= 0:
                st.error("Informe o credor e o valor da divida.")
            else:
                try:
                    salvar_divida(
                        data_divida,
                        credor.strip(),
                        tipo_divida,
                        saldo_original,
                        desconto,
                        saldo_final,
                        parcela_possivel,
                        vencimento_divida,
                        prioridade,
                        consequencia.strip(),
                        status_divida,
                        proxima_acao.strip(),
                        anotacoes.strip(),
                    )
                    st.success("Dívida salva com sucesso!")
                    st.rerun()
                except Exception as erro:
                    st.error(f"Não foi possível salvar a dívida: {mensagem_erro_usuario(erro)}")

    if len(df_dividas) > 0:
        df_dividas_view = df_dividas.copy()
        df_dividas_view["saldo_base"] = df_dividas_view["saldo_negociado"].where(
            df_dividas_view["saldo_negociado"] > 0,
            df_dividas_view["saldo_original"],
        )
        abertas = df_dividas_view[df_dividas_view["status"] != "Quitada"]
        total_aberto = abertas["saldo_base"].sum()
        parcelas_acordadas = abertas["parcela_possivel"].sum()
        economia_prevista = (
            df_dividas_view["saldo_original"] - df_dividas_view["saldo_base"]
        ).clip(lower=0).sum()
        prioritarias = len(df_dividas_view[df_dividas_view["prioridade"] == "Alta"])

        st.markdown(
            f"""<div class="debt-grid"><div class="metric-card"><div class="metric-label">Total aberto</div><div class="metric-value">{brl(total_aberto)}</div><div class="metric-foot">Saldo para negociar ou pagar</div></div><div class="metric-card"><div class="metric-label">Parcelas acordadas</div><div class="metric-value">{brl(parcelas_acordadas)}</div><div class="metric-foot">Compromisso mensal possível</div></div><div class="metric-card"><div class="metric-label">Economia prevista</div><div class="metric-value">{brl(economia_prevista)}</div><div class="metric-foot">Descontos registrados</div></div><div class="metric-card"><div class="metric-label">Prioridade alta</div><div class="metric-value">{prioritarias}</div><div class="metric-foot">Dívidas que exigem atenção</div></div></div>""",
            unsafe_allow_html=True,
        )

        with st.expander("⚡ E se eu quiser pagar mais rápido?", expanded=False):
            pagamento_base = parcelas_acordadas if parcelas_acordadas > 0 else 0
            valor_extra_divida = st.slider(
                "Quero pagar a mais por mês",
                min_value=0,
                max_value=5000,
                value=0,
                step=50,
                format="R$ %d",
            )
            pagamento_total = pagamento_base + valor_extra_divida

            if total_aberto > 0 and pagamento_total > 0:
                meses_atuais = math.ceil(total_aberto / pagamento_base) if pagamento_base > 0 else 0
                meses_novos = math.ceil(total_aberto / pagamento_total)
                meses_ganhos = max(meses_atuais - meses_novos, 0) if meses_atuais else 0
                texto_atual = (
                    f"No ritmo atual, faltam cerca de {texto_meses(meses_atuais)}."
                    if meses_atuais
                    else "Defina uma parcela mensal para comparar o prazo."
                )
                st.markdown(
                    f"""<div class="answer-grid">
                        <div class="answer-card answer-care"><div class="answer-question">Prazo atual</div><div class="answer-value">{texto_meses(meses_atuais) if meses_atuais else "Sem parcela"}</div><div class="answer-action">{texto_atual}</div></div>
                        <div class="answer-card answer-good"><div class="answer-question">Novo prazo</div><div class="answer-value">{texto_meses(meses_novos)}</div><div class="answer-action">Com {brl(pagamento_total)} por mês.</div></div>
                        <div class="answer-card answer-good"><div class="answer-question">Tempo ganho</div><div class="answer-value">{texto_meses(meses_ganhos)}</div><div class="answer-action">Estimativa simples, sem recalcular juros futuros.</div></div>
                    </div>""",
                    unsafe_allow_html=True,
                )
            else:
                st.info("Cadastre o saldo aberto e uma parcela possível para simular o prazo.")

        col_dividas, col_status = st.columns([1.25, 0.75])
        with col_dividas:
            tabela_dividas = df_dividas_view[
                [
                    "data",
                    "credor",
                    "tipo",
                    "saldo_original",
                    "desconto",
                    "saldo_base",
                    "parcela_possivel",
                    "vencimento",
                    "prioridade",
                    "status",
                    "proxima_acao",
                ]
            ].rename(
                columns={
                    "data": "Data",
                    "credor": "Credor",
                    "tipo": "Tipo",
                    "saldo_original": "Saldo original",
                    "desconto": "Desconto",
                    "saldo_base": "Saldo negociado",
                    "parcela_possivel": "Parcela possível",
                    "vencimento": "Vencimento",
                    "prioridade": "Prioridade",
                    "status": "Status",
                    "proxima_acao": "Próxima ação",
                }
            )
            tabela_dividas["Data"] = tabela_dividas["Data"].map(data_br)
            tabela_dividas["Vencimento"] = tabela_dividas["Vencimento"].map(data_br)
            for coluna_valor in ["Saldo original", "Desconto", "Saldo negociado", "Parcela possível"]:
                tabela_dividas[coluna_valor] = tabela_dividas[coluna_valor].map(brl)

            st.markdown("#### Tabela de dívidas")
            st.dataframe(
                tabela_dividas,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Data": st.column_config.TextColumn("Data"),
                    "Vencimento": st.column_config.TextColumn("Vencimento"),
                },
            )

        with col_status:
            status_dividas = (
                df_dividas_view.groupby("status", as_index=False)["saldo_base"]
                .sum()
                .sort_values("saldo_base", ascending=False)
            )
            fig_dividas = px.bar(
                status_dividas,
                x="status",
                y="saldo_base",
                color="status",
                title="Dívidas por status",
                color_discrete_map={
                    "Mapear": "#17486f",
                    "Negociar": "#cc8a2f",
                    "Acordada": "#28c7b7",
                    "Em pagamento": "#0d906f",
                    "Quitada": "#7b8da3",
                },
                labels={"status": "Status", "saldo_base": "Saldo"},
            )
            fig_dividas.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig_dividas), use_container_width=True)

        st.markdown("#### Anotações e próximas ações")
        for _, divida in df_dividas_view.iterrows():
            saldo_base = divida["saldo_base"]
            credor_item = escape(str(divida["credor"] or "Credor não informado"))
            tipo_item = escape(str(divida["tipo"] or "Tipo não informado"))
            status_item = escape(str(divida["status"] or "Sem status"))
            prioridade_item = escape(str(divida["prioridade"] or "Sem prioridade"))
            proxima_item = escape(str(divida["proxima_acao"] or "Sem próxima ação"))
            anotacoes_item = escape(str(divida["anotacoes"] or "Sem anotações"))
            vencimento_item = escape(data_br(divida["vencimento"]))

            st.markdown(
                f"""<div class="debt-item"><div><div class="debt-title">{credor_item}</div><div class="debt-meta">{tipo_item} • Vencimento: {vencimento_item} • Status: {status_item} • Prioridade: {prioridade_item}</div><div class="debt-note"><strong>Próxima ação:</strong> {proxima_item}<br><strong>Anotações:</strong> {anotacoes_item}</div></div><div><div class="debt-value">{brl(saldo_base)}</div><div class="debt-return">Parcela possível: {brl(divida["parcela_possivel"])}</div></div></div>""",
                unsafe_allow_html=True,
            )

            if st.button("🗑️ Apagar dívida", key=f"del_divida{divida['id']}"):
                excluir_divida(divida["id"])
                salvar_divida(data_div, credor, tipo_divida, saldo_original, desconto, saldo_final, parcela, vencimento, prioridade, "", status, proxima_acao, anotacoes)
                st.success("Divida salva.")
                st.rerun()
    if len(dividas):
        df_div = dividas.copy()
        df_div["saldo_base"] = df_div["saldo_negociado"].where(df_div["saldo_negociado"] > 0, df_div["saldo_original"])
        col_d1, col_d2, col_d3 = st.columns(3)
        col_d1.metric("Total aberto", brl(df_div[df_div["status"] != "Quitada"]["saldo_base"].sum()))
        col_d2.metric("Parcelas", brl(df_div[df_div["status"] != "Quitada"]["parcela_possivel"].sum()))
        col_d3.metric("Economia prevista", brl((df_div["saldo_original"] - df_div["saldo_base"]).clip(lower=0).sum()))
        tabela = df_div[["data", "credor", "tipo", "saldo_original", "saldo_base", "parcela_possivel", "vencimento", "prioridade", "status", "proxima_acao"]].copy()
        tabela["data"] = tabela["data"].map(data_br)
        tabela["vencimento"] = tabela["vencimento"].map(data_br)
        for coluna in ["saldo_original", "saldo_base", "parcela_possivel"]:
            tabela[coluna] = tabela[coluna].map(brl)
        st.dataframe(tabela, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma dívida cadastrada ainda.")
        st.info("Nenhuma divida cadastrada.")

# ====================== ABA 5 ======================
with aba[4]:
    st.subheader("Histórico")
    relatorio_pdf = gerar_relatorio_pdf(df, df_investimentos, df_dividas, df_metas)
    if st.session_state.get("historico_limpo"):
        st.success(st.session_state.pop("historico_limpo"))

    st.subheader("Historico")
    st.markdown(
        f"""<div class="history-summary"><strong>Relatório financeiro detalhado</strong><br>Baixe um PDF com o resumo do período, todas as movimentações, investimentos, dívidas e metas cadastradas. Registros incluídos: {len(df)} movimentações, {len(df_investimentos)} investimentos, {len(df_dividas)} dívidas e {len(df_metas)} metas.</div>""",
        f"""<div class="history-summary"><strong>Relatorio detalhado</strong><br>Registros incluidos: {len(df)} movimentacoes, {len(investimentos)} investimentos e {len(dividas)} dividas.</div>""",
        unsafe_allow_html=True,
    )
    st.download_button(
        "⬇️ Baixar relatório detalhado em PDF",
        data=relatorio_pdf,
        "Baixar relatorio em PDF",
        data=gerar_pdf(df, investimentos, dividas),
        file_name=f"relatorio-financeiro-{date.today().strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
    )

    if len(df) > 0:
        with st.expander("🧹 Limpar histórico completo", expanded=False):
            st.caption(
                "Apaga todas as movimentações cadastradas e importadas no histórico. "
                "Investimentos e dívidas não serão alterados."
            )
            confirmar_limpeza = st.checkbox(
                "Confirmo que quero apagar todas as movimentações do histórico",
                key="confirmar_limpeza_historico",
            )
            if st.button("Apagar todo o histórico", disabled=not confirmar_limpeza):
                try:
                    total_apagado = len(df)
                    limpar_historico_movimentacoes()
                    st.session_state["historico_limpo"] = (
                        f"Histórico limpo com sucesso. {total_apagado} movimentações foram apagadas."
                    )
                    st.rerun()
                except Exception as erro:
                    st.error(f"Não foi possível limpar o histórico: {mensagem_erro_usuario(erro)}")

        st.markdown("#### Histórico consolidado")
        st.caption(
            "Todos os registros manuais e importados por planilha ficam disponíveis nesta mesma visão."
    if len(df):
        with st.expander("Limpar historico completo", expanded=False):
            confirmar = st.checkbox("Confirmo que quero apagar todas as movimentacoes")
            if st.button("Apagar todo o historico", disabled=not confirmar):
                limpar_historico()
                st.success("Historico limpo.")
                st.rerun()
        col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
        busca = col_f1.text_input("Buscar", placeholder="Descricao, categoria ou pagamento")
        tipo_filtro = col_f2.selectbox("Tipo", ["Todos", "Entrada", "Saida"])
        categorias = sorted(df["categoria"].replace("", "Sem categoria").unique())
        categoria_filtro = col_f3.selectbox("Categoria", ["Todas", *categorias])
        df_hist = df.copy()
        if busca.strip():
            termo = normalizar(busca)
            df_hist = df_hist[
                df_hist.apply(
                    lambda row: termo in normalizar(" ".join([row["descricao"], row["categoria"], row["cartao"]])),
                    axis=1,
                )
            ]
        if tipo_filtro != "Todos":
            df_hist = df_hist[df_hist["tipo"] == tipo_filtro]
        if categoria_filtro != "Todas":
            df_hist = df_hist[df_hist["categoria"] == categoria_filtro]
        tabela = df_hist[["data", "descricao", "categoria", "tipo", "cartao", "valor"]].rename(
            columns={"data": "Data", "descricao": "Descricao", "categoria": "Categoria", "tipo": "Tipo", "cartao": "Pagamento", "valor": "Valor"}
        )

        col_busca, col_tipo, col_categoria = st.columns([2, 1, 1])
        with col_busca:
            busca_historico = st.text_input(
                "Buscar movimentação",
                placeholder="Descrição, categoria ou pagamento",
                key="busca_historico",
            )
        with col_tipo:
            tipo_historico = st.selectbox(
                "Tipo",
                ["Todos", "Entrada", "Saída"],
                key="tipo_historico",
            )
        with col_categoria:
            categorias_historico = sorted(
                texto_planilha(categoria, "Sem categoria")
                for categoria in df["categoria"].dropna().unique()
            )
            categoria_historico_filtro = st.selectbox(
                "Categoria",
                ["Todas", *categorias_historico],
                key="categoria_historico",
            )

        df_historico = df.copy()
        if busca_historico.strip():
            termo = normalizar_coluna(busca_historico)
            mascara = df_historico.apply(
                lambda registro: termo
                in normalizar_coluna(
                    " ".join(
                        [
                            texto_planilha(registro.get("descricao")),
                            texto_planilha(registro.get("categoria")),
                            texto_planilha(registro.get("cartao")),
                        ]
                    )
                ),
                axis=1,
            )
            df_historico = df_historico[mascara]
        if tipo_historico != "Todos":
            df_historico = df_historico[df_historico["tipo"] == tipo_historico]
        if categoria_historico_filtro != "Todas":
            df_historico = df_historico[df_historico["categoria"] == categoria_historico_filtro]

        tabela_historico = df_historico[
            ["data", "descricao", "categoria", "tipo", "cartao", "valor"]
        ].rename(
            columns={
                "data": "Data",
                "descricao": "Descrição",
                "categoria": "Categoria",
                "tipo": "Tipo",
                "cartao": "Forma de Pagamento",
                "valor": "Valor",
            }
        )
        tabela_historico["Data"] = tabela_historico["Data"].map(data_br)
        tabela_historico["Valor"] = tabela_historico["Valor"].map(brl)
        st.dataframe(
            tabela_historico,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Data": st.column_config.TextColumn("Data"),
                "Valor": st.column_config.TextColumn("Valor"),
            },
        )
        st.caption(f"{len(df_historico)} de {len(df)} movimentações exibidas.")

        st.markdown("#### Detalhes e exclusão")
        for _, row in df_historico.iterrows():
        tabela["Data"] = tabela["Data"].map(data_br)
        tabela["Valor"] = tabela["Valor"].map(brl)
        st.dataframe(tabela, use_container_width=True, hide_index=True)
        st.caption(f"{len(df_hist)} de {len(df)} movimentacoes exibidas.")
        for _, row in df_hist.iterrows():
            classe = "positive" if row["valor"] >= 0 else "negative"
            descricao_historico = escape(str(row["descricao"] or "Sem descrição"))
            categoria_historico = escape(str(row["categoria"] or "Sem categoria"))
            data_historico = escape(data_br(row["data"]))
            cartao_historico = escape(str(row["cartao"] or "Sem forma de pagamento"))

            st.markdown(
                f"""<div class="history-item"><div><div class="history-title">{descricao_historico}</div><div class="history-meta">{categoria_historico} • {data_historico} • {cartao_historico}</div></div><div class="{classe}">{brl(row["valor"])}</div></div>""",
                f"""<div class="history-card"><strong>{escape(row["descricao"])}</strong><br><span>{escape(row["categoria"])} - {data_br(row["data"])} - {escape(row["cartao"])}</span><div class="{classe}">{brl(row["valor"])}</div></div>""",
                unsafe_allow_html=True,
            )

            if st.button("🗑️ Apagar", key=f"del{row['id']}"):
            if st.button("Apagar", key=f"del{row['id']}"):
                excluir_transacao(row["id"])
                st.rerun()
        if df_historico.empty:
            st.info("Nenhuma movimentação corresponde aos filtros selecionados.")
    else:
        st.info("Nenhum registro ainda.")

st.caption("Dashboard Financeiro • Visão financeira clara • Experiência premium")
st.caption("Dashboard Financeiro - Visao financeira clara - Experiencia premium")
