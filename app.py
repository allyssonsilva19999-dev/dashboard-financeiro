import base64
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
)

# ====================== IMAGEM DE FUNDO ======================
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

    @media (max-width: 900px) {{
        .hero,
        .metric-grid,
        .investment-grid,
        .debt-grid,
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
    .debt-return {{
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
    .debt-title {{
        color: var(--navy);
    }}

    .history-item,
    .investment-item,
    .debt-item {{
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
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def data_br(valor):
    data_convertida = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data_convertida):
        return "Sem data"
    return data_convertida.strftime("%d/%m/%Y")


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

    if pd.isna(data_convertida):
        hoje = date.today()
        ano = hoje.year
        return date(ano, numero_mes, 1).isoformat()

    ano = data_convertida.year if data_convertida.year > 1970 else date.today().year
    ultimo_dia = pd.Period(year=ano, month=numero_mes, freq="M").days_in_month
    dia = min(data_convertida.day, ultimo_dia)
    return date(ano, numero_mes, dia).isoformat()


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
        return "Saída"
    if "entradas" in texto_contexto or "receitas" in texto_contexto:
        return "Entrada"
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

    linhas.append(
        {
            "data": data_final,
            "descricao": descricao,
            "categoria": categoria,
            "valor": valor_final,
            "tipo": tipo,
            "cartao": pagamento,
        }
    )
    return True


def adicionar_blocos_de_entrada(linhas, df_mes, numero_mes):
    adicionadas = 0
    for linha_bloco in range(len(df_mes.index)):
        for coluna_bloco in range(len(df_mes.columns)):
            if normalizar_coluna(celula_planilha(df_mes, linha_bloco, coluna_bloco)) != "entradas":
                continue

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
                    )
                    if coluna is not None:
                        colunas[chave] = coluna
                contexto = texto_contexto_planilha(df_mes, numero_linha, coluna_nome, coluna_valor)
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
                    continue
                linhas_vazias = 0

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
            return pd.DataFrame(), 0
        primeira_aba = primeira_aba.copy()
        primeira_aba.columns = primeira_aba.iloc[0]
        df_planilha = primeira_aba.iloc[1:].reset_index(drop=True)

    aliases = {
        "data": ["data", "dt", "dia", "date"],
        "descricao": ["descricao", "descrição", "historico", "histórico", "lancamento", "lançamento", "detalhe"],
        "categoria": ["categoria", "grupo", "classificacao", "classificação"],
        "valor": ["valor", "valor r$", "valor rs", "amount", "preco", "preço"],
        "tipo": ["tipo", "natureza", "entrada saida", "receita despesa"],
        "cartao": ["cartao", "cartão", "forma de pagamento", "pagamento", "conta", "banco"],
    }

    colunas_normais = {normalizar_coluna(coluna): coluna for coluna in df_planilha.columns}
    mapa_colunas = {}

    for destino, opcoes in aliases.items():
        for opcao in opcoes:
            chave = normalizar_coluna(opcao)
            if chave in colunas_normais:
                mapa_colunas[destino] = colunas_normais[chave]
                break

    if "valor" not in mapa_colunas:
        raise ValueError(
            "Não encontrei uma coluna de valor nem abas mensais no formato da Organização Financeira."
        )

    linhas = []
    ignoradas = 0

    for _, linha in df_planilha.iterrows():
        valor_original = converter_valor(linha.get(mapa_colunas["valor"]))
        if valor_original is None or valor_original == 0:
            ignoradas += 1
            continue

        data_original = linha.get(mapa_colunas.get("data"), date.today())
        data_convertida = pd.to_datetime(data_original, errors="coerce", dayfirst=True)
        if pd.isna(data_convertida):
            data_convertida = pd.Timestamp(date.today())

        tipo_original = str(linha.get(mapa_colunas.get("tipo"), "")).strip().lower()
        if any(palavra in tipo_original for palavra in ["saida", "saída", "despesa", "debito", "débito", "gasto"]):
            tipo = "Saída"
        elif any(palavra in tipo_original for palavra in ["entrada", "receita", "credito", "crédito", "salario", "salário"]):
            tipo = "Entrada"
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


def conciliar_movimentacoes(df_importado, df_existente):
    if df_importado.empty:
        return df_importado.copy(), 0

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


def gerar_relatorio_pdf(df_transacoes, df_investimentos, df_dividas=None):
    if df_dividas is None:
        df_dividas = pd.DataFrame()

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

    adicionar("Dashboard Financeiro - Relatório Financeiro Detalhado", "titulo", 62)
    adicionar(f"Emitido em {date.today().strftime('%d/%m/%Y')}", "pequeno")
    adicionar()
    adicionar("Resumo financeiro", "secao")
    adicionar(f"Entradas totais: {brl(entradas)}")
    adicionar(f"Saídas totais: {brl(saidas)}")
    adicionar(f"Saldo atual: {brl(saldo)}")
    adicionar(f"Patrimônio investido: {brl(investimentos)}")
    adicionar(f"Dívidas monitoradas: {brl(total_dividas)}")
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


def carregar_dados():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query(
        "SELECT * FROM transacoes ORDER BY data DESC, id DESC",
        conn,
    )
    conn.close()
    return normalizar_dataframe_financeiro(
        df,
        ["descricao", "categoria", "tipo", "cartao"],
    )


def carregar_investimentos():
    conn = sqlite3.connect(DB_FILE)
    df_investimentos = pd.read_sql_query(
        "SELECT * FROM investimentos ORDER BY data DESC, id DESC",
        conn,
    )
    conn.close()
    return normalizar_dataframe_financeiro(
        df_investimentos,
        ["tipo", "rentabilidade", "descricao", "status"],
    )


def carregar_dividas():
    conn = sqlite3.connect(DB_FILE)
    df_dividas = pd.read_sql_query(
        "SELECT * FROM dividas ORDER BY data DESC, id DESC",
        conn,
    )
    conn.close()
    return normalizar_dataframe_dividas(df_dividas)


def excluir_transacao(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def limpar_historico_movimentacoes():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes")
    conn.commit()
    conn.close()


def excluir_investimento(iid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM investimentos WHERE id = ?", (iid,))
    conn.commit()
    conn.close()


def excluir_divida(did):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM dividas WHERE id = ?", (did,))
    conn.commit()
    conn.close()


def salvar_transacao(data_movimentacao, descricao, categoria, valor, tipo, cartao):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        """
        INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (data_movimentacao, descricao, categoria, valor, tipo, cartao),
    )
    conn.commit()
    conn.close()


def salvar_investimento(data_investimento, tipo, valor, rentabilidade, descricao, status):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        """
        INSERT INTO investimentos
            (data, tipo, valor, rentabilidade, descricao, status)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (data_investimento, tipo, valor, rentabilidade, descricao, status),
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
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        """
        INSERT INTO dividas
            (
                data, credor, tipo, saldo_original, desconto, saldo_negociado,
                parcela_possivel, vencimento, prioridade, consequencia,
                status, proxima_acao, anotacoes
            )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
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
        ),
    )
    conn.commit()
    conn.close()


def style_plot(fig):
    fig.update_layout(
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0.62)",
        font=dict(color="#071426", size=12),
        title=dict(font=dict(size=19, color="#081b33"), x=0.04, xanchor="left"),
        legend=dict(
            bgcolor="rgba(255,255,255,0)",
            font=dict(color="#071426"),
            orientation="h",
            yanchor="bottom",
            y=1.01,
            xanchor="right",
            x=1,
        ),
        hoverlabel=dict(bgcolor="#ffffff", font_color="#071426"),
        margin=dict(l=28, r=24, t=78, b=34),
        height=410,
        separators=",.",
    )
    fig.update_xaxes(
        gridcolor="rgba(8,27,51,0.08)",
        linecolor="rgba(8,27,51,0.10)",
        zerolinecolor="rgba(8,27,51,0.10)",
        title_font=dict(color="#17486f"),
        tickfont=dict(color="#17486f"),
    )
    fig.update_yaxes(
        gridcolor="rgba(8,27,51,0.08)",
        linecolor="rgba(8,27,51,0.10)",
        zerolinecolor="rgba(8,27,51,0.10)",
        title_font=dict(color="#17486f"),
        tickfont=dict(color="#17486f"),
    )
    return fig


# Acesso direto: o usuário entra no dashboard sem etapa de login.
init_db()
try:
    df = carregar_dados()
    df_investimentos = carregar_investimentos()
    df_dividas = carregar_dividas()
except Exception as erro:
    st.error(f"Não foi possível carregar seus dados: {mensagem_erro_usuario(erro)}")
    st.info("Atualize a página e tente novamente.")
    st.stop()

hero_total_entradas = df[df["valor"] > 0]["valor"].sum() if len(df) > 0 else 0
hero_total_saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) > 0 else 0
hero_saldo = df["valor"].sum() if len(df) > 0 else 0
hero_total_investido = df_investimentos["valor"].sum() if len(df_investimentos) > 0 else 0

# ====================== HERO ======================
st.markdown(
    f"""<section class="site-nav"><div class="brand-lockup"><span class="brand-mark">DF</span><span>Dashboard Financeiro</span></div><div class="nav-links"><span>Organização</span><span>Dashboard</span><span>Dívidas</span><span>Histórico</span></div><span class="nav-cta">Controle financeiro</span></section><section class="hero"><div class="hero-card"><div class="hero-top"><span class="eyebrow">Painel inteligente</span><span class="pill">Financeiro 2025</span></div><h1>Dashboard<br>Financeiro</h1><p class="hero-subtitle">Organize decisões, acompanhe seu patrimônio e transforme pequenas escolhas financeiras em progresso consistente.</p><blockquote class="hero-quote">“Preço é o que você paga; valor é o que você recebe.”<cite>Benjamin Graham</cite></blockquote></div><aside class="utility-card"><div class="utility-head"><strong>Visão geral</strong><span>Atualizado agora</span></div><div class="utility-grid"><div class="utility-item"><div class="utility-label">Saldo atual</div><div class="utility-value">{brl(hero_saldo)}</div></div><div class="utility-item"><div class="utility-label">Entradas</div><div class="utility-value">{brl(hero_total_entradas)}</div></div><div class="utility-item"><div class="utility-label">Saídas</div><div class="utility-value">{brl(hero_total_saidas)}</div></div><div class="utility-item"><div class="utility-label">Investimentos</div><div class="utility-value">{brl(hero_total_investido)}</div></div></div></aside></section>""",
    unsafe_allow_html=True,
)

# ====================== NAVEGAÇÃO ======================
aba = st.tabs(["➕ Nova Movimentação", "📊 Dashboard", "💼 Investimentos", "🤝 Dívidas", "📋 Histórico"])

# ====================== ABA 1 ======================
with aba[0]:
    st.subheader("Adicionar Nova Movimentação")

    tipo = st.radio("Tipo", ["💰 Entrada", "💸 Saída"], horizontal=True)
    tipo_limpo = "Entrada" if "Entrada" in tipo else "Saída"

    with st.form("nova"):
        col1, col2 = st.columns(2)

        with col1:
            data = st.date_input("Data", value=date.today(), format="DD/MM/YYYY")
            descricao = st.text_input("Descrição")

            if tipo_limpo == "Entrada":
                categoria = st.selectbox(
                    "Categoria",
                    ["Salário", "Rendas Extras", "Freelance", "Reembolso"],
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

# ====================== ABA 2 ======================
with aba[1]:
    st.subheader("Dashboard em Tempo Real")
    if st.session_state.get("resultado_importacao"):
        st.success(st.session_state.pop("resultado_importacao"))

    total_entradas = df[df["valor"] > 0]["valor"].sum() if len(df) > 0 else 0
    total_saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) > 0 else 0
    saldo = df["valor"].sum() if len(df) > 0 else 0

    st.markdown(
        f"""<div class="metric-grid"><div class="metric-card"><div class="metric-label">Entradas</div><div class="metric-value">{brl(total_entradas)}</div><div class="metric-foot">Receitas registradas</div></div><div class="metric-card"><div class="metric-label">Saídas</div><div class="metric-value">{brl(total_saidas)}</div><div class="metric-foot">Despesas acumuladas</div></div><div class="metric-card"><div class="metric-label">Saldo</div><div class="metric-value">{brl(saldo)}</div><div class="metric-foot">Resultado atual</div></div><div class="metric-card"><div class="metric-label">Registros</div><div class="metric-value">{len(df)}</div><div class="metric-foot">Movimentações salvas</div></div></div>""",
        unsafe_allow_html=True,
    )

    with st.expander("📤 Subir planilha de movimentações", expanded=False):
        st.caption(
            "Os registros novos serão integrados ao mesmo histórico, saldo e gráficos dos "
            "cadastros manuais. Lançamentos já existentes não serão duplicados."
        )

        st.download_button(
            "Baixar modelo Excel com tabela de valores",
            data=gerar_modelo_excel(),
            file_name="modelo-dashboard-financeiro.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        arquivo_planilha = st.file_uploader(
            "Escolha a planilha",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=False,
        )

        if arquivo_planilha is not None:
            try:
                df_planilha = ler_planilha_movimentacoes(arquivo_planilha)
                df_importado, linhas_ignoradas = preparar_movimentacoes_importadas(df_planilha)

                if df_importado.empty:
                    st.warning("Não encontrei movimentações válidas nessa planilha.")
                else:
                    df_novo, duplicadas_encontradas = conciliar_movimentacoes(df_importado, df)
                    entradas_importadas = df_novo[df_novo["valor"] > 0]["valor"].sum()
                    saidas_importadas = abs(df_novo[df_novo["valor"] < 0]["valor"].sum())
                    saldo_importado = df_novo["valor"].sum()

                    col_valor1, col_valor2, col_valor3, col_valor4 = st.columns(4)
                    col_valor1.metric("Novas entradas", brl(entradas_importadas))
                    col_valor2.metric("Novas saídas", brl(saidas_importadas))
                    col_valor3.metric("Impacto no saldo", brl(saldo_importado))
                    col_valor4.metric("Novos registros", len(df_novo))

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
            except Exception as erro:
                st.error(f"Não consegui importar essa planilha: {mensagem_erro_usuario(erro)}")

    if len(df) > 0:
        df_chart = df.copy()
        df_chart["valor_abs"] = df_chart["valor"].abs()
        df_chart["data_convertida"] = pd.to_datetime(df_chart["data"], errors="coerce")

        categoria_total = (
            df_chart.groupby("categoria", as_index=False)["valor_abs"]
            .sum()
            .sort_values("valor_abs", ascending=False)
        )

        fluxo_total = df_chart.groupby(["categoria", "tipo"], as_index=False)["valor_abs"].sum()
        fluxo_mensal = (
            df_chart.dropna(subset=["data_convertida"])
            .assign(
                mes=lambda dados: dados["data_convertida"].dt.to_period("M").dt.to_timestamp(),
                entradas=lambda dados: dados["valor"].clip(lower=0),
                saidas=lambda dados: dados["valor"].clip(upper=0).abs(),
            )
            .groupby("mes", as_index=False)[["entradas", "saidas"]]
            .sum()
        )

        st.markdown(
            f"""<div class="chart-intro"><strong>Visão consolidada</strong><br>
            Os {len(df)} lançamentos cadastrados manualmente e integrados por planilha alimentam
            automaticamente os gráficos e o histórico abaixo.</div>""",
            unsafe_allow_html=True,
        )

        col_g1, col_g2 = st.columns(2)

        with col_g1:
            fig = px.pie(
                categoria_total,
                names="categoria",
                values="valor_abs",
                title="Distribuição por Categoria",
                hole=0.58,
                color_discrete_sequence=[
                    "#081b33",
                    "#28c7b7",
                    "#17486f",
                    "#8be6d4",
                    "#3a6388",
                    "#0d906f",
                    "#78a8c8",
                ],
            )
            fig.update_traces(
                textposition="outside",
                textinfo="percent+label",
                insidetextorientation="radial",
                marker=dict(line=dict(color="rgba(255,255,255,0.86)", width=2)),
                pull=[0.018] * len(categoria_total),
                hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>",
            )
            fig = style_plot(fig)
            fig.update_layout(
                height=500,
                showlegend=True,
                uniformtext_minsize=11,
                uniformtext_mode="hide",
                margin=dict(l=42, r=70, t=86, b=92),
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.12,
                    xanchor="center",
                    x=0.5,
                    bgcolor="rgba(255,255,255,0)",
                    font=dict(color="#071426", size=11),
                ),
            )
            st.plotly_chart(fig, use_container_width=True)

        with col_g2:
            fig2 = px.bar(
                fluxo_total,
                x="categoria",
                y="valor_abs",
                color="tipo",
                title="Entradas x Saídas",
                color_discrete_map={
                    "Entrada": "#0d906f",
                    "Saída": "#cc4a5b",
                },
                labels={"categoria": "Categoria", "valor_abs": "Valor", "tipo": "Tipo"},
            )
            fig2.update_traces(
                marker_line_color="rgba(255,255,255,0.78)",
                marker_line_width=1,
                hovertemplate="<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>",
            )
            fig2.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig2), use_container_width=True)

        if len(fluxo_mensal) > 0:
            fig_mensal = px.bar(
                fluxo_mensal,
                x="mes",
                y=["entradas", "saidas"],
                barmode="group",
                title="Evolução mensal consolidada",
                color_discrete_map={"entradas": "#0d906f", "saidas": "#cc4a5b"},
                labels={"mes": "Mês", "value": "Valor", "variable": "Movimentação"},
            )
            fig_mensal.for_each_trace(
                lambda trace: trace.update(
                    name="Entradas" if trace.name == "entradas" else "Saídas",
                    marker_line_color="rgba(255,255,255,0.76)",
                    marker_line_width=1,
                    hovertemplate="<b>%{x|%m/%Y}</b><br>R$ %{y:,.2f}<extra></extra>",
                )
            )
            fig_mensal.update_xaxes(tickformat="%m/%Y")
            fig_mensal.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig_mensal), use_container_width=True)
    else:
        st.info("Adicione uma movimentação para visualizar os gráficos.")

# ====================== ABA 3 ======================
with aba[2]:
    st.subheader("Investimentos")

    with st.form("novo_investimento"):
        st.markdown("#### Adicionar investimento")
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

        with col3:
            descricao_investimento = st.text_input(
                "Descrição",
                placeholder="Ex.: CDB Banco X - reserva",
            )
            status_investimento = st.selectbox(
                "Status",
                ["Ativo", "Planejado", "Resgatado"],
            )

        if st.form_submit_button("Salvar investimento"):
            if valor_investimento <= 0:
                st.error("Informe um valor de investimento maior que zero.")
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
                st.rerun()
    else:
        st.info("Nenhum investimento cadastrado ainda.")

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
                st.rerun()
    else:
        st.info("Nenhuma dívida cadastrada ainda.")

# ====================== ABA 5 ======================
with aba[4]:
    st.subheader("Histórico")
    relatorio_pdf = gerar_relatorio_pdf(df, df_investimentos, df_dividas)
    if st.session_state.get("historico_limpo"):
        st.success(st.session_state.pop("historico_limpo"))

    st.markdown(
        f"""<div class="history-summary"><strong>Relatório financeiro detalhado</strong><br>Baixe um PDF com o resumo do período, todas as movimentações, investimentos e dívidas cadastradas. Registros incluídos: {len(df)} movimentações, {len(df_investimentos)} investimentos e {len(df_dividas)} dívidas.</div>""",
        unsafe_allow_html=True,
    )
    st.download_button(
        "⬇️ Baixar relatório detalhado em PDF",
        data=relatorio_pdf,
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
            classe = "positive" if row["valor"] >= 0 else "negative"
            descricao_historico = escape(str(row["descricao"] or "Sem descrição"))
            categoria_historico = escape(str(row["categoria"] or "Sem categoria"))
            data_historico = escape(data_br(row["data"]))
            cartao_historico = escape(str(row["cartao"] or "Sem forma de pagamento"))

            st.markdown(
                f"""<div class="history-item"><div><div class="history-title">{descricao_historico}</div><div class="history-meta">{categoria_historico} • {data_historico} • {cartao_historico}</div></div><div class="{classe}">{brl(row["valor"])}</div></div>""",
                unsafe_allow_html=True,
            )

            if st.button("🗑️ Apagar", key=f"del{row['id']}"):
                excluir_transacao(row["id"])
                st.rerun()
        if df_historico.empty:
            st.info("Nenhuma movimentação corresponde aos filtros selecionados.")
    else:
        st.info("Nenhum registro ainda.")

st.caption("Dashboard Financeiro • Visão financeira clara • Experiência premium")
