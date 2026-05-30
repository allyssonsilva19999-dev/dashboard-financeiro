import base64
import sqlite3
import textwrap
from datetime import date
from html import escape
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="CreativeStyle Finance • 2025",
    layout="wide",
    page_icon="🌄",
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

    @media (max-width: 900px) {{
        .hero,
        .metric-grid,
        .investment-grid,
        .investment-item,
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
</style>
""",
    unsafe_allow_html=True,
)

# ====================== BANCO ======================
DB_FILE = "financeiro.db"


def brl(valor):
    texto = f"R$ {valor:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def data_br(valor):
    data_convertida = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data_convertida):
        return "Sem data"
    return data_convertida.strftime("%d/%m/%Y")


def _pdf_texto(texto):
    return (
        str(texto)
        .replace("\\", "\\\\")
        .replace("(", "\\(")
        .replace(")", "\\)")
        .encode("latin-1", errors="replace")
        .decode("latin-1")
    )


def gerar_relatorio_pdf(df_transacoes, df_investimentos):
    linhas = []

    def adicionar(texto="", estilo="corpo", largura=92):
        partes = textwrap.wrap(str(texto), width=largura) or [""]
        for parte in partes:
            linhas.append((parte, estilo))

    entradas = df_transacoes[df_transacoes["valor"] > 0]["valor"].sum() if len(df_transacoes) else 0
    saidas = abs(df_transacoes[df_transacoes["valor"] < 0]["valor"].sum()) if len(df_transacoes) else 0
    saldo = df_transacoes["valor"].sum() if len(df_transacoes) else 0
    investimentos = df_investimentos["valor"].sum() if len(df_investimentos) else 0

    adicionar("CreativeStyle Finance - Relatório Financeiro Detalhado", "titulo", 62)
    adicionar(f"Emitido em {date.today().strftime('%d/%m/%Y')}", "pequeno")
    adicionar()
    adicionar("Resumo financeiro", "secao")
    adicionar(f"Entradas totais: {brl(entradas)}")
    adicionar(f"Saídas totais: {brl(saidas)}")
    adicionar(f"Saldo atual: {brl(saldo)}")
    adicionar(f"Patrimônio investido: {brl(investimentos)}")
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
        """
    )
    conn.commit()
    conn.close()


def carregar_dados():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query(
        "SELECT * FROM transacoes ORDER BY data DESC, id DESC",
        conn,
    )
    conn.close()
    return df


def excluir_transacao(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def carregar_investimentos():
    conn = sqlite3.connect(DB_FILE)
    df_investimentos = pd.read_sql_query(
        "SELECT * FROM investimentos ORDER BY data DESC, id DESC",
        conn,
    )
    conn.close()
    return df_investimentos


def excluir_investimento(iid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM investimentos WHERE id = ?", (iid,))
    conn.commit()
    conn.close()


def style_plot(fig):
    fig.update_layout(
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(color="#1f1d1a"),
        title=dict(font=dict(size=20, color="#1f1d1a")),
        legend=dict(bgcolor="rgba(255,255,255,0)", font=dict(color="#29251f")),
        margin=dict(l=20, r=20, t=60, b=20),
    )
    fig.update_xaxes(
        gridcolor="rgba(31,29,26,0.10)",
        zerolinecolor="rgba(31,29,26,0.12)",
    )
    fig.update_yaxes(
        gridcolor="rgba(31,29,26,0.10)",
        zerolinecolor="rgba(31,29,26,0.12)",
    )
    return fig


init_db()
df = carregar_dados()
df_investimentos = carregar_investimentos()
hero_total_entradas = df[df["valor"] > 0]["valor"].sum() if len(df) > 0 else 0
hero_total_saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) > 0 else 0
hero_saldo = df["valor"].sum() if len(df) > 0 else 0
hero_total_investido = df_investimentos["valor"].sum() if len(df_investimentos) > 0 else 0

# ====================== HERO ======================
st.markdown(
    f"""<section class="hero"><div class="hero-card"><div class="hero-top"><span>Creativestyle_</span><span class="pill">Finance 2025</span></div><h1>CreativeStyle<br>Finance</h1><p class="hero-subtitle">Organize decisões, acompanhe seu patrimônio e transforme pequenas escolhas financeiras em progresso consistente.</p><blockquote class="hero-quote">“Preço é o que você paga; valor é o que você recebe.”<cite>Benjamin Graham</cite></blockquote></div><aside class="utility-card"><div class="utility-head"><strong>Visão geral</strong><span>Atualizado agora</span></div><div class="utility-grid"><div class="utility-item"><div class="utility-label">Saldo atual</div><div class="utility-value">{brl(hero_saldo)}</div></div><div class="utility-item"><div class="utility-label">Entradas</div><div class="utility-value">{brl(hero_total_entradas)}</div></div><div class="utility-item"><div class="utility-label">Saídas</div><div class="utility-value">{brl(hero_total_saidas)}</div></div><div class="utility-item"><div class="utility-label">Investimentos</div><div class="utility-value">{brl(hero_total_investido)}</div></div></div></aside></section>""",
    unsafe_allow_html=True,
)

# ====================== NAVEGAÇÃO ======================
aba = st.tabs(["➕ Nova Movimentação", "📊 Dashboard", "💼 Investimentos", "📋 Histórico"])

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
            valor_final = valor if tipo_limpo == "Entrada" else -abs(valor)

            conn = sqlite3.connect(DB_FILE)
            conn.execute(
                """
                INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (data, descricao, categoria, valor_final, tipo_limpo, cartao),
            )
            conn.commit()
            conn.close()

            st.success("Movimentação salva com sucesso!")
            st.rerun()

# ====================== ABA 2 ======================
with aba[1]:
    st.subheader("Dashboard em Tempo Real")

    total_entradas = df[df["valor"] > 0]["valor"].sum() if len(df) > 0 else 0
    total_saidas = abs(df[df["valor"] < 0]["valor"].sum()) if len(df) > 0 else 0
    saldo = df["valor"].sum() if len(df) > 0 else 0

    st.markdown(
        f"""<div class="metric-grid"><div class="metric-card"><div class="metric-label">Entradas</div><div class="metric-value">{brl(total_entradas)}</div><div class="metric-foot">Receitas registradas</div></div><div class="metric-card"><div class="metric-label">Saídas</div><div class="metric-value">{brl(total_saidas)}</div><div class="metric-foot">Despesas acumuladas</div></div><div class="metric-card"><div class="metric-label">Saldo</div><div class="metric-value">{brl(saldo)}</div><div class="metric-foot">Resultado atual</div></div><div class="metric-card"><div class="metric-label">Registros</div><div class="metric-value">{len(df)}</div><div class="metric-foot">Movimentações salvas</div></div></div>""",
        unsafe_allow_html=True,
    )

    if len(df) > 0:
        df_chart = df.copy()
        df_chart["valor_abs"] = df_chart["valor"].abs()

        categoria_total = (
            df_chart.groupby("categoria", as_index=False)["valor_abs"]
            .sum()
            .sort_values("valor_abs", ascending=False)
        )

        fluxo_total = df_chart.groupby(["categoria", "tipo"], as_index=False)["valor_abs"].sum()

        col_g1, col_g2 = st.columns(2)

        with col_g1:
            fig = px.pie(
                categoria_total,
                names="categoria",
                values="valor_abs",
                title="Distribuição por Categoria",
                hole=0.58,
                color_discrete_sequence=[
                    "#1f1d1a",
                    "#d99b45",
                    "#f3dcc1",
                    "#6f4d2d",
                    "#b98244",
                    "#efe4d5",
                    "#3d6d45",
                ],
            )
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(style_plot(fig), use_container_width=True)

        with col_g2:
            fig2 = px.bar(
                fluxo_total,
                x="categoria",
                y="valor_abs",
                color="tipo",
                title="Entradas x Saídas",
                color_discrete_map={
                    "Entrada": "#3d6d45",
                    "Saída": "#9c4f2f",
                },
            )
            st.plotly_chart(style_plot(fig2), use_container_width=True)
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
            conn = sqlite3.connect(DB_FILE)
            conn.execute(
                """
                INSERT INTO investimentos
                    (data, tipo, valor, rentabilidade, descricao, status)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    data_investimento,
                    tipo_investimento,
                    valor_investimento,
                    rentabilidade,
                    descricao_investimento,
                    status_investimento,
                ),
            )
            conn.commit()
            conn.close()

            st.success("Investimento salvo com sucesso!")
            st.rerun()

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
                    "#1f1d1a",
                    "#d99b45",
                    "#f3dcc1",
                    "#6f4d2d",
                    "#b98244",
                    "#3d6d45",
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
                    "Ativo": "#3d6d45",
                    "Planejado": "#d99b45",
                    "Resgatado": "#6f4d2d",
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
    st.subheader("Histórico")
    relatorio_pdf = gerar_relatorio_pdf(df, df_investimentos)

    st.markdown(
        f"""<div class="history-summary"><strong>Relatório financeiro detalhado</strong><br>Baixe um PDF com o resumo do período, todas as movimentações e os investimentos cadastrados. Registros incluídos: {len(df)} movimentações e {len(df_investimentos)} investimentos.</div>""",
        unsafe_allow_html=True,
    )
    st.download_button(
        "⬇️ Baixar relatório detalhado em PDF",
        data=relatorio_pdf,
        file_name=f"relatorio-financeiro-{date.today().strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
    )

    if len(df) > 0:
        for _, row in df.iterrows():
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
    else:
        st.info("Nenhum registro ainda.")

st.caption("CreativeStyle Finance • Visão financeira clara • Vidro fosco premium")
