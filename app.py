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

# ====================== CONFIG ======================
st.set_page_config(
    page_title="Dashboard Financeiro",
    layout="wide",
    page_icon="📊",
)

DB_FILE = "financeiro.db"
BACKGROUND_IMAGE = Path(__file__).parent / "assets" / "background-blue-dunes.jpg"

# ====================== UTILITÁRIOS ======================
def image_to_base64(path: Path) -> str:
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode()


def brl(valor) -> str:
    try:
        numero = float(valor or 0)
    except (TypeError, ValueError):
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
    data_convertida = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.isna(data_convertida):
        return "Sem data"
    return data_convertida.strftime("%d/%m/%Y")


def limpar_texto(valor, padrao: str = "") -> str:
    if pd.isna(valor):
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


# ====================== ESTILO ======================
bg_base64 = image_to_base64(BACKGROUND_IMAGE)
bg_css = (
    f"url('data:image/jpeg;base64,{bg_base64}') center/cover fixed no-repeat"
    if bg_base64
    else "linear-gradient(145deg, #eef1f5, #f8f9fb)"
)

st.markdown(
    f"""
<style>
    :root {{
        --ink: #111318;
        --muted: #747985;
        --lime: #d9ff00;
        --blue: #8fb1ff;
        --line: rgba(17,19,24,.08);
        --shadow: 0 26px 80px rgba(17,19,24,.12);
        --soft: 0 14px 38px rgba(17,19,24,.08);
        --card: rgba(255,255,255,.92);
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
    .metric-grid, .indicator-grid, .answer-grid, .goal-grid, .debt-grid, .investment-grid {{
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.85rem;
        margin: 1rem 0 1.15rem;
    }}
    .metric-card, .indicator-card, .answer-card, .goal-card, .history-item, .investment-item, .debt-item {{
        position: relative;
        overflow: hidden;
        min-height: 8.5rem;
        padding: 1.1rem;
        border-radius: 24px;
        background: var(--card);
        border: 1px solid rgba(255,255,255,.92);
        box-shadow: var(--soft);
    }}
    .metric-card.accent {{
        background: linear-gradient(135deg, var(--lime), #caff00);
    }}
    .metric-card.blue {{
        background: linear-gradient(135deg, #fff, #e9efff);
    }}
    .metric-card.dark {{
        background: #111318;
        color: white;
    }}
    .metric-card.dark .metric-label,
    .metric-card.dark .metric-foot,
    .metric-card.dark .metric-value {{
        color: white;
    }}
    .metric-label, .indicator-top, .answer-question {{
        color: var(--muted);
        font-size: 0.74rem;
        font-weight: 850;
        text-transform: uppercase;
    }}
    .metric-value, .indicator-value, .answer-value {{
        margin-top: 0.9rem;
        color: var(--ink);
        font-size: clamp(1.35rem, 2.6vw, 2.1rem);
        line-height: 1.05;
        font-weight: 900;
    }}
    .metric-foot, .indicator-note, .answer-action, .goal-meta, .debt-meta, .investment-meta {{
        margin-top: 0.45rem;
        color: var(--muted);
        font-size: 0.84rem;
        line-height: 1.4;
    }}
    .positive {{ color: #0d906f; font-weight: 900; }}
    .negative {{ color: #cc4a5b; font-weight: 900; }}
    .goal-progress, .progress-track {{
        overflow: hidden;
        height: 0.55rem;
        margin-top: 0.75rem;
        border-radius: 999px;
        background: rgba(17,19,24,0.08);
    }}
    .goal-progress span, .progress-track span {{
        display: block;
        height: 100%;
        border-radius: inherit;
        background: linear-gradient(90deg, var(--lime), var(--blue));
    }}
    .chart-intro, .history-summary {{
        margin: 0.9rem 0;
        padding: 1rem 1.1rem;
        border-radius: 22px;
        background: rgba(255,255,255,0.88);
        border: 1px solid rgba(255,255,255,0.92);
        box-shadow: var(--soft);
        color: var(--ink);
    }}
    .answer-good {{
        border-color: rgba(13,144,111,0.25);
        background: rgba(226,248,242,0.90);
    }}
    .answer-care {{
        border-color: rgba(204,138,47,0.28);
        background: rgba(255,244,222,0.92);
    }}
    .answer-risk {{
        border-color: rgba(204,74,91,0.28);
        background: rgba(255,232,235,0.92);
    }}
    div[data-testid="stTabs"] button {{
        border-radius: 999px;
        color: var(--ink);
        font-weight: 800;
        background: rgba(255,255,255,0.78);
        border: 1px solid rgba(255,255,255,0.90);
        box-shadow: var(--soft);
    }}
    div[data-testid="stTabs"] button[aria-selected="true"] {{
        color: white;
        background: #111318;
        border-color: #111318;
    }}
    div[data-testid="stPlotlyChart"],
    div[data-testid="stDataFrame"],
    div[data-testid="stExpander"],
    div[data-testid="stForm"],
    [data-testid="stAlert"] {{
        border-radius: 22px;
        background: rgba(255,255,255,0.90);
        border: 1px solid rgba(255,255,255,0.92);
        box-shadow: var(--soft);
    }}
    .stButton > button,
    [data-testid="stFormSubmitButton"] button {{
        min-height: 2.7rem;
        border: 0;
        border-radius: 999px;
        color: white;
        background: #111318;
        font-weight: 800;
        box-shadow: var(--soft);
    }}
    .stButton > button:hover,
    [data-testid="stFormSubmitButton"] button:hover {{
        background: #2a2f3a;
        color: white;
    }}
    div[data-testid="stDownloadButton"] button {{
        min-height: 2.7rem;
        border-radius: 999px;
        color: var(--ink);
        background: rgba(255,255,255,0.94);
        border: 1px solid rgba(17,19,24,0.10);
        font-weight: 800;
    }}
    @media (max-width: 980px) {{
        .metric-grid, .indicator-grid, .answer-grid, .goal-grid, .debt-grid, .investment-grid {{
            grid-template-columns: 1fr;
        }}
    }}
</style>
""",
    unsafe_allow_html=True,
)

# ====================== BANCO DE DADOS ======================
def init_db():
    conn = sqlite3.connect(DB_FILE)
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
    conn = sqlite3.connect(DB_FILE)
    try:
        df = pd.read_sql_query("SELECT * FROM transacoes ORDER BY data DESC, id DESC", conn)
    except Exception:
        df = pd.DataFrame(columns=["id", "data", "descricao", "categoria", "valor", "tipo", "cartao"])
    conn.close()
    return df


def carregar_investimentos() -> pd.DataFrame:
    conn = sqlite3.connect(DB_FILE)
    try:
        df = pd.read_sql_query("SELECT * FROM investimentos ORDER BY data DESC, id DESC", conn)
    except Exception:
        df = pd.DataFrame(columns=["id", "data", "tipo", "valor", "rentabilidade", "descricao", "status"])
    conn.close()
    return df


def carregar_dividas() -> pd.DataFrame:
    conn = sqlite3.connect(DB_FILE)
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
    conn = sqlite3.connect(DB_FILE)
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
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        "INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao) VALUES (?, ?, ?, ?, ?, ?)",
        (str(data), descricao, categoria, float(valor), tipo, cartao),
    )
    conn.commit()
    conn.close()


def excluir_transacao(id_):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes WHERE id = ?", (int(id_),))
    conn.commit()
    conn.close()


def limpar_historico():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes")
    conn.commit()
    conn.close()


def salvar_investimento(data, tipo, valor, rentabilidade, descricao, status):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        "INSERT INTO investimentos (data, tipo, valor, rentabilidade, descricao, status) VALUES (?, ?, ?, ?, ?, ?)",
        (str(data), tipo, float(valor), rentabilidade, descricao, status),
    )
    conn.commit()
    conn.close()


def excluir_investimento(id_):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM investimentos WHERE id = ?", (int(id_),))
    conn.commit()
    conn.close()


def salvar_divida(
    data, credor, tipo, saldo_original, desconto, saldo_negociado,
    parcela_possivel, vencimento, prioridade, consequencia, status,
    proxima_acao, anotacoes,
):
    conn = sqlite3.connect(DB_FILE)
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
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM dividas WHERE id = ?", (int(id_),))
    conn.commit()
    conn.close()


def salvar_meta(data, nome, valor_meta, valor_atual, aporte_mensal, prazo, status, anotacoes):
    conn = sqlite3.connect(DB_FILE)
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
    conn = sqlite3.connect(DB_FILE)
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


def style_plot(fig):
    fig.update_layout(
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(color="#111318", size=12),
        title=dict(font=dict(size=18, color="#111318"), x=0.04),
        margin=dict(l=34, r=26, t=70, b=48),
        height=400,
        legend=dict(orientation="h", y=1.02, x=1, xanchor="right"),
        hoverlabel=dict(bgcolor="#111318", font_color="#ffffff"),
        separators=",.",
    )
    fig.update_xaxes(gridcolor="rgba(17,19,24,0.08)", tickfont=dict(color="#747985"))
    fig.update_yaxes(gridcolor="rgba(17,19,24,0.08)", tickfont=dict(color="#747985"))
    return fig


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

    conn = sqlite3.connect(DB_FILE)
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
<div style="display:flex; justify-content:space-between; align-items:flex-start; gap:1rem; margin-bottom:1.2rem;">
    <div>
        <div style="color:#747985; font-weight:850; font-size:0.85rem; text-transform:uppercase;">Painel inteligente</div>
        <h1 style="margin:0.3rem 0 0.4rem; font-size:clamp(2.4rem,5vw,4.2rem); line-height:0.98; font-weight:900;">Dashboard<br>Financeiro</h1>
        <p style="color:#747985; max-width:40rem; line-height:1.6;">Controle entradas, gastos, dívidas, metas e investimentos em uma visão clara e simples.</p>
    </div>
    <div style="display:flex; gap:0.6rem; align-items:center;">
        <div style="padding:0.45rem 1rem; border-radius:999px; background:#111318; color:white; font-weight:800; font-size:0.86rem;">Saúde: {score}/100</div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# Cards principais
st.markdown(
    f"""
<div class="metric-grid">
    <div class="metric-card accent">
        <div class="metric-label">Saldo atual</div>
        <div class="metric-value">{brl(saldo)}</div>
        <div class="metric-foot">Resultado de tudo que foi registrado</div>
    </div>
    <div class="metric-card">
        <div class="metric-label">Entradas</div>
        <div class="metric-value">{brl(entradas)}</div>
        <div class="metric-foot">Receitas manuais e importadas</div>
    </div>
    <div class="metric-card blue">
        <div class="metric-label">Saídas</div>
        <div class="metric-value">{brl(saidas)}</div>
        <div class="metric-foot">Despesas acumuladas</div>
    </div>
    <div class="metric-card dark">
        <div class="metric-label">Patrimônio</div>
        <div class="metric-value">{brl(total_investido)}</div>
        <div class="metric-foot">Investimentos registrados</div>
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
            color_discrete_map={"Bom": "#0d906f", "Normal": "#28c7b7", "Apertado": "#cc4a5b"},
        )
        fig.update_yaxes(tickprefix="R$ ")
        st.plotly_chart(style_plot(fig), use_container_width=True)

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
            )
            fig_saldo.update_traces(line=dict(color="#111318", width=3), fillcolor="rgba(217,255,0,0.34)")
            fig_saldo.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig_saldo), use_container_width=True)

        cat_total = df_chart.groupby("categoria", as_index=False)["valor_abs"].sum().sort_values("valor_abs", ascending=False)
        fluxo_tipo = df_chart.groupby(["categoria", "tipo"], as_index=False)["valor_abs"].sum()

        c1, c2 = st.columns(2)
        with c1:
            fig = px.pie(
                cat_total, names="categoria", values="valor_abs", hole=0.58,
                title="Distribuição por categoria",
                color_discrete_sequence=["#d9ff00", "#111318", "#8fb1ff", "#d7dbe2", "#efffb4", "#5c6573"],
            )
            fig.update_traces(textposition="outside", textinfo="percent+label")
            st.plotly_chart(style_plot(fig), use_container_width=True)
        with c2:
            fig2 = px.bar(
                fluxo_tipo, x="categoria", y="valor_abs", color="tipo",
                title="Entradas x Saídas",
                color_discrete_map={"Entrada": "#d9ff00", "Saída": "#111318"},
            )
            fig2.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig2), use_container_width=True)

        c3, c4 = st.columns(2)
        with c3:
            top = cat_total.head(8).sort_values("valor_abs", ascending=True)
            fig3 = px.bar(
                top, x="valor_abs", y="categoria", orientation="h",
                title="Categorias que mais movimentam dinheiro",
                color="valor_abs", color_continuous_scale=["#e9edf2", "#8fb1ff", "#111318"],
            )
            fig3.update_layout(coloraxis_showscale=False)
            fig3.update_xaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig3), use_container_width=True)
        with c4:
            pag = df_chart.groupby("cartao", as_index=False)["valor_abs"].sum().sort_values("valor_abs", ascending=False).head(7)
            fig4 = px.pie(
                pag, names="cartao", values="valor_abs", hole=0.62,
                title="Formas de pagamento",
                color_discrete_sequence=["#111318", "#d9ff00", "#8fb1ff", "#d7dbe2", "#efffb4"],
            )
            st.plotly_chart(style_plot(fig4), use_container_width=True)

        if len(fluxo):
            fig_m = px.bar(
                fluxo, x="mes", y=["entradas", "saidas"], barmode="group",
                title="Evolução mensal consolidada",
                color_discrete_map={"entradas": "#d9ff00", "saidas": "#111318"},
            )
            fig_m.update_xaxes(tickformat="%m/%Y")
            fig_m.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig_m), use_container_width=True)
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
            fig = px.pie(tipo_inv, names="tipo", values="valor", hole=0.58, title="Distribuição dos investimentos")
            st.plotly_chart(style_plot(fig), use_container_width=True)
        with c2:
            fig2 = px.bar(
                investimentos.groupby("status", as_index=False)["valor"].sum(),
                x="status", y="valor", color="status", title="Valores por status",
            )
            st.plotly_chart(style_plot(fig2), use_container_width=True)

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
        cats = sorted(df["categoria"].replace("", "Sem categoria").unique())
        cat_f = c3.selectbox("Categoria", ["Todas", *cats])

        df_h = df.copy()
        if busca.strip():
            termo = normalizar(busca)
            df_h = df_h[df_h.apply(
                lambda r: termo in normalizar(" ".join([str(r["descricao"]), str(r["categoria"]), str(r["cartao"])])),
                axis=1,
            )]
        if tipo_f != "Todos":
            df_h = df_h[df_h["tipo"] == tipo_f]
        if cat_f != "Todas":
            df_h = df_h[df_h["categoria"] == cat_f]

        tabela = df_h[["data", "descricao", "categoria", "tipo", "cartao", "valor"]].rename(
            columns={"data": "Data", "descricao": "Descrição", "categoria": "Categoria",
                     "tipo": "Tipo", "cartao": "Pagamento", "valor": "Valor"}
        )
        tabela["Data"] = tabela["Data"].map(data_br)
        tabela["Valor"] = tabela["Valor"].map(brl)
        st.dataframe(tabela, use_container_width=True, hide_index=True)
        st.caption(f"{len(df_h)} de {len(df)} movimentações exibidas.")

        for _, row in df_h.iterrows():
            classe = "positive" if row["valor"] >= 0 else "negative"
            st.markdown(
                f"""
<div class="history-item">
    <div>
        <div style="font-weight:850;">{escape(str(row['descricao'] or 'Sem descrição'))}</div>
        <div style="color:#747985; font-size:0.88rem;">{escape(str(row['categoria']))} • {data_br(row['data'])} • {escape(str(row['cartao'] or '—'))}</div>
    </div>
    <div class="{classe}">{brl(row['valor'])}</div>
</div>
""",
                unsafe_allow_html=True,
            )
            if st.button("🗑️ Apagar", key=f"del_t_{row['id']}"):
                excluir_transacao(row["id"])
                st.rerun()
    else:
        st.info("Nenhum registro ainda.")

st.caption("Dashboard Financeiro • Visão clara e simples • Uso pessoal")
