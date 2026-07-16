import re
import sqlite3
import textwrap
import zipfile
from datetime import date
from html import escape
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


st.set_page_config(
    page_title="Dashboard Financeiro",
    layout="wide",
    page_icon="DF",
)

DB_FILE = "financeiro.db"


def brl(valor):
    texto = f"R$ {float(valor or 0):,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


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


def brl_compacto(valor):
    return brl_curto(valor)


def pct(valor):
    return f"{float(valor or 0):.0f}%".replace(".", ",")


def data_br(valor):
    data_convertida = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.isna(data_convertida):
        return "Sem data"
    return data_convertida.strftime("%d/%m/%Y")


def limpar_texto(valor, padrao=""):
    if pd.isna(valor):
        return padrao
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return padrao
    return texto


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


def mensagem_erro_usuario(_erro):
    return "Nao conseguimos concluir agora. Tente novamente."


def converter_valor(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    if not texto:
        return None
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

PALAVRAS_ENTRADA_PLANILHA = ["entrada", "receita", "salario", "renda", "freelance", "reembolso"]
PALAVRAS_SAIDA_PLANILHA = ["saida", "despesa", "debito", "gasto", "conta", "cartao"]


def data_planilha_mensal(nome_aba, valor_data=None):
    mes = MESES_PLANILHA.get(normalizar_coluna(nome_aba), date.today().month)
    data_convertida = pd.to_datetime(valor_data, errors="coerce", dayfirst=True)
    if pd.isna(data_convertida):
        return date(date.today().year, mes, 1).isoformat()
    return data_convertida.date().isoformat()


def texto_contexto_planilha(df_aba, linha, col_inicio, col_fim):
    textos = []
    for idx_linha in range(max(0, linha - 4), linha + 1):
        for idx_coluna in range(max(0, col_inicio), min(len(df_aba.columns), col_fim + 1)):
            textos.append(normalizar_coluna(celula_planilha(df_aba, idx_linha, idx_coluna)))
    return " ".join(textos)


def inferir_tipo_mensal(descricao="", categoria="", tipo="", contexto=""):
    texto = normalizar_coluna(" ".join([str(descricao), str(categoria), str(tipo), str(contexto)]))
    texto_direto = normalizar_coluna(" ".join([str(descricao), str(categoria), str(tipo)]))
    if "saida" in texto or any(palavra in texto_direto for palavra in ["debito", "credito", "despesa", "gasto"]):
        return "Saída"
    if any(palavra in texto for palavra in PALAVRAS_ENTRADA_PLANILHA):
        return "Entrada"
    if any(palavra in texto for palavra in PALAVRAS_SAIDA_PLANILHA):
        return "Saída"
    return "Saída"


def adicionar_movimentacao_mensal(linhas, data_mov, descricao, categoria, valor, tipo, cartao):
    if valor is None or valor == 0:
        return
    tipo_final = "Entrada" if tipo == "Entrada" else "Saída"
    valor_final = abs(valor) if tipo_final == "Entrada" else -abs(valor)
    linhas.append(
        {
            "data": data_mov,
            "descricao": texto_planilha(descricao, "Importado da planilha"),
            "categoria": texto_planilha(categoria, "Importado"),
            "valor": valor_final,
            "tipo": tipo_final,
            "cartao": texto_planilha(cartao, "Planilha"),
        }
    )


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


def preparar_modelo_organizacao_financeira(planilhas):
    linhas = []
    for nome_aba, df_aba in planilhas.items():
        if MESES_PLANILHA.get(normalizar_coluna(nome_aba)) is None:
            continue
        for linha_idx in range(len(df_aba.index)):
            for col_idx in range(len(df_aba.columns) - 1):
                if normalizar_coluna(celula_planilha(df_aba, linha_idx, col_idx)) != "entradas":
                    continue
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

            for idx in range(linha_idx + 1, len(df_aba.index)):
                descricao = texto_planilha(celula_planilha(df_aba, idx, col_descricao))
                valor_original = converter_valor(celula_planilha(df_aba, idx, col_valor))
                desc_norm = normalizar_coluna(descricao)
                if not descricao and valor_original is None:
                    break
                if not descricao:
                    continue
                if desc_norm.startswith("total") or desc_norm in ["saidas", "entradas", "investimentos", "reserva"]:
                    break
                if valor_original is None or valor_original == 0:
                    continue
                tipo_original = texto_planilha(celula_planilha(df_aba, idx, col_tipo), "Planilha") if col_tipo is not None else "Planilha"
                categoria = texto_planilha(celula_planilha(df_aba, idx, col_categoria), "Receita" if "entrada" in contexto else "Outros") if col_categoria is not None else ("Receita" if "entrada" in contexto else "Outros")
                tipo = inferir_tipo_mensal(descricao, categoria, tipo_original, contexto)
                data_mov = data_planilha_mensal(nome_aba, celula_planilha(df_aba, idx, col_data) if col_data is not None else None)
                adicionar_movimentacao_mensal(linhas, data_mov, descricao, categoria, valor_original, tipo, tipo_original)
    return pd.DataFrame(linhas).drop_duplicates() if linhas else pd.DataFrame()


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


def preparar_movimentacoes_importadas(dados):
    if isinstance(dados, dict):
        mensal = preparar_modelo_organizacao_financeira(dados)
        if not mensal.empty:
            return mensal, 0
        primeira = next(iter(dados.values()), pd.DataFrame())
        if primeira.empty:
            return pd.DataFrame(), 0
        primeira = primeira.copy()
        primeira.columns = primeira.iloc[0]
        dados = primeira.iloc[1:].reset_index(drop=True)

    aliases = {
        "data": ["data", "dt", "dia"],
        "descricao": ["descricao", "descrição", "historico", "histórico", "nome"],
        "categoria": ["categoria", "grupo"],
        "valor": ["valor", "valor r$", "valor rs"],
        "tipo": ["tipo", "natureza"],
        "cartao": ["forma de pagamento", "pagamento", "cartao", "cartão", "conta"],
    }
    colunas = {normalizar_coluna(coluna): coluna for coluna in dados.columns}
    mapa = {}
    for destino, opcoes in aliases.items():
        for opcao in opcoes:
            if normalizar_coluna(opcao) in colunas:
                mapa[destino] = colunas[normalizar_coluna(opcao)]
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
        data_convertida = pd.to_datetime(row.get(mapa.get("data"), date.today()), errors="coerce", dayfirst=True)
        if pd.isna(data_convertida):
            data_convertida = pd.Timestamp(date.today())
        tipo_texto = normalizar_coluna(row.get(mapa.get("tipo"), ""))
        if any(palavra in tipo_texto for palavra in PALAVRAS_ENTRADA_PLANILHA):
            tipo = "Entrada"
        elif any(palavra in tipo_texto for palavra in PALAVRAS_SAIDA_PLANILHA):
            tipo = "Saída"
        else:
            tipo = "Entrada" if valor_original > 0 else "Saída"
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


def normalizar_dataframe_financeiro(df):
    return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()


def normalizar_dataframe_dividas(df):
    return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()


def normalizar_dataframe_metas(df):
    return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()


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


def carregar_dados():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM transacoes ORDER BY data DESC, id DESC", conn)
    conn.close()
    if df.empty:
        return pd.DataFrame(columns=["id", "data", "descricao", "categoria", "valor", "tipo", "cartao"])
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)
    for coluna in ["data", "descricao", "categoria", "tipo", "cartao"]:
        df[coluna] = df[coluna].fillna("").astype(str)
    return df


def carregar_investimentos():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM investimentos ORDER BY data DESC, id DESC", conn)
    conn.close()
    if df.empty:
        return pd.DataFrame(columns=["id", "data", "tipo", "valor", "rentabilidade", "descricao", "status"])
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)
    for coluna in ["data", "tipo", "rentabilidade", "descricao", "status"]:
        df[coluna] = df[coluna].fillna("").astype(str)
    return df


def carregar_dividas():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM dividas ORDER BY data DESC, id DESC", conn)
    conn.close()
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


def excluir_transacao(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def excluir_investimento(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM investimentos WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def excluir_divida(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM dividas WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def excluir_meta(tid):
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM metas WHERE id = ?", (tid,))
    conn.commit()
    conn.close()


def limpar_historico():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes")
    conn.commit()
    conn.close()


def limpar_historico_movimentacoes():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("DELETE FROM transacoes")
    conn.commit()
    conn.close()


def salvar_investimento(data_inv, tipo, valor, rentabilidade, descricao, status):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        "INSERT INTO investimentos (data, tipo, valor, rentabilidade, descricao, status) VALUES (?, ?, ?, ?, ?, ?)",
        (str(data_inv), tipo, valor, rentabilidade, descricao, status),
    )
    conn.commit()
    conn.close()


def salvar_divida(data_divida, credor, tipo, saldo_original, desconto, saldo_negociado, parcela_possivel, vencimento, prioridade, consequencia, status, proxima_acao, anotacoes):
    conn = sqlite3.connect(DB_FILE)
    conn.execute(
        """
        INSERT INTO dividas
            (data, credor, tipo, saldo_original, desconto, saldo_negociado, parcela_possivel, vencimento, prioridade, consequencia, status, proxima_acao, anotacoes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            str(data_divida),
            credor,
            tipo,
            saldo_original,
            desconto,
            saldo_negociado,
            parcela_possivel,
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
        (str(data_meta), nome, valor_meta, valor_atual, aporte_mensal, prazo, status, anotacoes),
    )
    conn.commit()
    conn.close()


def chave_movimentacao(registro):
    data_mov = pd.to_datetime(registro.get("data"), errors="coerce", dayfirst=True)
    data_norm = data_mov.date().isoformat() if not pd.isna(data_mov) else limpar_texto(registro.get("data"))
    return (
        data_norm,
        normalizar_coluna(registro.get("descricao")),
        round(float(registro.get("valor") or 0), 2),
        normalizar_coluna(registro.get("tipo")),
    )


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


def importar_movimentacoes(df_importado):
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

            col_descricao = next((i for i, item in enumerate(valores_linha) if item in ["descricao", "nome"]), None)
            col_valor = next((i for i, item in enumerate(valores_linha) if item == "valor" and i > col_descricao), None)
            col_tipo = next((i for i, item in enumerate(valores_linha) if item in ["tipo", "pagamento"] and i > col_descricao), None)
            col_categoria = next((i for i, item in enumerate(valores_linha) if item == "categoria" and i > col_descricao), None)
            if col_descricao is None or col_valor is None:
                continue

            contexto = " ".join(
                normalizar(df_aba.iat[i, j])
                for i in range(max(0, linha_idx - 4), linha_idx + 1)
                for j in range(max(0, col_descricao - 2), min(len(df_aba.columns), col_valor + 3))
            )

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
    return pd.DataFrame(linhas).drop_duplicates() if linhas else pd.DataFrame()


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


def gerar_pdf(df, investimentos, dividas):
    linhas = []

    def add(texto="", largura=92):
        for parte in textwrap.wrap(str(texto), width=largura) or [""]:
            linhas.append(parte)

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


PLOTLY_CONFIG = {"displayModeBar": False, "responsive": True}


def style_plot(fig, height=360):
    fig.update_layout(
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(color="#111318", size=11),
        title=dict(font=dict(size=17, color="#111318"), x=0.02, xanchor="left"),
        margin=dict(l=24, r=18, t=62, b=78),
        height=height,
        autosize=True,
        legend=dict(
            orientation="h",
            y=-0.18,
            x=0,
            xanchor="left",
            yanchor="top",
            font=dict(size=10, color="#111318"),
            title=None,
        ),
        hoverlabel=dict(bgcolor="#111318", font_color="#ffffff"),
        separators=",.",
        uniformtext_minsize=9,
        uniformtext_mode="hide",
    )
    fig.update_xaxes(gridcolor="rgba(17,19,24,0.08)", tickfont=dict(color="#747985", size=10), automargin=True)
    fig.update_yaxes(gridcolor="rgba(17,19,24,0.08)", tickfont=dict(color="#747985", size=10), automargin=True)
    return fig


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
    *, *::before, *::after {{
        box-sizing: border-box;
    }}
    body, p, label, span, div, button, input, textarea {{
        letter-spacing: 0 !important;
    }}
    [data-testid="stAppViewContainer"] {{
        background:
            radial-gradient(circle at 12% 8%, rgba(255,255,255,.95), transparent 22rem),
            radial-gradient(circle at 88% 12%, rgba(217,255,0,.18), transparent 24rem),
            radial-gradient(circle at 72% 86%, rgba(143,177,255,.22), transparent 28rem),
            linear-gradient(145deg, #f6f8fb 0%, #eef3f8 46%, #ffffff 100%);
        color: var(--ink);
    }}
    [data-testid="stHeader"] {{
        background: transparent;
    }}
    #MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {{
        display: none !important;
        visibility: hidden !important;
    }}
    .main .block-container {{
        max-width: 1220px;
        padding-top: 1.6rem;
        padding-bottom: 3rem;
    }}
    .main .block-container, .classic-header, .hero-card, .overview-card,
    .metric-card, .indicator-card, .history-card,
    div[data-testid="column"], div[data-testid="stVerticalBlock"] {{
        min-width: 0;
    }}
    h1, h2, h3 {{
        color: var(--ink) !important;
        letter-spacing: 0;
    }}
    .main p, .main label, [data-testid="stWidgetLabel"] p,
    [data-testid="stRadio"] label, [data-testid="stRadio"] p {{
        color: var(--ink) !important;
        overflow-wrap: anywhere;
        word-break: normal;
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
        overflow: visible;
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
        overflow-wrap: anywhere;
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
        min-height: 2.65rem;
        white-space: normal !important;
        overflow-wrap: anywhere;
    }}
    div[data-testid="stTabs"] [data-baseweb="tab-list"] {{
        gap: .55rem;
        flex-wrap: wrap;
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
    div[data-testid="stForm"] {{
        padding: 1rem;
    }}
    div[data-testid="stDataFrame"], div[data-testid="stPlotlyChart"] {{
        overflow-x: auto;
    }}
    div[data-testid="stPlotlyChart"] {{
        padding: .55rem .35rem .15rem;
    }}
    .js-plotly-plot .plotly .main-svg {{
        overflow: visible !important;
    }}
    .stButton > button, [data-testid="stDownloadButton"] button, [data-testid="stFormSubmitButton"] button {{
        min-height: 2.85rem;
        white-space: normal !important;
        overflow-wrap: anywhere;
        line-height: 1.2;
        border-radius: 999px !important;
    }}
    [data-baseweb="input"] > div,
    [data-baseweb="textarea"] > div,
    [data-baseweb="select"] > div {{
        background: #ffffff !important;
        color: var(--ink) !important;
        border: 1px solid rgba(17,19,24,.12) !important;
        border-radius: 16px !important;
        min-height: 3rem;
    }}
    [data-baseweb="input"] input,
    [data-baseweb="textarea"] textarea,
    [data-baseweb="select"] span,
    [data-baseweb="select"] div {{
        color: var(--ink) !important;
        -webkit-text-fill-color: var(--ink) !important;
    }}
    [data-baseweb="input"] input::placeholder,
    [data-baseweb="textarea"] textarea::placeholder {{
        color: rgba(17,19,24,.44) !important;
        -webkit-text-fill-color: rgba(17,19,24,.44) !important;
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
    @media (max-width: 760px) {{
        [data-testid="stAppViewContainer"] {{
            background:
                radial-gradient(circle at 18% 4%, rgba(217,255,0,.16), transparent 16rem),
                linear-gradient(145deg, #f8fbfd 0%, #eef5f7 58%, #ffffff 100%);
        }}
        .main .block-container {{
            padding: .85rem .75rem 2rem;
        }}
        .classic-header {{
            gap: .75rem;
            margin-bottom: .85rem;
        }}
        .hero-card, .overview-card {{
            border-radius: 22px;
            padding: 1rem;
            box-shadow: 0 14px 42px rgba(17,19,24,.09);
        }}
        .hero-card {{
            min-height: auto;
        }}
        .hero-top {{
            align-items: flex-start;
            flex-direction: column;
        }}
        .brand-mark {{
            width: 100%;
            font-size: .95rem;
        }}
        .header-pill {{
            width: 100%;
            justify-content: center;
        }}
        .hero-card h1 {{
            margin: 1.35rem 0 .7rem;
            font-size: clamp(2rem, 13vw, 3rem);
            line-height: 1.03;
        }}
        .hero-card p, .header-quote {{
            font-size: .92rem;
            line-height: 1.55;
        }}
        .header-quote {{
            margin-top: 1rem;
            padding-left: .8rem;
        }}
        .overview-head {{
            align-items: flex-start;
            flex-direction: column;
            gap: .35rem;
        }}
        .overview-item, .metric-card, .indicator-card, .history-card {{
            min-height: auto;
            padding: .95rem;
            border-radius: 18px;
        }}
        .overview-value, .metric-value, .indicator-value {{
            font-size: clamp(1.25rem, 8vw, 1.85rem);
            line-height: 1.14;
        }}
        .metric-foot, .indicator-note {{
            font-size: .82rem;
        }}
        h2 {{
            font-size: 1.35rem !important;
            line-height: 1.2 !important;
        }}
        h3 {{
            font-size: 1.15rem !important;
            line-height: 1.25 !important;
        }}
        div[data-testid="stForm"] {{
            padding: .95rem .8rem;
            border-radius: 22px;
        }}
        [data-baseweb="input"] > div,
        [data-baseweb="textarea"] > div,
        [data-baseweb="select"] > div {{
            min-height: 3.25rem;
        }}
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {{
            display: flex;
            gap: .45rem;
        }}
        div[data-testid="stTabs"] button {{
            flex: 1 1 calc(50% - .45rem);
            justify-content: center;
            min-width: 8.8rem;
            padding: .6rem .75rem;
        }}
        div[data-testid="stPlotlyChart"] {{
            border-radius: 18px;
            padding: .35rem .15rem 0;
        }}
    }}
    @media (max-width: 430px) {{
        .main .block-container {{
            padding-left: .55rem;
            padding-right: .55rem;
        }}
        .hero-card h1 {{
            font-size: clamp(1.8rem, 12vw, 2.55rem);
        }}
        .overview-value, .metric-value, .indicator-value {{
            font-size: clamp(1.18rem, 7.5vw, 1.65rem);
        }}
        div[data-testid="stTabs"] button {{
            flex-basis: 100%;
            min-width: 100%;
        }}
    }}
</style>
""",
    unsafe_allow_html=True,
)


init_db()
df = carregar_dados()
investimentos = carregar_investimentos()
dividas = carregar_dividas()

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

aba = st.tabs(["Nova Movimentacao", "Dashboard", "Investimentos", "Dividas", "Historico"])

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
                st.error(f"Nao consegui importar essa planilha: {erro}")

    if len(df):
        df_chart = df.copy()
        df_chart["data_convertida"] = pd.to_datetime(df_chart["data"], errors="coerce")
        df_chart["valor_abs"] = df_chart["valor"].abs()
        df_chart["categoria"] = df_chart["categoria"].replace("", "Sem categoria")
        df_chart["cartao"] = df_chart["cartao"].replace("", "Nao informado")
        df_timeline = df_chart.dropna(subset=["data_convertida"]).sort_values("data_convertida")
        if len(df_timeline):
            df_timeline["saldo_acumulado"] = df_timeline["valor"].cumsum()
            fig_saldo = px.area(df_timeline, x="data_convertida", y="saldo_acumulado", title="Saldo acumulado ao longo do tempo")
            fig_saldo.update_traces(line=dict(color="#111318", width=3), fillcolor="rgba(217,255,0,.34)")
            fig_saldo.update_yaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig_saldo, height=340), use_container_width=True, config=PLOTLY_CONFIG)

        col_g1, col_g2 = st.columns(2)
        categoria_total = df_chart.groupby("categoria", as_index=False)["valor_abs"].sum().sort_values("valor_abs", ascending=False)
        categoria_grafico = categoria_total.copy()
        if len(categoria_grafico) > 7:
            principais = categoria_grafico.head(6)
            outras = pd.DataFrame([{"categoria": "Outras", "valor_abs": categoria_grafico.iloc[6:]["valor_abs"].sum()}])
            categoria_grafico = pd.concat([principais, outras], ignore_index=True)
        fluxo = df_chart.groupby(["categoria", "tipo"], as_index=False)["valor_abs"].sum()
        with col_g1:
            fig = px.pie(categoria_grafico, names="categoria", values="valor_abs", hole=0.58, title="Distribuicao por categoria", color_discrete_sequence=["#d9ff00", "#111318", "#8fb1ff", "#d7dbe2", "#efffb4", "#5c6573"])
            fig.update_traces(
                textposition="inside",
                textinfo="percent",
                insidetextorientation="radial",
                marker=dict(line=dict(color="#ffffff", width=2)),
                hovertemplate="<b>%{label}</b><br>Valor: R$ %{value:,.2f}<br>%{percent}<extra></extra>",
            )
            st.plotly_chart(style_plot(fig, height=350), use_container_width=True, config=PLOTLY_CONFIG)
        with col_g2:
            fig2 = px.bar(fluxo, x="categoria", y="valor_abs", color="tipo", title="Entradas x Saidas", color_discrete_map={"Entrada": "#d9ff00", "Saida": "#111318"})
            fig2.update_yaxes(tickprefix="R$ ")
            fig2.update_xaxes(tickangle=-25)
            st.plotly_chart(style_plot(fig2, height=350), use_container_width=True, config=PLOTLY_CONFIG)

        col_g3, col_g4 = st.columns(2)
        with col_g3:
            top_categorias = categoria_total.head(8).sort_values("valor_abs", ascending=True)
            fig3 = px.bar(top_categorias, x="valor_abs", y="categoria", orientation="h", title="Categorias que mais movimentam dinheiro", color="valor_abs", color_continuous_scale=["#e9edf2", "#8fb1ff", "#111318"])
            fig3.update_layout(coloraxis_showscale=False)
            fig3.update_xaxes(tickprefix="R$ ")
            st.plotly_chart(style_plot(fig3, height=360), use_container_width=True, config=PLOTLY_CONFIG)
        with col_g4:
            pagamentos_total = df_chart.groupby("cartao", as_index=False)["valor_abs"].sum().sort_values("valor_abs", ascending=False)
            pagamentos = pagamentos_total.copy()
            if len(pagamentos) > 7:
                principais = pagamentos.head(6)
                outras = pd.DataFrame([{"cartao": "Outras", "valor_abs": pagamentos.iloc[6:]["valor_abs"].sum()}])
                pagamentos = pd.concat([principais, outras], ignore_index=True)
            fig4 = px.pie(pagamentos, names="cartao", values="valor_abs", hole=0.62, title="Formas de pagamento", color_discrete_sequence=["#111318", "#d9ff00", "#8fb1ff", "#d7dbe2", "#efffb4"])
            fig4.update_traces(
                textposition="inside",
                textinfo="percent",
                insidetextorientation="radial",
                marker=dict(line=dict(color="#ffffff", width=2)),
                hovertemplate="<b>%{label}</b><br>Valor: R$ %{value:,.2f}<br>%{percent}<extra></extra>",
            )
            st.plotly_chart(style_plot(fig4, height=350), use_container_width=True, config=PLOTLY_CONFIG)
    else:
        st.info("Cadastre ou importe movimentacoes para ver os graficos.")

with aba[2]:
    st.subheader("Investimentos")
    with st.form("form_investimento"):
        col1, col2, col3 = st.columns(3)
        with col1:
            data_inv = st.date_input("Data", value=date.today(), format="DD/MM/YYYY", key="data_inv")
            tipo_inv = st.selectbox("Tipo", ["Reserva de emergencia", "Tesouro Direto", "CDB", "LCI / LCA", "Acoes", "FII", "Outro"])
        with col2:
            valor_inv = st.number_input("Valor investido", value=0.0, min_value=0.0, step=0.01)
            rentabilidade = st.text_input("Rentabilidade")
        with col3:
            descricao_inv = st.text_input("Descricao do investimento")
            status_inv = st.selectbox("Status", ["Ativo", "Planejado", "Resgatado"])
        if st.form_submit_button("Salvar investimento"):
            if valor_inv <= 0:
                st.error("Informe um valor maior que zero.")
            else:
                salvar_investimento(data_inv, tipo_inv, valor_inv, rentabilidade, descricao_inv, status_inv)
                st.success("Investimento salvo.")
                st.rerun()
    if len(investimentos):
        st.dataframe(investimentos.assign(data=investimentos["data"].map(data_br), valor=investimentos["valor"].map(brl)), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum investimento cadastrado.")

with aba[3]:
    st.subheader("Dividas e negociacao")
    with st.form("form_divida"):
        col1, col2, col3 = st.columns(3)
        with col1:
            data_div = st.date_input("Data do registro", value=date.today(), format="DD/MM/YYYY")
            credor = st.text_input("Credor")
            tipo_divida = st.selectbox("Tipo", ["Cartao de credito", "Emprestimo", "Conta atrasada", "Financiamento", "Acordo", "Outro"])
        with col2:
            saldo_original = st.number_input("Saldo original", value=0.0, min_value=0.0, step=0.01)
            desconto = st.number_input("Desconto", value=0.0, min_value=0.0, step=0.01)
            saldo_negociado = st.number_input("Saldo negociado", value=0.0, min_value=0.0, step=0.01)
        with col3:
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
        st.info("Nenhuma divida cadastrada.")

with aba[4]:
    st.subheader("Historico")
    st.markdown(
        f"""<div class="history-summary"><strong>Relatorio detalhado</strong><br>Registros incluidos: {len(df)} movimentacoes, {len(investimentos)} investimentos e {len(dividas)} dividas.</div>""",
        unsafe_allow_html=True,
    )
    st.download_button(
        "Baixar relatorio em PDF",
        data=gerar_pdf(df, investimentos, dividas),
        file_name=f"relatorio-financeiro-{date.today().strftime('%d-%m-%Y')}.pdf",
        mime="application/pdf",
    )
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
        tabela["Data"] = tabela["Data"].map(data_br)
        tabela["Valor"] = tabela["Valor"].map(brl)
        st.dataframe(tabela, use_container_width=True, hide_index=True)
        st.caption(f"{len(df_hist)} de {len(df)} movimentacoes exibidas.")
        for _, row in df_hist.iterrows():
            classe = "positive" if row["valor"] >= 0 else "negative"
            st.markdown(
                f"""<div class="history-card"><strong>{escape(row["descricao"])}</strong><br><span>{escape(row["categoria"])} - {data_br(row["data"])} - {escape(row["cartao"])}</span><div class="{classe}">{brl(row["valor"])}</div></div>""",
                unsafe_allow_html=True,
            )
            if st.button("Apagar", key=f"del{row['id']}"):
                excluir_transacao(row["id"])
                st.rerun()
    else:
        st.info("Nenhum registro ainda.")

st.caption("Dashboard Financeiro - Visao financeira clara - Experiencia premium")
