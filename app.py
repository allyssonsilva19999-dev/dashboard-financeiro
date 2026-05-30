import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
from datetime import date

# ====================== CONFIGURAÇÃO ======================
st.set_page_config(
    page_title="Dashboard Financeiro Premium",
    layout="wide",
    page_icon="💰",
    initial_sidebar_state="expanded"
)

# CSS Responsivo + Sofisticado
st.markdown("""
<style>
    .main {background-color: #0E1117;}
    h1, h2, h3 {color: #00FFAA;}
    .stPlotlyChart {border-radius: 12px;}
    .mobile-only {display: block;}
    @media (max-width: 768px) { .stColumns > div {width: 100% !important;} }
</style>
""", unsafe_allow_html=True)

st.title("💰 Dashboard Financeiro Premium")
st.caption("Sistema Autônomo • Dados salvos • Compatível com Celular")

# ====================== BANCO DE DADOS ======================
DB_FILE = "financeiro.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    conn.executescript('''
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
    ''')
    conn.commit()
    conn.close()

def carregar_transacoes():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM transacoes", conn)
    conn.close()
    return df

def carregar_investimentos():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM investimentos", conn)
    conn.close()
    return df

init_db()

# ====================== MENU LATERAL ====================
st.sidebar.header("📌 Menu Principal")
aba = st.sidebar.radio("Ir para:", [
    "➕ Nova Movimentação",
    "📊 Dashboard",
    "💼 Investimentos",
    "📋 Todas as Movimentações"
])

# ====================== FILTRO GLOBAL POR MÊS ======================
st.sidebar.header("🔍 Filtro por Mês")
mes_filtro = st.sidebar.selectbox(
    "Filtrar por mês:",
    ["Todos os meses", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
     "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
    index=0
)

# ====================== NOVA MOVIMENTAÇÃO ======================
if aba == "➕ Nova Movimentação":
    st.subheader("Adicionar Nova Movimentação")
    tipo_mov = st.selectbox("Tipo", ["Saída (Despesa)", "Entrada (Receita)"])
    
    with st.form("form_mov"):
        col1, col2 = st.columns(2)
        with col1:
            data = st.date_input("Data", value=date.today())
            descricao = st.text_input("Descrição")
            if tipo_mov == "Entrada (Receita)":
                categoria = st.selectbox("Categoria", ["Salário", "Rendas Extras", "Freelance", "Investimentos", "Reembolso", "Outro"])
            else:
                categoria = st.selectbox("Categoria", ["Mercado", "Necessidades", "Roupa", "Beleza", "Lazer", "Aluguel", "Contas", "Uber", "IFood", "Outro"])
        with col2:
            valor = st.number_input("Valor R$", value=0.0, step=0.01)
            cartao = st.text_input("Forma de Pagamento / Cartão")
        
        if st.form_submit_button("💾 Salvar"):
            valor_final = valor if tipo_mov == "Entrada (Receita)" else -valor
            conn = sqlite3.connect(DB_FILE)
            conn.execute("INSERT INTO transacoes (data, descricao, categoria, valor, tipo, cartao) VALUES (?,?,?,?,?,?)",
                         (data, descricao, categoria, valor_final, tipo_mov, cartao))
            conn.commit()
            conn.close()
            st.success("✅ Salvo com sucesso!")
            st.rerun()

# ====================== DASHBOARD ======================
elif aba == "📊 Dashboard":
    df = carregar_transacoes()
    
    # Aplicar filtro de mês
    if mes_filtro != "Todos os meses":
        df = df[df['data'].str.contains(mes_filtro[:3], na=False)]  # filtro simples por texto

    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("💰 Entradas", f"R$ {df[df['valor']>0]['valor'].sum():,.2f}")
    with col2: st.metric("💸 Saídas", f"R$ {abs(df[df['valor']<0]['valor'].sum()):,.2f}")
    with col3: st.metric("📊 Saldo", f"R$ {df['valor'].sum():,.2f}")
    with col4: st.metric("Total Mov.", len(df))

    tab1, tab2 = st.tabs(["📊 Gráficos", "📋 Tabela"])
    with tab1:
        col_a, col_b = st.columns(2)
        with col_a:
            fig = px.pie(df, names='categoria', values=abs(df['valor']), title="Gastos por Categoria")
            st.plotly_chart(fig, use_container_width=True)
        with col_b:
            fig2 = px.bar(df, x='categoria', y=abs(df['valor']), color='tipo', title="Entradas × Saídas")
            st.plotly_chart(fig2, use_container_width=True)

# ====================== INVESTIMENTOS ======================
elif aba == "💼 Investimentos":
    st.subheader("💼 Registro de Investimentos")
    
    with st.form("form_invest"):
        col1, col2 = st.columns(2)
        with col1:
            data_inv = st.date_input("Data do Investimento", value=date.today())
            tipo_inv = st.selectbox("Tipo", ["Renda Fixa", "Ações", "Fundos", "Cripto", "Tesouro Direto", "Poupança", "Outro"])
            valor_inv = st.number_input("Valor Investido R$", value=0.0, step=0.01)
        with col2:
            rentab = st.text_input("Rentabilidade Esperada (%)", "12.5")
            status = st.selectbox("Status", ["Em andamento", "Resgatado", "Vencido"])
            desc_inv = st.text_input("Descrição / Nome do Ativo")
        
        if st.form_submit_button("💾 Registrar Investimento"):
            conn = sqlite3.connect(DB_FILE)
            conn.execute("INSERT INTO investimentos (data, tipo, valor, rentabilidade, descricao, status) VALUES (?,?,?,?,?,?)",
                         (data_inv, tipo_inv, valor_inv, rentab, desc_inv, status))
            conn.commit()
            conn.close()
            st.success("✅ Investimento registrado!")
            st.rerun()

    # Mostrar investimentos cadastrados
    inv_df = carregar_investimentos()
    if len(inv_df) > 0:
        st.dataframe(inv_df, use_container_width=True)

# ====================== TODAS AS MOVIMENTAÇÕES ======================
elif aba == "📋 Todas as Movimentações":
    df = carregar_transacoes()
    if len(df) > 0:
        st.dataframe(df.sort_values('data', ascending=False), use_container_width=True, height=700)
    else:
        st.info("Nenhuma movimentação ainda.")

st.sidebar.caption("✅ Dados salvos automaticamente no SQLite\n📱 Totalmente compatível com celular")