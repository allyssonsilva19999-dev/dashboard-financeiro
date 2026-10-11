# Dashboard Financeiro

Aplicativo Streamlit para registrar movimentações, investimentos, importar planilhas e gerar relatórios.

## Persistência por usuário

O aplicativo usa Supabase para manter os registros após atualizações e reinicializações do Streamlit Cloud. Cada pessoa cria uma conta e acessa somente os próprios dados.

1. Crie um projeto no Supabase.
2. Abra o SQL Editor e execute o conteúdo de `supabase_setup.sql`.
3. No Streamlit Cloud, abra **Manage app > Settings > Secrets**.
4. Adicione:

```toml
[supabase]
url = "https://SEU-PROJETO.supabase.co"
anon_key = "SUA_CHAVE_ANON_PUBLICA"
```

Sem esses Secrets, o aplicativo usa SQLite apenas para desenvolvimento local. Dados locais não são permanentes no Streamlit Cloud.

## Análise financeira (50/30/20)

A página **Análise** do menu diagnostica o mês escolhido: taxa de poupança, alerta de déficit, categoria de maior peso, distribuição Necessidades / Desejos / Poupança comparada à regra 50/30/20 (com a variação 60/25/15 como referência para o Brasil), gastos fora do padrão e sugestões práticas. A lógica fica em `analise_financeira.py`; a classificação das categorias nesses grupos é uma heurística, ajustável nas constantes do módulo.

## Organizador standalone (`organizador.html`)

Versão independente do diagnóstico, em um único arquivo HTML. Basta abrir no navegador: não usa servidor nem login, os dados ficam só no `localStorage` do aparelho e podem ser exportados/importados em JSON. Tem registro de transações, gráficos (Chart.js via CDN; o resto funciona offline), filtro por período, desfazer exclusão e as mesmas análises e sugestões da página Análise.

## Testes

```bash
pip install pytest
python -m pytest tests                  # lógica da análise (Python)
node --test tests/organizador.test.js   # lógica do organizador.html (JS)
```
