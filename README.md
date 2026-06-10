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
