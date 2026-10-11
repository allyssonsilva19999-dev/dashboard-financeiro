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

Planejador financeiro completo em um único arquivo HTML. Basta abrir no navegador: não usa servidor nem login, os dados ficam só no `localStorage` do aparelho e podem ser exportados/importados em JSON (backups da versão anterior continuam válidos). Funciona offline; só a fonte vem do Google Fonts, com fonte do sistema como alternativa.

- **Lançamentos**: receitas e despesas com categoria, data, forma de pagamento, situação (pago / a pagar), compras parceladas no crédito, observação, busca, filtros, edição e desfazer.
- **Importar extrato**: lê CSV (detecta separador, codificação e colunas dos principais bancos), OFX e PDF de extrato ou fatura, inclusive PDF com senha. Mostra uma prévia para conferir tipo e categoria, sugere a categoria (lembrando as que você já usou) e deixa de fora saldos, pagamentos de fatura e lançamentos repetidos. A leitura de PDF usa o pdf.js (Mozilla, Apache 2.0) em `vendor/pdfjs`, no próprio aparelho. PDFs escaneados (imagem) não têm texto para ler: use o CSV ou OFX do banco.
- **Orçamento**: limite mensal por categoria, alertas a partir de 80% e comparação com a regra 50/30/20.
- **Contas fixas**: aluguel, luz, assinaturas; marca como paga no mês e avisa atrasos.
- **Metas**: valor, prazo e quanto guardar por mês, com previsão de conclusão.
- **Dívidas**: saldo, juros e parcela, simulação de quitação e ordem "maior juro primeiro" ou "menor saldo primeiro".
- **Investimentos**: patrimônio, rendimento e reserva de emergência comparada ao custo de vida.
- **Diagnóstico**: nota de saúde financeira, regra 50/30/20 e as mesmas sugestões da página Análise, mais as do planejamento.

Na primeira visita ele abre com dados de exemplo (marcados e removíveis em "Começar do zero"). Responsivo do celular ao desktop, com tema claro/escuro do sistema e animações que respeitam a preferência de movimento reduzido.

### Publicar e instalar no celular

O repositório já tem o necessário para virar app instalável (`manifest.webmanifest`, ícones, `sw.js` para funcionar offline e `index.html`, que abre o organizador):

1. No GitHub, abra **Settings → Pages**.
2. Em **Build and deployment**, escolha **Deploy from a branch**, selecione o branch (`main`, ou o branch desta PR para testar antes) e a pasta **/ (root)**. Salve.
3. Em um ou dois minutos o site fica em `https://allyssonsilva19999-dev.github.io/dashboard-financeiro/`.
4. No iPhone, abra esse endereço no **Safari**, toque em **Compartilhar → Adicionar à Tela de Início**. O app abre em tela cheia, com ícone próprio, e funciona sem internet.

No iPhone, o app da Tela de Início guarda os dados separado do Safari. Use sempre o ícone e faça **Ajustes → Exportar** de vez em quando como backup.

## Testes

```bash
pip install pytest
python -m pytest tests                  # lógica da análise (Python)
node --test tests/organizador.test.js   # lógica do organizador.html (JS)
```
