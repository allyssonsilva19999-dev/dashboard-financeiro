import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from analise_financeira import (  # noqa: E402
    DESEJOS,
    NECESSIDADES,
    POUPANCA,
    analisar,
    grupo_da_categoria,
    meses_disponiveis,
)


def mov(data, categoria, valor):
    return {"data": data, "categoria": categoria, "valor": valor}


def titulos(a):
    return [s["titulo"] for s in a["sugestoes"]]


def test_classificacao_de_grupos():
    assert grupo_da_categoria("Energia") == NECESSIDADES
    assert grupo_da_categoria("Cartão de crédito") == NECESSIDADES
    assert grupo_da_categoria("Streaming") == DESEJOS
    assert grupo_da_categoria("Investimento / Reserva") == POUPANCA
    assert grupo_da_categoria("Categoria nova") == DESEJOS
    assert grupo_da_categoria(None) == DESEJOS


def test_mes_saudavel_50_30_20():
    a = analisar([
        mov("2026-05-05", "Salário", 5000),
        mov("2026-05-06", "Moradia / Aluguel", -1500),
        mov("2026-05-07", "Mercado / Feira", -1000),
        mov("2026-05-10", "Lazer / Entretenimento", -1000),
        mov("2026-05-11", "Streaming", -500),
    ])
    assert a["mes"] == "2026-05"
    assert a["receitas"] == 5000
    assert a["despesas"] == 4000
    assert a["distribuicao"][NECESSIDADES]["pct"] == 50.0
    assert a["distribuicao"][DESEJOS]["pct"] == 30.0
    assert a["distribuicao"][POUPANCA]["pct"] == 20.0
    assert a["taxa_poupanca"] == 20.0
    assert not a["deficit"]
    assert any(s["nivel"] == "ok" for s in a["sugestoes"])
    assert "Automatize a poupança" in titulos(a)


def test_deficit_gera_alerta_e_aponta_maior_desejo():
    a = analisar([
        mov("2026-05-05", "Salário", 2000),
        mov("2026-05-06", "Moradia / Aluguel", -1200),
        mov("2026-05-07", "Delivery / Restaurante", -600),
        mov("2026-05-08", "Viagem", -700),
    ])
    assert a["deficit"]
    assert a["saldo"] == -500
    assert a["taxa_poupanca"] < 0
    assert a["distribuicao"][POUPANCA]["valor"] == 0
    primeira = a["sugestoes"][0]
    assert primeira["nivel"] == "alerta" and primeira["titulo"] == "Déficit no mês"
    assert "Viagem" in primeira["texto"]
    assert "Automatize a poupança" not in titulos(a)


def test_investimento_conta_como_poupanca_e_nao_como_despesa():
    a = analisar([
        mov("2026-05-05", "Salário", 4000),
        mov("2026-05-06", "Moradia / Aluguel", -1600),
        mov("2026-05-07", "Investimento / Reserva", -800),
    ])
    assert a["despesas"] == 1600
    assert a["investido"] == 800
    assert a["taxa_poupanca"] == 60.0
    assert all(r["categoria"] != "Investimento / Reserva" for r in a["ranking"])


def test_necessidades_acima_de_60_sugere_variacao_brasil():
    a = analisar([
        mov("2026-05-05", "Salário", 3000),
        mov("2026-05-06", "Moradia / Aluguel", -1500),
        mov("2026-05-07", "Mercado / Feira", -500),
    ])
    assert a["distribuicao"][NECESSIDADES]["pct"] > 60
    assert any("60/25/15" in s["texto"] for s in a["sugestoes"])


def test_moradia_e_divida_acima_do_limite():
    a = analisar([
        mov("2026-05-05", "Salário", 3000),
        mov("2026-05-06", "Moradia / Aluguel", -1000),
        mov("2026-05-07", "Cartão de crédito", -700),
    ])
    t = titulos(a)
    assert any(x.startswith("Moradia consome") for x in t)
    assert any(x.startswith("Dívidas consomem") for x in t)


def test_anomalia_exige_historico_e_aumento_relevante():
    base = [mov(f"2026-0{m}-05", "Salário", 3000) for m in (3, 4, 5)]
    base += [mov("2026-03-10", "Delivery / Restaurante", -100),
             mov("2026-04-10", "Delivery / Restaurante", -120)]
    maio = base + [mov("2026-05-10", "Delivery / Restaurante", -400)]
    assert any("Delivery / Restaurante fora do padrão" in t for t in titulos(analisar(maio)))

    # Um único mês anterior não basta para falar em "padrão".
    pouco = [mov("2026-04-05", "Salário", 3000), mov("2026-04-10", "Streaming", -50),
             mov("2026-05-05", "Salário", 3000), mov("2026-05-10", "Streaming", -300)]
    assert not any("fora do padrão" in t for t in titulos(analisar(pouco)))


def test_mes_explicito_ignora_outros_meses():
    regs = [
        mov("2026-04-05", "Salário", 1000), mov("2026-04-06", "Energia", -100),
        mov("2026-05-05", "Salário", 2000), mov("2026-05-06", "Energia", -300),
    ]
    assert analisar(regs, "2026-04")["receitas"] == 1000
    assert analisar(regs)["receitas"] == 2000
    assert meses_disponiveis(regs) == ["2026-05", "2026-04"]


def test_sem_dados_e_sem_receita_nao_quebram():
    vazio = analisar([])
    assert not vazio["tem_dados"] and vazio["taxa_poupanca"] is None
    assert vazio["sugestoes"][0]["nivel"] == "dica"

    so_despesa = analisar([mov("2026-05-06", "Energia", -100)])
    assert so_despesa["sugestoes"][0]["titulo"] == "Sem receitas no mês"

    sujo = analisar([{"data": "lixo", "categoria": "x", "valor": "abc"}, {"data": None}])
    assert not sujo["tem_dados"]


def test_categorias_genericas_pedem_categorizacao():
    a = analisar([
        mov("2026-05-05", "Salário", 3000),
        mov("2026-05-06", "Energia", -300),
        mov("2026-05-07", "Outro", -300),
    ])
    assert "Muitas despesas sem categoria" in titulos(a)
