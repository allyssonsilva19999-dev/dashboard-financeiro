"""Análise financeira mensal: taxa de poupança, déficit, regra 50/30/20 e sugestões.

Sem dependências externas. Recebe registros no formato do app: dicts com
``data`` (AAAA-MM-DD), ``categoria`` e ``valor`` (entrada > 0, saída < 0).
"""

from __future__ import annotations

import math
from collections import defaultdict
from datetime import date, datetime

NECESSIDADES = "Necessidades"
DESEJOS = "Desejos"
POUPANCA = "Poupança"

# Alvos da regra 50/30/20 e a variação 60/25/15, mais realista para o Brasil.
ALVO_PADRAO = {NECESSIDADES: 50, DESEJOS: 30, POUPANCA: 20}
ALVO_BRASIL = {NECESSIDADES: 60, DESEJOS: 25, POUPANCA: 15}

CATEGORIAS_NECESSIDADES = {
    "Moradia / Aluguel", "Condomínio", "Energia", "Água", "Internet / Telefone", "Gás",
    "Mercado / Feira", "Transporte / Combustível", "Saúde / Farmácia", "Plano de saúde",
    "Educação / Cursos", "Casa / Manutenção", "Filhos / Família", "Impostos / Taxas",
    "Seguros", "Trabalho / Home office",
    # Pagamento de dívidas conta como necessidade: é obrigação já assumida.
    "Cartão de crédito", "Empréstimo / Financiamento", "Dívidas",
}
CATEGORIAS_DESEJOS = {
    "Delivery / Restaurante", "Uber / App", "Assinaturas", "Lazer / Entretenimento",
    "Streaming", "Roupas / Calçados", "Beleza / Cuidados pessoais", "Pet",
    "Móveis / Eletro", "Viagem", "Presentes", "Doações",
}
CATEGORIAS_POUPANCA = {"Investimento / Reserva"}
CATEGORIAS_DIVIDA = {"Cartão de crédito", "Empréstimo / Financiamento", "Dívidas"}
CATEGORIAS_MORADIA = {"Moradia / Aluguel", "Condomínio"}
CATEGORIAS_GENERICAS = {"Outro", "Outros", "Importado", "Sem categoria", ""}

LIMITE_MORADIA = 30  # % da renda
LIMITE_DIVIDA = 20  # % da renda
LIMITE_CONCENTRACAO = 35  # % das despesas numa única categoria
LIMITE_GENERICAS = 20  # % das despesas em "Outro"/sem categoria
ANOMALIA_FATOR = 1.5  # gasto do mês vs média dos meses anteriores
ANOMALIA_MINIMO = 50.0  # R$ de aumento mínimo para alertar


def grupo_da_categoria(categoria: str | None) -> str:
    """Classifica uma categoria de saída em necessidade, desejo ou poupança.

    Categorias genéricas ou desconhecidas caem em Desejos: sem informação,
    assume-se gasto não essencial (o diagnóstico pede para categorizar melhor).
    """
    nome = (categoria or "").strip()
    if nome in CATEGORIAS_POUPANCA:
        return POUPANCA
    if nome in CATEGORIAS_NECESSIDADES:
        return NECESSIDADES
    return DESEJOS


def _parse_data(valor) -> date | None:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    try:
        return datetime.strptime(str(valor)[:10], "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return None


def _valor(registro) -> float:
    try:
        return float(registro.get("valor") or 0)
    except (TypeError, ValueError):
        return 0.0


def meses_disponiveis(registros) -> list[str]:
    """Meses com movimentação, no formato AAAA-MM, do mais recente ao mais antigo."""
    meses = set()
    for r in registros:
        d = _parse_data(r.get("data"))
        if d:
            meses.add(f"{d.year:04d}-{d.month:02d}")
    return sorted(meses, reverse=True)


def _pct(parte: float, total: float) -> float:
    return round(parte / total * 100, 1) if total > 0 else 0.0


def analisar(registros, mes: str | None = None) -> dict:
    """Diagnóstico do mês ``mes`` (AAAA-MM). Sem ``mes``, usa o mais recente."""
    registros = list(registros)
    meses = meses_disponiveis(registros)
    if mes is None and meses:
        mes = meses[0]

    por_mes: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    receitas = 0.0
    for r in registros:
        d = _parse_data(r.get("data"))
        v = _valor(r)
        if not d or v == 0:
            continue
        chave = f"{d.year:04d}-{d.month:02d}"
        if v < 0:
            cat = (r.get("categoria") or "").strip() or "Sem categoria"
            por_mes[chave][cat] += abs(v)
        elif chave == mes:
            receitas += v

    despesas_cat = dict(por_mes.get(mes, {})) if mes else {}
    grupos = {NECESSIDADES: 0.0, DESEJOS: 0.0, POUPANCA: 0.0}
    for cat, v in despesas_cat.items():
        grupos[grupo_da_categoria(cat)] += v

    investido = grupos[POUPANCA]
    gasto = grupos[NECESSIDADES] + grupos[DESEJOS]  # despesas, sem investimentos
    saldo = receitas - gasto - investido
    # Poupança no sentido da regra: tudo que não foi consumido (sobra + investido).
    poupado = max(receitas - gasto, 0.0)
    taxa_poupanca = round((receitas - gasto) / receitas * 100, 1) if receitas > 0 else None

    ranking = sorted(
        ((c, v) for c, v in despesas_cat.items() if grupo_da_categoria(c) != POUPANCA),
        key=lambda x: x[1],
        reverse=True,
    )
    resultado = {
        "mes": mes,
        "tem_dados": bool(mes) and (receitas > 0 or bool(despesas_cat)),
        "receitas": round(receitas, 2),
        "despesas": round(gasto, 2),
        "investido": round(investido, 2),
        "saldo": round(saldo, 2),
        "deficit": gasto > receitas and (receitas > 0 or gasto > 0),
        "taxa_poupanca": taxa_poupanca,
        "distribuicao": {
            NECESSIDADES: {"valor": round(grupos[NECESSIDADES], 2), "pct": _pct(grupos[NECESSIDADES], receitas)},
            DESEJOS: {"valor": round(grupos[DESEJOS], 2), "pct": _pct(grupos[DESEJOS], receitas)},
            POUPANCA: {"valor": round(poupado, 2), "pct": _pct(poupado, receitas)},
        },
        "ranking": [
            {"categoria": c, "valor": round(v, 2), "pct_despesas": _pct(v, gasto), "pct_renda": _pct(v, receitas),
             "grupo": grupo_da_categoria(c)}
            for c, v in ranking
        ],
    }
    resultado["sugestoes"] = gerar_sugestoes(resultado, _historico_categorias(por_mes, mes))
    return resultado


def _historico_categorias(por_mes, mes) -> dict[str, list[float]]:
    """Gasto por categoria nos meses anteriores ao analisado (só meses com despesa)."""
    historico: dict[str, list[float]] = defaultdict(list)
    anteriores = [m for m in por_mes if mes and m < mes and por_mes[m]]
    for m in anteriores:
        for cat in por_mes[m]:
            historico[cat].append(por_mes[m][cat])
    # Meses anteriores em que a categoria não apareceu contam como zero.
    for cat in historico:
        historico[cat] += [0.0] * (len(anteriores) - len(historico[cat]))
    return historico


def arredondar(x: float) -> int:
    """Arredonda para o inteiro mais próximo, metades para cima (42,5 -> 43)."""
    return int(math.floor(x + 0.5))


def _brl(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _sug(nivel: str, titulo: str, texto: str) -> dict:
    return {"nivel": nivel, "titulo": titulo, "texto": texto}


def gerar_sugestoes(a: dict, historico: dict[str, list[float]] | None = None) -> list[dict]:
    """Sugestões ordenadas por gravidade: alerta > atencao > dica > ok."""
    if not a["tem_dados"]:
        return [_sug("dica", "Comece registrando", "Cadastre receitas e despesas do mês para receber o diagnóstico.")]

    s: list[dict] = []
    receitas, gasto = a["receitas"], a["despesas"]
    nec = a["distribuicao"][NECESSIDADES]["pct"]
    des = a["distribuicao"][DESEJOS]["pct"]
    ranking = a["ranking"]

    if receitas <= 0:
        s.append(_sug("alerta", "Sem receitas no mês",
                      f"Há {_brl(gasto)} em despesas e nenhuma receita registrada. "
                      "Lance os ganhos do mês para o diagnóstico fazer sentido."))
        return s

    if a["deficit"]:
        falta = gasto - receitas
        desejos_top = next((r for r in ranking if r["grupo"] == DESEJOS), None)
        corte = (f" Comece por {desejos_top['categoria']} ({_brl(desejos_top['valor'])})."
                 if desejos_top else " Revise as despesas fixas.")
        s.append(_sug("alerta", "Déficit no mês",
                      f"Você gastou {_brl(falta)} a mais do que ganhou.{corte}"))

    taxa = a["taxa_poupanca"]
    if taxa is not None and not a["deficit"]:
        if taxa < 10:
            s.append(_sug("atencao", f"Taxa de poupança baixa ({arredondar(taxa)}%)",
                          "Sobrou menos de 10% da renda. Corte um gasto de desejo e reserve o valor assim que receber."))
        elif taxa < 20:
            s.append(_sug("dica", f"Taxa de poupança razoável ({arredondar(taxa)}%)",
                          "Você está perto da meta de 20%. Um pequeno ajuste em desejos fecha a diferença."))
        else:
            s.append(_sug("ok", f"Boa taxa de poupança ({arredondar(taxa)}%)",
                          "Você guardou 20% ou mais da renda. Direcione para reserva de emergência e investimentos."))

    if nec > 60:
        s.append(_sug("alerta" if nec > 75 else "atencao", f"Necessidades pesadas ({arredondar(nec)}% da renda)",
                      "Acima de 60%, a regra 50/30/20 fica apertada. Use a variação 60/25/15 por ora "
                      "e renegocie contratos fixos (aluguel, planos, tarifas) para voltar perto de 50%."))
    elif nec > 50:
        s.append(_sug("dica", f"Necessidades um pouco acima do ideal ({arredondar(nec)}%)",
                      "O ideal é até 50%, mas até 60% é aceitável no Brasil (variação 60/25/15)."))

    if des > 30:
        topo = next((r for r in ranking if r["grupo"] == DESEJOS), None)
        extra = f" O maior é {topo['categoria']} ({_brl(topo['valor'])})." if topo else ""
        s.append(_sug("atencao", f"Desejos acima de 30% ({arredondar(des)}%)",
                      f"Lazer e consumo não essencial passaram do limite.{extra}"))

    if ranking:
        top = ranking[0]
        if top["pct_despesas"] >= LIMITE_CONCENTRACAO:
            s.append(_sug("atencao", f"Gasto concentrado em {top['categoria']}",
                          f"{arredondar(top['pct_despesas'])}% das despesas ({_brl(top['valor'])}) estão em uma só categoria."))

        moradia = sum(r["pct_renda"] for r in ranking if r["categoria"] in CATEGORIAS_MORADIA)
        if moradia > LIMITE_MORADIA:
            s.append(_sug("atencao", f"Moradia consome {arredondar(moradia)}% da renda",
                          f"O recomendado é até {LIMITE_MORADIA}%. Avalie renegociar aluguel ou condomínio."))

        divida = sum(r["pct_renda"] for r in ranking if r["categoria"] in CATEGORIAS_DIVIDA)
        if divida > LIMITE_DIVIDA:
            s.append(_sug("alerta", f"Dívidas consomem {arredondar(divida)}% da renda",
                          "Priorize quitar as dívidas de juros mais altos (cartão, cheque especial) "
                          "antes de investir, e tente renegociar."))

        genericas = sum(r["pct_despesas"] for r in ranking if r["categoria"] in CATEGORIAS_GENERICAS)
        if genericas > LIMITE_GENERICAS:
            s.append(_sug("dica", "Muitas despesas sem categoria",
                          f"{arredondar(genericas)}% das despesas estão em \"Outro\" ou sem categoria. "
                          "Categorizar melhor deixa o diagnóstico mais preciso."))

    for cat, valores in (historico or {}).items():
        atual = next((r["valor"] for r in ranking if r["categoria"] == cat), 0.0)
        media = sum(valores) / len(valores) if valores else 0.0
        if len(valores) >= 2 and media > 0 and atual >= media * ANOMALIA_FATOR and atual - media >= ANOMALIA_MINIMO:
            s.append(_sug("atencao", f"{cat} fora do padrão",
                          f"{_brl(atual)} neste mês contra média de {_brl(media)} nos anteriores "
                          f"(+{arredondar((atual / media - 1) * 100)}%)."))

    if not a["deficit"] and a["distribuicao"][POUPANCA]["valor"] > 0:
        s.append(_sug("dica", "Automatize a poupança",
                      "Agende a transferência para a reserva no dia em que o salário cair, "
                      "antes de gastar. O que sobra no fim do mês costuma sumir."))

    ordem = {"alerta": 0, "atencao": 1, "dica": 2, "ok": 3}
    s.sort(key=lambda x: ordem[x["nivel"]])
    return s
