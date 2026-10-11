// Testa a lógica pura de organizador.html: node --test tests/organizador.test.js
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const html = fs.readFileSync(path.join(__dirname, "..", "organizador.html"), "utf8");
const bloco = html.match(/LOGICA:INICIO ====[^\n]*\n([\s\S]*?)\/\* ==== LOGICA:FIM/);
const plano = html.match(/PLANO:INICIO ====[^\n]*\n([\s\S]*?)\/\* ==== PLANO:FIM/);
assert.ok(bloco, "bloco LOGICA não encontrado em organizador.html");
assert.ok(plano, "bloco PLANO não encontrado em organizador.html");
const importacao = html.match(/IMPORTACAO:INICIO ====[^\n]*\n([\s\S]*?)\/\* ==== IMPORTACAO:FIM/);
assert.ok(importacao, "bloco IMPORTACAO não encontrado em organizador.html");
const L = new Function(`${bloco[1]}
  return { analisar, parseValor, dataValida, grupoDaCategoria, validarImportacao, mesesDisponiveis, brl };`)();
const P = new Function(`${bloco[1]}\n${plano[1]}
  return { somarMeses, dataNoMes, difMeses, gerarParcelas, progressoOrcamentos, mediaPorCategoria, custoMensalMedio, planoMeta,
    simularDivida, ordenarDividas, pagarParcelaDivida, statusContas, calcularScore, sugestoesPlano, validarEstado };`)();

const mov = (data, tipo, categoria, reais) => ({ id: `${data}-${categoria}-${reais}`, data, tipo, categoria, centavos: Math.round(reais * 100) });
const rec = (data, reais) => mov(data, "receita", "Salário", reais);
const dsp = (data, cat, reais) => mov(data, "despesa", cat, reais);
const titulos = a => a.sugestoes.map(s => s.titulo);

test("parseValor aceita formatos brasileiros e rejeita lixo", () => {
  assert.strictEqual(L.parseValor("1.234,56"), 123456);
  assert.strictEqual(L.parseValor("1234,5"), 123450);
  assert.strictEqual(L.parseValor("1234.56"), 123456);
  assert.strictEqual(L.parseValor("R$ 12"), 1200);
  assert.strictEqual(L.parseValor("1.234.567"), 123456700);
  assert.strictEqual(L.parseValor("0"), null);
  assert.strictEqual(L.parseValor("-5"), null);
  assert.strictEqual(L.parseValor("abc"), null);
  assert.strictEqual(L.parseValor("1,234,5"), null);
  assert.strictEqual(L.parseValor(""), null);
  assert.strictEqual(L.parseValor("0,10") + L.parseValor("0,20"), 30); // sem erro de ponto flutuante
});

test("dataValida", () => {
  assert.ok(L.dataValida("2026-02-28"));
  assert.ok(!L.dataValida("2026-02-30"));
  assert.ok(!L.dataValida("26-02-01"));
  assert.ok(!L.dataValida(20260201));
});

test("classificação de grupos igual ao Python", () => {
  assert.strictEqual(L.grupoDaCategoria("Energia"), "Necessidades");
  assert.strictEqual(L.grupoDaCategoria("Cartão de crédito"), "Necessidades");
  assert.strictEqual(L.grupoDaCategoria("Streaming"), "Desejos");
  assert.strictEqual(L.grupoDaCategoria("Investimento / Reserva"), "Poupança");
  assert.strictEqual(L.grupoDaCategoria("Nova"), "Desejos");
  assert.strictEqual(L.grupoDaCategoria(null), "Desejos");
});

test("mês saudável 50/30/20", () => {
  const a = L.analisar([
    rec("2026-05-05", 5000), dsp("2026-05-06", "Moradia / Aluguel", 1500), dsp("2026-05-07", "Mercado / Feira", 1000),
    dsp("2026-05-10", "Lazer / Entretenimento", 1000), dsp("2026-05-11", "Streaming", 500),
  ], "2026-05");
  assert.strictEqual(a.distribuicao.Necessidades.pct, 50);
  assert.strictEqual(a.distribuicao.Desejos.pct, 30);
  assert.strictEqual(a.distribuicao["Poupança"].pct, 20);
  assert.strictEqual(a.taxaPoupanca, 20);
  assert.ok(!a.deficit);
  assert.ok(a.sugestoes.some(s => s.nivel === "ok"));
  assert.ok(titulos(a).includes("Automatize a poupança"));
});

test("déficit gera alerta e aponta o maior desejo", () => {
  const a = L.analisar([
    rec("2026-05-05", 2000), dsp("2026-05-06", "Moradia / Aluguel", 1200),
    dsp("2026-05-07", "Delivery / Restaurante", 600), dsp("2026-05-08", "Viagem", 700),
  ], "2026-05");
  assert.ok(a.deficit);
  assert.strictEqual(a.saldo, -50000);
  assert.strictEqual(a.distribuicao["Poupança"].valor, 0);
  assert.strictEqual(a.sugestoes[0].nivel, "alerta");
  assert.match(a.sugestoes[0].texto, /Viagem/);
  assert.ok(!titulos(a).includes("Automatize a poupança"));
});

test("investimento conta como poupança, não como despesa", () => {
  const a = L.analisar([rec("2026-05-05", 4000), dsp("2026-05-06", "Moradia / Aluguel", 1600),
    dsp("2026-05-07", "Investimento / Reserva", 800)], "2026-05");
  assert.strictEqual(a.despesas, 160000);
  assert.strictEqual(a.investido, 80000);
  assert.strictEqual(a.taxaPoupanca, 60);
  assert.ok(a.ranking.every(r => r.categoria !== "Investimento / Reserva"));
});

test("necessidades > 60% sugere 60/25/15; moradia e dívida acima do limite", () => {
  const a = L.analisar([rec("2026-05-05", 3000), dsp("2026-05-06", "Moradia / Aluguel", 1500), dsp("2026-05-07", "Mercado / Feira", 500)], "2026-05");
  assert.ok(a.sugestoes.some(s => s.texto.includes("60/25/15")));
  const b = L.analisar([rec("2026-05-05", 3000), dsp("2026-05-06", "Moradia / Aluguel", 1000), dsp("2026-05-07", "Cartão de crédito", 700)], "2026-05");
  assert.ok(titulos(b).some(t => t.startsWith("Moradia consome")));
  assert.ok(titulos(b).some(t => t.startsWith("Dívidas consomem")));
});

test("anomalia exige 2+ meses de histórico", () => {
  const base = [rec("2026-03-05", 3000), rec("2026-04-05", 3000), rec("2026-05-05", 3000),
    dsp("2026-03-10", "Delivery / Restaurante", 100), dsp("2026-04-10", "Delivery / Restaurante", 120)];
  const maio = base.concat(dsp("2026-05-10", "Delivery / Restaurante", 400));
  assert.ok(titulos(L.analisar(maio, "2026-05")).includes("Delivery / Restaurante fora do padrão"));
  const pouco = [rec("2026-04-05", 3000), dsp("2026-04-10", "Streaming", 50), rec("2026-05-05", 3000), dsp("2026-05-10", "Streaming", 300)];
  assert.ok(!titulos(L.analisar(pouco, "2026-05")).some(t => t.includes("fora do padrão")));
});

test("período 'all' soma tudo e não calcula anomalias", () => {
  const regs = [rec("2026-04-05", 1000), dsp("2026-04-06", "Energia", 100), rec("2026-05-05", 2000), dsp("2026-05-06", "Energia", 300)];
  assert.strictEqual(L.analisar(regs, "all").receitas, 300000);
  assert.strictEqual(L.analisar(regs, "2026-04").receitas, 100000);
  assert.deepStrictEqual(L.mesesDisponiveis(regs), ["2026-05", "2026-04"]);
});

test("sem dados / sem receita não quebram", () => {
  const v = L.analisar([], "all");
  assert.ok(!v.temDados);
  assert.strictEqual(v.taxaPoupanca, null);
  assert.strictEqual(v.sugestoes[0].nivel, "dica");
  const d = L.analisar([dsp("2026-05-06", "Energia", 100)], "2026-05");
  assert.strictEqual(d.sugestoes[0].titulo, "Sem receitas no período");
});

test("validarImportacao descarta dados inválidos e limita tamanhos", () => {
  assert.strictEqual(L.validarImportacao({ foo: 1 }), null);
  assert.strictEqual(L.validarImportacao("x"), null);
  const v = L.validarImportacao({ transacoes: [
    { id: "a", tipo: "receita", categoria: "Salário", valor: 100.5, data: "2026-05-01", descricao: "ok" },
    { id: "b", tipo: "despesa", categoria: "X".repeat(500), valor: -20, data: "2026-05-02", descricao: "D".repeat(500) },
    { tipo: "outro", valor: 1, data: "2026-05-01" },
    { tipo: "receita", valor: "10", data: "2026-05-01" },
    { tipo: "receita", valor: 10, data: "2026-13-01" },
    { tipo: "receita", valor: Infinity, data: "2026-05-01" },
    null,
  ] });
  assert.strictEqual(v.transacoes.length, 2);
  assert.strictEqual(v.ignoradas, 5);
  assert.strictEqual(v.transacoes[0].centavos, 10050);
  assert.strictEqual(v.transacoes[1].centavos, 2000);
  assert.strictEqual(v.transacoes[1].categoria.length, 60);
  assert.strictEqual(v.transacoes[1].descricao.length, 80);
  assert.strictEqual(L.validarImportacao([{ tipo: "receita", valor: 5, data: "2026-05-01" }]).transacoes.length, 1);
});

/* ---------------- Planejamento (bloco PLANO) ---------------- */
const tx = (data, tipo, categoria, reais, extra) => Object.assign({ id: `${data}-${categoria}-${reais}`, data, tipo, categoria, centavos: Math.round(reais * 100), status: "pago" }, extra);

test("datas: meses somados atravessam o ano e o dia é limitado ao mês", () => {
  assert.strictEqual(P.somarMeses("2026-11", 3), "2027-02");
  assert.strictEqual(P.somarMeses("2026-01", -1), "2025-12");
  assert.strictEqual(P.somarMeses("2026-03", -15), "2024-12");
  assert.strictEqual(P.dataNoMes("2026-02", 31), "2026-02-28");
  assert.strictEqual(P.dataNoMes("2028-02", 30), "2028-02-29");
  assert.strictEqual(P.difMeses("2026-10", "2027-06"), 8);
});

test("gerarParcelas divide os centavos, avança os meses e deixa as futuras a pagar", () => {
  const ps = P.gerarParcelas({ id: "x", tipo: "despesa", categoria: "Móveis / Eletro", centavos: 100000, data: "2026-01-31", descricao: "Geladeira", status: "pago" }, 3);
  assert.deepStrictEqual(ps.map(p => p.centavos), [33334, 33333, 33333]);
  assert.deepStrictEqual(ps.map(p => p.data), ["2026-01-31", "2026-02-28", "2026-03-31"]);
  assert.deepStrictEqual(ps.map(p => p.status), ["pago", "pendente", "pendente"]);
  assert.strictEqual(ps[2].descricao, "Geladeira (3/3)");
  assert.ok(ps.every(p => p.parcela.grupo === ps[0].parcela.grupo && p.parcela.total === 3));
  assert.strictEqual(new Set(ps.map(p => p.id)).size, 3);
  assert.strictEqual(P.gerarParcelas({ centavos: 500, data: "2026-01-01" }, 1).length, 1);
});

test("orçamento: estados ok, alerta (80%+) e estourado; média dos meses anteriores", () => {
  const txs = [tx("2026-05-02", "despesa", "Mercado / Feira", 850), tx("2026-05-03", "despesa", "Delivery / Restaurante", 320), tx("2026-05-04", "despesa", "Lazer / Entretenimento", 50),
    tx("2026-04-02", "despesa", "Mercado / Feira", 900), tx("2026-03-02", "despesa", "Mercado / Feira", 700)];
  const r = P.progressoOrcamentos({ "Mercado / Feira": 100000, "Delivery / Restaurante": 30000, "Lazer / Entretenimento": 20000, Pet: 0 }, txs, "2026-05");
  assert.deepStrictEqual(r.map(o => [o.categoria, o.estado]), [["Delivery / Restaurante", "estourado"], ["Mercado / Feira", "alerta"], ["Lazer / Entretenimento", "ok"]]);
  assert.strictEqual(r[0].restante, -2000);
  assert.deepStrictEqual(P.mediaPorCategoria(txs, "2026-05"), { "Mercado / Feira": 80000 });
  assert.strictEqual(P.custoMensalMedio(txs.concat(tx("2026-05-05", "despesa", "Investimento / Reserva", 999)), "2026-05"), Math.round((122000 + 90000 + 70000) / 3));
});

test("planoMeta: concluída, no ritmo, fora do ritmo, vencida e sem prazo", () => {
  const hoje = "2026-10-11";
  assert.strictEqual(P.planoMeta({ alvo: 1000, atual: 1000 }, hoje).status, "concluida");
  const a = P.planoMeta({ alvo: 600000, atual: 215000, prazo: "2027-06-28", aporte: 50000 }, hoje);
  assert.strictEqual(a.status, "no-ritmo"); assert.strictEqual(a.mesesRestantes, 8); assert.strictEqual(a.aporteNecessario, 48125); assert.strictEqual(a.previsao, "2027-06");
  assert.strictEqual(P.planoMeta({ alvo: 550000, atual: 90000, prazo: "2027-03-28", aporte: 30000 }, hoje).status, "atrasada");
  assert.strictEqual(P.planoMeta({ alvo: 1000, atual: 0, prazo: "2026-01-01", aporte: 0 }, hoje).status, "vencida");
  const s = P.planoMeta({ alvo: 1000, atual: 0, aporte: 300 }, hoje);
  assert.strictEqual(s.status, "sem-prazo"); assert.strictEqual(s.previsao, "2027-02");
});

test("dívidas: simulação, ordem avalanche/bola de neve e pagamento de parcela", () => {
  assert.deepStrictEqual(P.simularDivida({ saldo: 100000, juros: 0, parcela: 25000 }), { meses: 4, juros: 0 });
  assert.strictEqual(P.simularDivida({ saldo: 100000, juros: 10, parcela: 10000 }), null); // parcela = juros: nunca quita
  const sim = P.simularDivida({ saldo: 240000, juros: 12.5, parcela: 40000 });
  assert.ok(sim.meses > 6 && sim.juros > 0);
  const ds = [{ id: "a", saldo: 240000, juros: 12.5 }, { id: "b", saldo: 2180000, juros: 1.49 }, { id: "c", saldo: 50000, juros: 3 }, { id: "q", saldo: 0, juros: 20 }];
  assert.deepStrictEqual(P.ordenarDividas(ds, "avalanche").map(d => d.id), ["a", "c", "b"]);
  assert.deepStrictEqual(P.ordenarDividas(ds, "bola-de-neve").map(d => d.id), ["c", "a", "b"]);
  const r = P.pagarParcelaDivida({ saldo: 240000, juros: 12.5, parcela: 40000, restantes: 8 });
  assert.strictEqual(r.pago, 40000); assert.strictEqual(r.divida.saldo, 230000); assert.strictEqual(r.divida.restantes, 7);
  assert.strictEqual(P.pagarParcelaDivida({ saldo: 10000, juros: 0, parcela: 40000 }).pago, 10000);
});

test("contas fixas: paga, agendada, atrasada e aberta, com atrasadas primeiro", () => {
  const contas = [{ id: "luz", nome: "Luz", valor: 1, dia: 15 }, { id: "alu", nome: "Aluguel", valor: 1, dia: 5 }, { id: "net", nome: "Net", valor: 1, dia: 20 },
    { id: "gym", nome: "Gym", valor: 1, dia: 1 }, { id: "off", nome: "Pausada", valor: 1, dia: 1, ativa: false }];
  const txs = [tx("2026-10-15", "despesa", "Energia", 1, { contaId: "luz", status: "pendente" }), tx("2026-10-01", "despesa", "Lazer", 1, { contaId: "gym" }), tx("2026-09-05", "despesa", "Moradia", 1, { contaId: "alu" })];
  const r = P.statusContas(contas, txs, "2026-10", "2026-10-11");
  assert.deepStrictEqual(r.map(x => [x.conta.id, x.estado]), [["alu", "atrasada"], ["luz", "agendada"], ["net", "aberta"], ["gym", "paga"]]);
  assert.strictEqual(r[0].dias, -6);
});

test("calcularScore fica entre 0 e 100 e melhora com mais sobra e reserva", () => {
  const base = { receitas: 700000, despesas: 600000, reserva: 0, custoMensal: 600000, totalDividas: 0, progressoMetas: 0 };
  const pior = P.calcularScore(Object.assign({}, base, { despesas: 900000, totalDividas: 5000000 }));
  const melhor = P.calcularScore(Object.assign({}, base, { despesas: 400000, reserva: 1800000, progressoMetas: 80 }));
  assert.ok(pior >= 0 && melhor <= 100 && melhor > P.calcularScore(base) && P.calcularScore(base) > pior);
  assert.strictEqual(P.calcularScore({ receitas: 0, despesas: 0 }), 50);
});

test("sugestoesPlano aponta conta atrasada, orçamento estourado, juros altos, reserva e meta", () => {
  const s = P.sugestoesPlano({
    contas: [{ estado: "atrasada", vencimento: "2026-10-05", conta: { nome: "Aluguel", valor: 180000 } }],
    orcamentos: [{ categoria: "Delivery / Restaurante", estado: "estourado", gasto: 32000, limite: 30000, pct: 106.7, restante: -2000 }],
    dividas: [{ nome: "Cartão", saldo: 240000, juros: 12.5, parcela: 40000 }, { nome: "Agiota", saldo: 100000, juros: 20, parcela: 10000 }],
    custoMensal: 500000, reserva: 500000, reservaAlvo: 3000000, mesesReserva: 6,
    metas: [{ m: { nome: "Notebook", prazo: "2027-03-28", aporte: 30000 }, p: { status: "atrasada", aporteNecessario: 92000 } }], temDespesas: true,
  });
  const t = s.map(x => x.titulo);
  assert.ok(t.includes("Conta atrasada: Aluguel"));
  assert.ok(t.includes("Orçamento de Delivery / Restaurante estourado"));
  assert.ok(t.includes("Juros altos em Cartão"));
  assert.ok(t.includes("Agiota: a parcela não cobre os juros"));
  assert.ok(t.some(x => x.startsWith("Reserva cobre 1 mês")));
  assert.ok(t.includes('Meta "Notebook" fora do ritmo'));
});

test("validarEstado aceita backup v1 e limpa campos do v2", () => {
  const v1 = P.validarEstado({ versao: 1, transacoes: [{ id: "a", tipo: "receita", categoria: "Salário", valor: 10, data: "2026-05-01" }] });
  assert.strictEqual(v1.transacoes.length, 1);
  assert.strictEqual(v1.transacoes[0].status, "pago");
  assert.strictEqual(v1.perfil.mesesReserva, 6);
  const v2 = P.validarEstado({
    versao: 2,
    transacoes: [{ id: "t", tipo: "despesa", categoria: "X", valor: 30, data: "2026-05-01", status: "pendente", pagamento: "P".repeat(99), parcela: { grupo: "g", n: 5, total: 3 }, contaId: "c1", exemplo: true }],
    orcamentos: { "Mercado / Feira": 500.5, Ruim: -3, Texto: "10" },
    metas: [{ nome: "Viagem", alvo: 1000, atual: -5, prazo: "2027-02-30" }, { nome: "Sem alvo", alvo: 0 }],
    dividas: [{ nome: "Cartão", saldo: 100, juros: 400, parcela: 10, dia: 99 }],
    contas: [{ nome: "Luz", valor: 0 }, { nome: "Net", valor: 100, dia: 10, ativa: false }],
    investimentos: [{ nome: "X", tipo: "Golpe", aplicado: 10, atual: 12 }],
    perfil: { nome: "Ana", renda: 5000, mesesReserva: 99, exemplo: ["renda", "nome"] },
    exemploOrc: ["Mercado / Feira", 3, "Y".repeat(99)],
  });
  const t = v2.transacoes[0];
  assert.strictEqual(t.status, "pendente"); assert.strictEqual(t.pagamento.length, 30); assert.strictEqual(t.parcela, undefined);
  assert.strictEqual(t.contaId, "c1"); assert.strictEqual(t.exemplo, true);
  assert.deepStrictEqual(v2.orcamentos, { "Mercado / Feira": 50050 });
  assert.strictEqual(v2.metas.length, 1); assert.strictEqual(v2.metas[0].atual, 0); assert.strictEqual(v2.metas[0].prazo, "");
  assert.strictEqual(v2.dividas[0].juros, 100); assert.strictEqual(v2.dividas[0].dia, 31);
  assert.strictEqual(v2.dividas[0].categoria, "Dívidas"); // tipo ausente vira "Dívidas"
  assert.strictEqual(P.validarEstado({ transacoes: [], dividas: [{ nome: "C", categoria: "Cartão de crédito", saldo: 1, parcela: 1 }] }).dividas[0].categoria, "Cartão de crédito");
  assert.deepStrictEqual(v2.contas.map(c => [c.nome, c.ativa]), [["Net", false]]);
  assert.strictEqual(v2.investimentos[0].tipo, "Outro");
  assert.strictEqual(v2.perfil.mesesReserva, 24); assert.deepStrictEqual(v2.perfil.exemplo, ["renda"]); assert.strictEqual(v2.perfil.renda, 500000);
  assert.deepStrictEqual(v2.exemploOrc, ["Mercado / Feira"]);
  assert.strictEqual(P.validarEstado("lixo"), null);
});

test("conta fixa não aparece em meses anteriores ao cadastro", () => {
  const contas = [{ id: "a", nome: "Aluguel", valor: 1, dia: 5, criadaEm: "2026-10-01" }, { id: "b", nome: "Antiga", valor: 1, dia: 5 }];
  assert.deepStrictEqual(P.statusContas(contas, [], "2026-09", "2026-10-11").map(x => x.conta.id), ["b"]);
  assert.deepStrictEqual(P.statusContas(contas, [], "2026-10", "2026-10-11").map(x => x.conta.id).sort(), ["a", "b"]);
});

/* ---------------- Importação de extratos (bloco IMPORTACAO) ---------------- */
const I = new Function(`${bloco[1]}\n${importacao[1]}
  return { valorExtrato, dataExtrato, lerCSV, detectarColunas, linhasCSVParaLancamentos, lerOFX, lerLinhasExtrato, referenciaDoTexto, detectarModo, sugerirCategoria, prepararImportacao };`)();
const doCSV = (texto, ref = "2026-10") => { const l = I.lerCSV(texto), m = I.detectarColunas(l); return { mapa: m, itens: I.linhasCSVParaLancamentos(l, m, ref) }; };

test("valorExtrato entende os formatos de banco", () => {
  const casos = { "1.234,56": 123456, "-1.234,56": -123456, "1234.56": 123456, "R$ 12": 1200, "R$ -45,90": -4590, "- R$ 45,90": -4590, "(12,00)": -1200,
    "12,00-": -1200, "12,00 D": -1200, "12,00 C": 1200, "45,90C": 4590, "−3,50": -350, "+7,1": 710, "1.500": 150000, "-150.00": -15000, "0,99": 99 };
  for (const [t, v] of Object.entries(casos)) assert.strictEqual(I.valorExtrato(t), v, t);
  for (const t of ["", "abc", "12,3,4x", "R$", "D"]) assert.strictEqual(I.valorExtrato(t), null, t);
});

test("dataExtrato com e sem ano, numérica e por extenso", () => {
  assert.strictEqual(I.dataExtrato("05/10/2026", "2026-10"), "2026-10-05");
  assert.strictEqual(I.dataExtrato("05/10/26", "2026-10"), "2026-10-05");
  assert.strictEqual(I.dataExtrato("2026-10-05T10:00", "2026-10"), "2026-10-05");
  assert.strictEqual(I.dataExtrato("05.10.2026", "2026-10"), "2026-10-05");
  assert.strictEqual(I.dataExtrato("05/10", "2026-10"), "2026-10-05");
  assert.strictEqual(I.dataExtrato("20/12", "2026-01"), "2025-12-20"); // dezembro na fatura de janeiro
  assert.strictEqual(I.dataExtrato("12 OUT", "2026-10"), "2026-10-12");
  assert.strictEqual(I.dataExtrato("3 de março de 2026", "2026-10"), "2026-03-03");
  assert.strictEqual(I.dataExtrato("31/02/2026", "2026-10"), null);
  assert.strictEqual(I.dataExtrato("12 PIX", "2026-10"), null);
});

test("CSV da conta Nubank (vírgula, valores com sinal)", () => {
  const { mapa, itens } = doCSV("Data,Valor,Identificador,Descrição\n05/10/2026,-150.00,abc-1,Transferência enviada pelo Pix - JOAO\n06/10/2026,7200.00,abc-2,Transferência recebida - EMPRESA LTDA\n");
  assert.deepStrictEqual([mapa.data, mapa.valor, mapa.desc], [0, 1, [3]]);
  assert.deepStrictEqual(itens.map(i => [i.data, i.valor]), [["2026-10-05", -15000], ["2026-10-06", 720000]]);
  assert.strictEqual(I.detectarModo(itens, mapa), "conta");
});

test("CSV da fatura Nubank (positivos são gastos) vira modo fatura", () => {
  const { mapa, itens } = doCSV("date,category,title,amount\n2026-10-01,restaurante,Ifood *Ifood,45.90\n2026-10-02,supermercado,Carrefour,320.10\n2026-10-03,transporte,Uber *Trip,23.50\n2026-10-04,,Pagamento recebido,-900.00\n");
  assert.strictEqual(I.detectarModo(itens, mapa), "fatura");
  const prev = I.prepararImportacao(itens, "fatura", null, []);
  assert.deepStrictEqual(prev.map(p => [p.tipo, p.categoria, p.incluir]), [["despesa", "Delivery / Restaurante", true], ["despesa", "Mercado / Feira", true], ["despesa", "Uber / App", true], ["receita", "Outro", false]]);
});

test("CSV com ponto e vírgula, duas colunas de descrição e aspas (Inter)", () => {
  const { mapa, itens } = doCSV('Extrato Conta Corrente\nConta ;123\n\nData Lançamento;Histórico;Descrição;Valor;Saldo\n05/10/2026;Pix enviado ;"Maria; Silva";-1.234,56;10.000,00\n06/10/2026;Pix recebido;Fulano;500,00;10.500,00\n');
  assert.strictEqual(mapa.cabecalho, 2); // linhas vazias não contam
  assert.deepStrictEqual(itens.map(i => [i.descricao, i.valor]), [["Pix enviado - Maria; Silva", -123456], ["Pix recebido - Fulano", 50000]]);
});

test("CSV com colunas de crédito e débito (Bradesco, C6) e sem cabeçalho", () => {
  const b = doCSV("Data;Histórico;Docto.;Crédito (R$);Débito (R$);Saldo (R$)\n01/10/2026;SALDO ANTERIOR;;;;1.000,00\n02/10/2026;CONTA DE LUZ ENEL;123;;210,35;789,65\n03/10/2026;TED RECEBIDA;9;1.000,00;;1.789,65\n");
  assert.deepStrictEqual(b.itens.map(i => i.valor), [-21035, 100000]);
  const c6 = doCSV("Data Lançamento,Data Contábil,Título,Descrição,Entrada(R$),Saída(R$),Saldo do Dia(R$)\n05/10/2026,05/10/2026,Pix,Padaria,0,12.5,100\n");
  assert.deepStrictEqual(c6.itens.map(i => [i.descricao, i.valor]), [["Pix - Padaria", -1250]]);
  const sem = doCSV("05/10/2026;Mercado Bom Preço;-87,40\n06/10/2026;Salário;3500,00\n07/10/2026;Farmácia;-23,10\n");
  assert.deepStrictEqual(sem.itens.map(i => [i.descricao, i.valor]), [["Mercado Bom Preço", -8740], ["Salário", 350000], ["Farmácia", -2310]]);
});

test("OFX", () => {
  const ofx = "OFXHEADER:100\n<OFX><BANKTRANLIST><STMTTRN><TRNTYPE>DEBIT<DTPOSTED>20261005120000[-3:BRT]<TRNAMT>-150.00<FITID>1<MEMO>PIX ENVIADO JOAO</STMTTRN><STMTTRN><TRNTYPE>CREDIT<DTPOSTED>20261006<TRNAMT>7200.00<FITID>2<NAME>SALARIO EMPRESA</STMTTRN></BANKTRANLIST></OFX>";
  assert.deepStrictEqual(I.lerOFX(ofx).map(i => [i.data, i.valor, i.descricao]), [["2026-10-05", -15000, "PIX ENVIADO JOAO"], ["2026-10-06", 720000, "SALARIO EMPRESA"]]);
});

test("linhas de PDF: valor e saldo no fim, datas sem ano, linhas que não são lançamento", () => {
  const linhas = ["EXTRATO DE CONTA CORRENTE - período 01/10/2026 a 31/10/2026", "Data Histórico Valor Saldo",
    "05/10 PIX ENVIADO JOAO -150,00 1.234,56", "06/10 SALDO DO DIA 1.234,56", "07/10 PAGAMENTO CONTA LUZ ENEL 210,35 D 1.024,21",
    "08/10 SALARIO EMPRESA LTDA 7.200,00 C 8.224,21", "Total de saídas 360,35", "12 OUT UBER *TRIP R$ 23,50"];
  const ref = I.referenciaDoTexto(linhas, "2027-01");
  assert.strictEqual(ref, "2026-10");
  const r = I.lerLinhasExtrato(linhas, ref);
  assert.deepStrictEqual(r.map(i => [i.data, i.descricao, i.valor]), [["2026-10-05", "PIX ENVIADO JOAO", -15000], ["2026-10-06", "SALDO DO DIA", 123456],
    ["2026-10-07", "PAGAMENTO CONTA LUZ ENEL", -21035], ["2026-10-08", "SALARIO EMPRESA LTDA", 720000], ["2026-10-12", "UBER *TRIP", 2350]]);
  const prev = I.prepararImportacao(r, "conta", null, []);
  assert.strictEqual(prev[1].incluir, false); // saldo
  assert.strictEqual(prev[2].categoria, "Energia");
  assert.strictEqual(prev[3].categoria, "Salário");
  assert.strictEqual(I.detectarModo(r, null, "FATURA DO CARTÃO VISA"), "fatura");
});

test("categoria: memória do usuário vem antes das palavras-chave; repetidos vêm desmarcados", () => {
  const mem = new Map([["uber *trip", "Transporte / Combustível"], ["pix enviado joao", "Salário"]]);
  assert.strictEqual(I.sugerirCategoria("UBER *TRIP", "despesa", mem), "Transporte / Combustível");
  assert.strictEqual(I.sugerirCategoria("Pix enviado Joao", "despesa", mem), "Outro"); // lembrada é de receita: ignora
  assert.strictEqual(I.sugerirCategoria("NETFLIX.COM", "despesa"), "Streaming");
  assert.strictEqual(I.sugerirCategoria("Mercado Livre", "despesa"), "Outro");
  assert.strictEqual(I.sugerirCategoria("Tim celular", "despesa"), "Internet / Telefone");
  assert.strictEqual(I.sugerirCategoria("Rendimento poupança", "receita"), "Rendimentos / Investimentos");
  const prev = I.prepararImportacao([{ data: "2026-10-05", valor: -15000, descricao: "PIX" }], "conta", null, [{ data: "2026-10-05", centavos: 15000, tipo: "despesa" }]);
  assert.deepStrictEqual([prev[0].incluir, prev[0].motivo], [false, "Já lançado"]);
});
