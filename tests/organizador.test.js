// Testa a lógica pura de organizador.html: node --test tests/
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const html = fs.readFileSync(path.join(__dirname, "..", "organizador.html"), "utf8");
const bloco = html.match(/LOGICA:INICIO ====[^\n]*\n([\s\S]*?)\/\* ==== LOGICA:FIM/);
assert.ok(bloco, "bloco LOGICA não encontrado em organizador.html");
const L = new Function(`${bloco[1]}
  return { analisar, parseValor, dataValida, grupoDaCategoria, validarImportacao, mesesDisponiveis, brl };`)();

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
