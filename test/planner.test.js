const assert = require("assert").strict;
const fs = require("fs");
const path = require("path");
const planner = require("../planner.js");

function test(name, run) {
  try {
    run();
    console.log(`✓ ${name}`);
  } catch (error) {
    console.error(`✗ ${name}`);
    throw error;
  }
}

const recipes = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "plan-recipes.json"), "utf8")).recipes;
const b01 = recipes.find((r) => r.code === "B01");
const b02 = recipes.find((r) => r.code === "B02");
const ids = [b01.id, b02.id];

function overlaps(a, b) {
  return a.start < b.end && b.start < a.end;
}

test("un plan con B01 y B02 coloca todos los pasos con datos de planificación", () => {
  const plan = planner.buildPlan(recipes, ids, { people: 2 });
  assert.equal(plan.ok, true);
  assert.equal(plan.tasks.length, b01.plan.steps.length + b02.plan.steps.length);
  plan.tasks.forEach((t) => {
    assert.ok(Number.isFinite(t.start) && Number.isFinite(t.end) && t.end >= t.start);
    assert.ok(t.text.length > 0);
  });
});

test("cada paso empieza después de que acabe el anterior de su receta", () => {
  const plan = planner.buildPlan(recipes, ids, { people: 2 });
  [b01, b02].forEach((r) => {
    const mine = plan.tasks.filter((t) => t.recipeId === r.id).sort((a, b) => a.block - b.block || a.step - b.step);
    for (let i = 1; i < mine.length; i++) {
      assert.ok(mine[i].start >= mine[i - 1].end, `${r.code} paso ${mine[i].step} empieza antes de acabar el anterior`);
    }
  });
});

test("una persona nunca hace dos pasos a la vez", () => {
  const plan = planner.buildPlan(recipes, ids, { people: 2 });
  for (let p = 0; p < 2; p++) {
    const mine = plan.tasks.filter((t) => t.person === p);
    for (let i = 0; i < mine.length; i++) {
      for (let j = i + 1; j < mine.length; j++) {
        const a = { start: mine[i].start, end: mine[i].start + mine[i].personMin };
        const b = { start: mine[j].start, end: mine[j].start + mine[j].personMin };
        assert.ok(!overlaps(a, b), `persona ${p + 1}: solapan ${mine[i].code}.${mine[i].step} y ${mine[j].code}.${mine[j].step}`);
      }
    }
  }
});

test("la Thermomix grande no atiende dos pasos de recetas distintas a la vez", () => {
  const plan = planner.buildPlan(recipes, ids, { people: 3 });
  const g = plan.spans.filter((s) => s.key === "TMX_G");
  for (let i = 0; i < g.length; i++) {
    for (let j = i + 1; j < g.length; j++) assert.ok(!overlaps(g[i], g[j]), "dos reservas solapadas en la Thermomix grande");
  }
});

test("B01 usa la Thermomix grande y B02 (cualquiera) usa la pequeña si la grande está ocupada", () => {
  const plan = planner.buildPlan(recipes, ids, { people: 2 });
  const keys = new Set(plan.spans.map((s) => s.key));
  assert.ok(keys.has("TMX_G"));
  assert.ok(keys.has("TMX_P"));
});

test("con 2 personas se acaba antes (o igual) que con 1", () => {
  const one = planner.buildPlan(recipes, ids, { people: 1 });
  const two = planner.buildPlan(recipes, ids, { people: 2 });
  assert.ok(two.end <= one.end);
  assert.ok(two.end < one.end, "con dos personas se esperaba ganar tiempo en este lote");
});

test("las personas solo trabajan dentro de la jornada", () => {
  const plan = planner.buildPlan(recipes, ids, { people: 2, startMin: 9 * 60, shiftMin: 4 * 60 });
  plan.tasks
    .filter((t) => t.person !== null)
    .forEach((t) => {
      const tm = t.start % planner.DAY;
      assert.ok(tm >= 9 * 60, `${t.code}.${t.step} empieza antes de la jornada`);
      assert.ok(tm + t.personMin <= 13 * 60, `${t.code}.${t.step} acaba fuera de la jornada`);
    });
});

test("regla del stock: una ficha previa fuera de la selección se da por hecha", () => {
  const withReq = JSON.parse(JSON.stringify(recipes));
  withReq.find((r) => r.code === "B02").plan.requires = ["B01", "B03"];
  const onlyB02 = planner.buildPlan(withReq, [b02.id], { people: 1 });
  assert.deepEqual(onlyB02.stock.map((s) => s.code).sort(), ["B01", "B03"]);
  const both = planner.buildPlan(withReq, ids, { people: 2 });
  // B01 está en la selección: B02 no empieza hasta que B01 termina
  const endB01 = Math.max(...both.tasks.filter((t) => t.code === "B01").map((t) => t.end));
  const startB02 = Math.min(...both.tasks.filter((t) => t.code === "B02").map((t) => t.start));
  assert.ok(startB02 >= endB01);
  assert.deepEqual(both.stock.map((s) => s.code), ["B03"]);
});

test("los pasos sin tiempo se suponen y quedan anotados", () => {
  const broken = JSON.parse(JSON.stringify(recipes));
  const step = broken.find((r) => r.code === "B01").plan.steps.find((s) => s.t === "Máquina" && s.mm === 45);
  step.mm = null;
  const plan = planner.buildPlan(broken, [b01.id], { people: 1 });
  assert.equal(plan.assumptions.length, 1);
  assert.equal(plan.assumptions[0].code, "B01");
});

test("una espera larga abre el día siguiente para los pasos marcados con Día 2", () => {
  const multi = JSON.parse(JSON.stringify(recipes));
  const r = multi.find((x) => x.code === "B01");
  const last = r.plan.steps[r.plan.steps.length - 1];
  last.d = 2;
  const plan = planner.buildPlan(multi, [r.id], { people: 1 });
  const t = plan.tasks.find((x) => x.step === last.s);
  assert.ok(t.start >= planner.DAY + planner.DEFAULTS.startMin, "el paso del día 2 no empieza el día siguiente");
  assert.equal(plan.days, 2);
});

test("un paso POR TANDA se repite según el rendimiento y ENTRE TANDAS una menos", () => {
  const tandas = JSON.parse(JSON.stringify(recipes));
  const r = tandas.find((x) => x.code === "B01");
  r.yieldAmount = 50;
  const maq = r.plan.steps.find((s) => s.t === "Máquina" && s.mm === 45);
  maq.rp = "POR TANDA";
  maq.ub = 10;
  const plan = planner.buildPlan(tandas, [r.id], { people: 1 });
  const t = plan.tasks.find((x) => x.step === maq.s);
  assert.equal(t.totalMin, 45 * 5);
});

test("las recetas sin datos de planificación no entran y se avisan", () => {
  const sin = recipes.concat([{ id: "sin-plan", code: "B99", name: "SIN PLAN", stages: [{ title: "x", ingredients: [], steps: ["a"] }] }]);
  const plan = planner.buildPlan(sin, ids.concat(["sin-plan"]), { people: 1 });
  assert.deepEqual(plan.skipped.map((s) => s.code), ["B99"]);
  assert.equal(plan.tasks.length, b01.plan.steps.length + b02.plan.steps.length);
});

test("el formato de horas y duraciones es legible", () => {
  assert.equal(planner.fmtClock(8 * 60 + 5), "08:05");
  assert.equal(planner.fmtClock(planner.DAY + 9 * 60), "09:00");
  assert.equal(planner.dayOf(planner.DAY + 1), 2);
  assert.equal(planner.fmtDuration(0.5), "30 s");
  assert.equal(planner.fmtDuration(45), "45 min");
  assert.equal(planner.fmtDuration(125), "2 h 5 min");
  assert.equal(planner.fmtDuration(1440), "1 día");
});

const withB03 = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "plan-recipes.json"), "utf8")).recipes;
const b03 = withB03.find((r) => r.code === "B03");

test("B03 (salmorreta) forma parte de los datos de prueba con su plan", () => {
  assert.ok(b03 && planner.hasPlan(b03));
});

test("añadir personas nunca retrasa el final del plan", () => {
  const all = withB03.map((r) => r.id);
  let prev = Infinity;
  for (let people = 1; people <= 6; people++) {
    const plan = planner.buildPlan(withB03, all, { people });
    assert.equal(plan.ok, true);
    assert.ok(plan.end <= prev, `con ${people} personas termina más tarde (${plan.end}) que con ${people - 1} (${prev})`);
    prev = plan.end;
  }
});

test("si B01 y B03 compiten por la Thermomix grande, gana la que más tarda en total", () => {
  const plan = planner.buildPlan(withB03, [b01.id, b03.id], { people: 3 });
  // Óptimo con una sola Thermomix grande: B03 primero y B01 después (10:44).
  assert.ok(plan.end <= 8 * 60 + 164, `el plan termina a las ${planner.fmtClock(plan.end)}, se esperaba 10:44 o antes`);
});
