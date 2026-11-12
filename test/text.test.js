const assert = require("assert").strict;
const T = require("../text.js");

let failures = 0;
function test(name, run) {
  try {
    run();
    console.log(`✓ ${name}`);
  } catch (error) {
    failures++;
    console.error(`✗ ${name}`);
    console.error(error.message);
  }
}

const cebolla = [
  { qty: 1000, unit: "g", name: "Cebolla blanca" },
  { qty: 40, unit: "g", name: "Aceite de Oliva" },
  { qty: 5, unit: "g", name: "Sal" },
];

test("al escalar, la cantidad escrita en el paso se actualiza", () => {
  const out = T.rescaleLiteralQuantities("Agregar los 40 gramos de aceite y los 5 gramos de sal.", cebolla, 1.4);
  assert.equal(out, "Agregar los 56 gramos de aceite y los 7 gramos de sal.");
});

test("sin escalar no se toca ni un número", () => {
  const texto = "Agregar los 40 gramos de aceite y los 5 gramos de sal.";
  assert.equal(T.rescaleLiteralQuantities(texto, cebolla, 1), texto);
});

test("los tiempos y las temperaturas nunca se escalan", () => {
  const texto = "Programar 30 minutos, 105 °C, velocidad 1, con el cubilete del revés.";
  assert.equal(T.rescaleLiteralQuantities(texto, cebolla, 2), texto);
});

test("una cantidad que no es de ningún ingrediente se deja como está", () => {
  const texto = "Envasar en raciones de 150 g de producto terminado.";
  assert.equal(T.rescaleLiteralQuantities(texto, cebolla, 2), texto);
});

test("un número que coincide pero no es ese ingrediente no se toca", () => {
  // 40 g, sí, pero de harina: no está en la ficha.
  const texto = "Espolvorear 40 g de harina por encima.";
  assert.equal(T.rescaleLiteralQuantities(texto, cebolla, 2), texto);
});

test("la unidad del texto manda aunque la ficha use otra de la misma familia", () => {
  const out = T.rescaleLiteralQuantities("Pelar 1 kg de cebolla.", cebolla, 1.4);
  assert.equal(out, "Pelar 1,4 kg de cebolla.");
});

test("una unidad de otra familia no se confunde con la de la ficha", () => {
  const texto = "Añadir 1000 ml de cebolla."; // absurdo a propósito: ml vs g
  assert.equal(T.rescaleLiteralQuantities(texto, cebolla, 2), texto);
});

test("el plural del ingrediente también cuenta", () => {
  const out = T.rescaleLiteralQuantities("Pelar 1000 g de cebollas.", cebolla, 0.5);
  assert.equal(out, "Pelar 500 g de cebollas.");
});

test("al paso que solo nombra el ingrediente se le añade la cantidad del lote", () => {
  const out = T.stepTextWithQuantities("Pelar las cebollas y trocearlas.", cebolla, 1.4);
  assert.match(out, /1400 g de cebollas/);
});

test("un paso con la cantidad ya escrita no la recibe dos veces", () => {
  const out = T.stepTextWithQuantities("Agregar los 40 gramos de aceite.", cebolla, 1.4);
  assert.equal(out, "Agregar los 56 gramos de aceite.");
});

test("stepTextWithQuantities sin escala se comporta como antes", () => {
  const out = T.stepTextWithQuantities("Agregar los 40 gramos de aceite.", cebolla, 1);
  assert.equal(out, "Agregar los 40 gramos de aceite.");
});

test("los decimales salen con coma, como en el resto de la app", () => {
  const out = T.rescaleLiteralQuantities("Añadir 5 g de sal.", cebolla, 1.3);
  assert.equal(out, "Añadir 6,5 g de sal.");
});

const pasosB02 = [
  "Pelar las cebollas, cortarlas en trozos medianos y colocarlas en el vaso.",
  "Bajar con una espátula la cebolla de las paredes del vaso.",
  "Agregar los 40 gramos de aceite y los 5 gramos de sal.",
  "La cebolla debe reducir y quedar bien transparente.",
];

test("la cantidad de cada ingrediente se dice una sola vez en toda la etapa", () => {
  const out = T.stepTexts(pasosB02, cebolla, 1.4);
  assert.match(out[0], /1400 g de cebollas/);
  assert.equal(out[1], "Bajar con una espátula la cebolla de las paredes del vaso.");
  assert.equal(out[3], "La cebolla debe reducir y quedar bien transparente.");
});

test("si la ficha ya escribió la cantidad, no se repite después", () => {
  const pasos = [
    "Agregar los 40 gramos de aceite.",
    "Remover el aceite con la cebolla.",
  ];
  const out = T.stepTexts(pasos, cebolla, 1.4);
  assert.equal(out[0], "Agregar los 56 gramos de aceite.");
  assert.equal(out[1], "Remover el aceite con los 1400 g de cebolla.");
});

test("sin escalar también se dice una sola vez", () => {
  const out = T.stepTexts(pasosB02, cebolla, 1);
  assert.match(out[0], /1000 g de cebollas/);
  assert.equal(out[1], "Bajar con una espátula la cebolla de las paredes del vaso.");
  assert.equal(out[2], "Agregar los 40 gramos de aceite y los 5 gramos de sal.");
});

test("un \"del\" no se parte por la mitad al añadir la cantidad", () => {
  const limon = [{ qty: 4, unit: "g", name: "Limón" }];
  assert.equal(T.stepTexts(["Rallar la piel del limón."], limon, 1)[0], "Rallar la piel del limón (4 g).");
});

/* ---- unidades contables: siempre hacia arriba ---- */

test("scaleQty redondea hacia arriba las unidades contables", () => {
  assert.equal(T.scaleQty(2, "ud", 1.4), 3); // 2,8 huevos -> 3
  assert.equal(T.scaleQty(4, "u", 1.4), 6); // 5,6 -> 6
  assert.equal(T.scaleQty(3, "unidades", 2), 6); // exacto, no sube
  assert.equal(T.scaleQty(10, "ud", 1), 10);
});

test("scaleQty deja los gramos y los mililitros con su decimal", () => {
  assert.equal(T.scaleQty(1000, "g", 1.4), 1400);
  assert.equal(T.scaleQty(5, "g", 1.3), 6.5);
  assert.ok(Math.abs(T.scaleQty(250, "ml", 1.1) - 275) < 1e-9);
});

test("isCountableUnit distingue las unidades de las medidas", () => {
  assert.ok(T.isCountableUnit("ud"));
  assert.ok(T.isCountableUnit("u"));
  assert.ok(T.isCountableUnit("Unidades"));
  assert.ok(!T.isCountableUnit("g"));
  assert.ok(!T.isCountableUnit("ml"));
});

test("la cantidad que se añade al paso sale entera si es contable", () => {
  const huevos = [{ qty: 2, unit: "ud", name: "Huevos" }];
  assert.equal(T.stepTexts(["Cascar los huevos."], huevos, 1.4)[0], "Cascar las 3 ud de huevos.");
});

test("la cantidad escrita en la ficha también se redondea hacia arriba", () => {
  const huevos = [{ qty: 2, unit: "ud", name: "Huevos" }];
  assert.equal(
    T.stepTexts(["Cocer las 2 ud de huevos 10 minutos."], huevos, 1.4)[0],
    "Cocer las 3 ud de huevos 10 minutos."
  );
});

if (failures) {
  console.error(`\n${failures} fallo(s)`);
  process.exit(1);
}

/* ---------------- Fases de ingredientes ---------------- */

const croquetas = {
  stages: [
    {
      title: "MASA",
      phases: ["Masa"],
      ingredients: [
        { name: "Leche", qty: 1400, unit: "ml", phase: "Masa" },
        { name: "Harina", qty: 155, unit: "g", phase: "Masa" },
        { name: "Huevos", qty: 6, unit: "ud", phase: "Empanado" },
        { name: "Pan rallado", qty: 200, unit: "g", phase: "Empanado" },
        { name: "Aceite de girasol", qty: 329, unit: "ml", phase: "Fritura" },
        { name: "Perejil", qty: 2, unit: "g", phase: "" },
      ],
      steps: ["1"],
    },
    { title: "BOLEADO", phases: ["Empanado"], ingredients: [], steps: ["Preparar un bol con los huevos batidos."] },
    { title: "FRITURA", phases: "Fritura", ingredients: [], steps: ["2"] },
    { title: "REGENERACIÓN", phases: [], ingredients: [], steps: ["3"] },
  ],
};

test("cada etapa alista solo los ingredientes de sus fases", () => {
  const names = (i) => T.ingredientsForStage(croquetas, i).map((x) => x.name);
  // el perejil no lo reclama ninguna etapa: cae en la primera, no desaparece
  assert.deepEqual(names(0), ["Leche", "Harina", "Perejil"]);
  assert.deepEqual(names(1), ["Huevos", "Pan rallado"]);
  assert.deepEqual(names(2), ["Aceite de girasol"]); // declaración como texto
  assert.deepEqual(names(3), []); // etapa sin ingredientes propios
});

test("una ficha sin fases se comporta como siempre", () => {
  const sinFases = {
    stages: [
      { title: "A", ingredients: [{ name: "Agua", qty: 1, unit: "ml" }], steps: ["1"] },
      { title: "B", ingredients: [], steps: ["2"] },
    ],
  };
  assert.equal(T.recipeUsesPhases(sinFases), false);
  assert.equal(T.ingredientsForStage(sinFases, 0).length, 1);
  assert.deepEqual(T.ingredientsForStage(sinFases, 1), []);
});

test("las fases se comparan sin acentos ni mayúsculas", () => {
  assert.equal(T.phaseKey(" EMPANADO "), "empanado");
  assert.equal(T.phaseKey("Regeneración"), T.phaseKey("REGENERACION"));
  const r = {
    stages: [
      { title: "A", phases: ["regeneracion"], ingredients: [{ name: "X", qty: 1, unit: "g", phase: "Regeneración" }], steps: ["1"] },
    ],
  };
  assert.equal(T.ingredientsForStage(r, 0).length, 1);
});

test("la tabla se agrupa por fases en el orden de la ficha", () => {
  const groups = T.groupByPhase(croquetas.stages[0].ingredients);
  assert.deepEqual(groups.map((g) => g.phase), ["Masa", "Empanado", "Fritura", ""]);
  assert.deepEqual(groups[1].items.map((x) => x.index), [2, 3]);
  // el índice tiene que ser el de la tabla sin agrupar: es el que escala
  assert.equal(groups[2].items[0].ingredient.name, "Aceite de girasol");
  assert.equal(groups[2].items[0].index, 4);
});

test("los pasos de una etapa posterior ya dicen la cantidad", () => {
  const stage = croquetas.stages[1];
  const out = T.stepTexts(stage.steps, T.ingredientsForStage(croquetas, 1), 1);
  assert.match(out[0], /6 ud de huevos/);
});

test("al escalar, la etapa posterior escala con la receta", () => {
  const out = T.stepTexts(croquetas.stages[1].steps, T.ingredientsForStage(croquetas, 1), 1.5);
  assert.match(out[0], /9 ud de huevos/);
});

test("la cantidad pegada al ingrediente se escala y no se duplica", () => {
  const ing = [
    { name: "Huevos", qty: 4, unit: "ud" },
    { name: "Pan rallado crujiente", qty: 300, unit: "g" },
  ];
  const paso = "Una bandeja con 4 huevos (4 ud) batidos y otra con 300 g de pan rallado.";
  // sin escalar sobra el paréntesis: el número de delante ya lo dice
  assert.equal(
    T.stepTexts([paso], ing, 1)[0],
    "Una bandeja con 4 huevos batidos y otra con 300 g de pan rallado."
  );
  // al escalar se reescribe el número escrito, sin contradecirse
  const x2 = T.stepTexts([paso], ing, 2)[0];
  assert.match(x2, /8 huevos batidos/);
  assert.doesNotMatch(x2, /\(/);
  assert.match(x2, /600 g de pan rallado/);
});

test("una cantidad parcial no se convierte en el lote entero", () => {
  const ing = [{ name: "Huevos", qty: 4, unit: "ud" }];
  // "2 huevos" no es la cantidad de la ficha: ni se escala ni se le añade el lote
  assert.equal(T.stepTexts(["Añadir 2 huevos"], ing, 2)[0], "Añadir 2 huevos");
  // la mención sin número sí recibe la cantidad
  assert.match(T.stepTexts(["Añadir los huevos"], ing, 2)[0], /8 ud de huevos/);
});

test('"del" cuenta como "de" al reescribir cantidades', () => {
  const ing = [{ name: "Aceite de Oliva", qty: 67, unit: "g" }];
  const paso = "Incorporar 67 g del aceite del sofrito reservado.";
  // sin escalar no se toca y, sobre todo, no se le añade la cantidad otra vez
  assert.equal(T.stepTexts([paso], ing, 1)[0], paso);
  assert.equal(T.stepTexts([paso], ing, 2)[0], "Incorporar 134 g del aceite del sofrito reservado.");
});

test("una porción no se confunde con el lote", () => {
  const jamon = [{ name: "Jamón a dados", qty: 225, unit: "g" }];
  // la ficha ya dice cuánto pesa esa tarrina: no se le añade el lote delante
  assert.equal(
    T.stepTexts(["Triturar 1 tarrina de jamón (75 g) con Turbo 3 s."], jamon, 1)[0],
    "Triturar 1 tarrina de jamón (75 g) con Turbo 3 s."
  );
  // una medida de cocina que no es la de la ficha tampoco atrae el lote
  const aceite = [{ name: "Aceite de girasol", qty: 329, unit: "ml" }];
  assert.equal(
    T.stepTexts(["Poner la freidora con 3 L de aceite de girasol a 185 °C."], aceite, 1)[0],
    "Poner la freidora con 3 L de aceite de girasol a 185 °C."
  );
  const pan = [{ name: "Pan de masa madre", qty: 100, unit: "g" }];
  assert.equal(
    T.stepTexts(["Poner las 2 rebanadas de pan de masa madre en la tostadora."], pan, 1)[0],
    "Poner las 2 rebanadas de pan de masa madre en la tostadora."
  );
  // pero una mención limpia sigue recibiendo su cantidad
  assert.match(T.stepTexts(["Añadir el aceite y calentar."], aceite, 1)[0], /329 ml de aceite/);
});

test("la cantidad escrita con una medida de cocina también escala", () => {
  const gel = [{ name: "Gelatina", qty: 11, unit: "ud" }];
  const paso = "Hidratar las 11 hojas de gelatina en agua fría.";
  assert.equal(T.stepTexts([paso], gel, 1)[0], paso);
  assert.match(T.stepTexts([paso], gel, 2)[0], /22 hojas de gelatina/);
  // contable: 16,5 hojas no existen
  assert.match(T.stepTexts([paso], gel, 1.5)[0], /17 hojas de gelatina/);
  // si el número no es el de la ficha es una porción: no se toca
  const jamon = [{ name: "Jamón a dados", qty: 225, unit: "g" }];
  assert.equal(
    T.stepTexts(["Triturar 1 tarrina de jamón con Turbo."], jamon, 2)[0],
    "Triturar 1 tarrina de jamón con Turbo."
  );
});
