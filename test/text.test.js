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
