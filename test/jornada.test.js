/* Backend de la jornada compartida (Code.gs) probado en local.
   Apps Script no se puede ejecutar aquí, así que se cargan sus funciones con
   PropertiesService, LockService y ContentService de mentira. Así el contrato
   de JORNADA_COMPARTIDA.md se comprueba antes de desplegar nada. */
const assert = require("assert").strict;
const fs = require("fs");
const path = require("path");

function test(name, run) {
  try {
    run();
    console.log(`✓ ${name}`);
  } catch (error) {
    console.error(`✗ ${name}`);
    throw error;
  }
}

function loadBackend() {
  const src = fs.readFileSync(path.join(__dirname, "..", "Code.gs"), "utf8");
  const store = new Map();
  const props = {
    getProperty: (k) => (store.has(k) ? store.get(k) : null),
    setProperty: (k, v) => store.set(k, String(v)),
    deleteProperty: (k) => store.delete(k),
    getKeys: () => [...store.keys()],
  };
  let locks = 0;
  const api = new Function(
    "PropertiesService",
    "LockService",
    "ContentService",
    "CacheService",
    "DriveApp",
    "SpreadsheetApp",
    "Logger",
    src + "\nreturn { doGet: doGet, doPost: doPost };"
  )(
    { getScriptProperties: () => props },
    {
      getScriptLock: () => ({
        tryLock: () => {
          locks++;
          return true;
        },
        releaseLock: () => locks--,
      }),
    },
    {
      MimeType: { JSON: "json" },
      createTextOutput: (text) => ({ text, setMimeType: () => ({ text, getContent: () => text }) }),
    },
    { getScriptCache: () => ({ get: () => null, getAll: () => ({}), putAll: () => {} }) },
    {},
    {},
    { log: () => {} }
  );
  const get = (parameter) => JSON.parse(api.doGet({ parameter }).text);
  const post = (body) => JSON.parse(api.doPost({ postData: { contents: JSON.stringify(body) } }).text);
  return { get, post, store, locked: () => locks };
}

const PLAN = {
  v: 1,
  plannerVersion: 1,
  start: 480,
  end: 1043,
  tasks: [
    ["salmorreta|1|1", 480, 495, 0, null, "Thermomix grande"],
    ["salmorreta|1|2", 495, 530, 1, null, "Olla 25 L (1)"],
  ],
  spans: [["TMX_G", "Thermomix grande", "B03", 480, 495]],
};
const INPUTS = { selected: ["salmorreta"], people: 3, start: "08:00", shiftH: 8 };

function newJornada(api, extra) {
  return api.post(Object.assign({ action: "create", code: "1234", name: "Leo", inputs: INPUTS, plan: PLAN, plannerVersion: 1, dataStamp: "2026-11-12T10:00:00.000Z" }, extra || {}));
}

test("crear una jornada devuelve código, slot y hora del servidor", () => {
  const api = loadBackend();
  const res = newJornada(api);
  assert.equal(res.ok, true);
  assert.equal(res.code, "1234");
  assert.equal(res.planRev, 1);
  assert.equal(typeof res.slot, "number");
  assert.ok(Math.abs(res.now - Date.now()) < 5000, "el servidor debe decir su hora");
  assert.equal(api.locked(), 0, "el candado se suelta siempre");
});

test("un código ocupado no se sobreescribe sin querer", () => {
  const api = loadBackend();
  newJornada(api);
  const res = newJornada(api, { name: "Otro" });
  assert.equal(res.ok, false);
  assert.equal(res.taken, true);
});

test("sin código, el servidor inventa uno de cuatro dígitos", () => {
  const api = loadBackend();
  const res = api.post({ action: "create", name: "Leo", inputs: INPUTS, plan: PLAN });
  assert.equal(res.ok, true);
  assert.match(res.code, /^\d{4}$/);
});

test("unirse devuelve el plan entero, el estado y un slot propio", () => {
  const api = loadBackend();
  newJornada(api);
  const res = api.post({ action: "join", code: "1234", name: "Marta" });
  assert.equal(res.ok, true);
  assert.deepEqual(res.plan, PLAN);
  assert.equal(res.meta.inputs.people, 3);
  assert.equal(res.meta.dataStamp, "2026-11-12T10:00:00.000Z");
  assert.ok(res.live.people.some((p) => p.n === "Marta"));
  assert.notEqual(res.slot, 0, "Leo ya ocupaba el puesto 0");
});

test("el mismo nombre recupera su puesto al volver", () => {
  const api = loadBackend();
  const created = newJornada(api);
  const again = api.post({ action: "join", code: "1234", name: "Leo" });
  assert.equal(again.slot, created.slot);
});

test("unirse a un código que no existe lo dice, no revienta", () => {
  const api = loadBackend();
  const res = api.post({ action: "join", code: "9999", name: "Leo" });
  assert.equal(res.ok, false);
  assert.equal(res.missing, true);
});

test("un código que no son cuatro dígitos se rechaza", () => {
  const api = loadBackend();
  assert.equal(api.post({ action: "join", code: "12", name: "x" }).ok, false);
  assert.equal(api.get({ jornada: "abc" }).ok, false);
});

test("marcar hecho se ve desde otro dispositivo", () => {
  const api = loadBackend();
  newJornada(api);
  const patch = api.post({ action: "patch", code: "1234", slot: 1, ops: [{ op: "done", key: "salmorreta|1|2" }] });
  assert.equal(patch.ok, true);
  const seen = api.get({ jornada: "1234" });
  assert.deepEqual(seen.live.done, { salmorreta: ["1|2:1"] }, "queda apuntado el paso y quién lo hizo");
});

test("desmarcar quita el hecho y no deja rastro doble", () => {
  const api = loadBackend();
  newJornada(api);
  api.post({ action: "patch", code: "1234", slot: 1, ops: [{ op: "done", key: "a|1|1" }, { op: "done", key: "a|1|1" }] });
  let live = api.get({ jornada: "1234" }).live;
  assert.deepEqual(live.done, { a: ["1|1:1"] }, "marcar dos veces no duplica");
  api.post({ action: "patch", code: "1234", slot: 2, ops: [{ op: "undone", key: "a|1|1" }] });
  live = api.get({ jornada: "1234" }).live;
  assert.deepEqual(live.done, {});
});

test("coger una tarea ya cogida se rechaza en vez de robarla", () => {
  const api = loadBackend();
  newJornada(api);
  const first = api.post({ action: "patch", code: "1234", slot: 1, ops: [{ op: "claim", key: "a|1|1" }] });
  assert.deepEqual(first.rejected, []);
  const second = api.post({ action: "patch", code: "1234", slot: 2, ops: [{ op: "claim", key: "a|1|1" }] });
  assert.equal(second.rejected.length, 1);
  assert.equal(second.rejected[0].owner, 1);
  assert.deepEqual(api.get({ jornada: "1234" }).live.owner, { a: ["1|1:1"] });
});

test("quien la tiene puede soltarla y otro no", () => {
  const api = loadBackend();
  newJornada(api);
  api.post({ action: "patch", code: "1234", slot: 1, ops: [{ op: "claim", key: "a|1|1" }] });
  api.post({ action: "patch", code: "1234", slot: 2, ops: [{ op: "release", key: "a|1|1" }] });
  assert.deepEqual(api.get({ jornada: "1234" }).live.owner, { a: ["1|1:1"] }, "un ajeno no la suelta");
  api.post({ action: "patch", code: "1234", slot: 1, ops: [{ op: "release", key: "a|1|1" }] });
  assert.deepEqual(api.get({ jornada: "1234" }).live.owner, {});
});

test("volver a coger la misma tarea no se rechaza a uno mismo", () => {
  const api = loadBackend();
  newJornada(api);
  api.post({ action: "patch", code: "1234", slot: 1, ops: [{ op: "claim", key: "a|1|1" }] });
  const again = api.post({ action: "patch", code: "1234", slot: 1, ops: [{ op: "claim", key: "a|1|1" }] });
  assert.deepEqual(again.rejected, []);
});

test("el endsAt del temporizador lo pone el reloj del servidor", () => {
  const api = loadBackend();
  newJornada(api);
  const res = api.post({ action: "patch", code: "1234", slot: 1, ops: [{ op: "timer", id: "1-7", key: "a|1|1", label: "Arroz", seconds: 600 }] });
  const timer = res.live.timers[0];
  assert.equal(timer.i, "1-7");
  assert.equal(timer.b, 1);
  assert.ok(Math.abs(timer.e - (res.now + 600000)) < 2000, "endsAt = ahora del servidor + segundos");
  api.post({ action: "patch", code: "1234", slot: 2, ops: [{ op: "timerStop", id: "1-7" }] });
  assert.equal(api.get({ jornada: "1234" }).live.timers.length, 0, "cualquiera puede pararlo");
});

test("el sondeo con el rev al día no manda el estado otra vez", () => {
  const api = loadBackend();
  newJornada(api);
  const full = api.get({ jornada: "1234" });
  const poll = api.get({ jornada: "1234", rev: String(full.rev) });
  assert.equal(poll.unchanged, true);
  assert.equal(poll.live, undefined);
  assert.equal(poll.plan, undefined);
  assert.ok(poll.now > 0);
  api.post({ action: "patch", code: "1234", slot: 0, ops: [{ op: "done", key: "a|1|1" }] });
  const next = api.get({ jornada: "1234", rev: String(full.rev) });
  assert.equal(next.unchanged, undefined);
  assert.equal(next.rev, full.rev + 1);
  assert.ok(next.live);
  assert.equal(next.plan, undefined, "el plan solo se manda cuando se pide");
  assert.deepEqual(api.get({ jornada: "1234", rev: String(next.rev), plan: "1" }).plan, PLAN);
});

test("recalcular el plan sube el planRev y conserva lo ya hecho", () => {
  const api = loadBackend();
  newJornada(api);
  api.post({ action: "patch", code: "1234", slot: 0, ops: [{ op: "done", key: "salmorreta|1|1" }] });
  const plan2 = Object.assign({}, PLAN, { end: 1200 });
  const res = newJornada(api, { replace: true, plan: plan2 });
  assert.equal(res.ok, true);
  assert.equal(res.planRev, 2);
  const after = api.get({ jornada: "1234" });
  assert.equal(after.plan.end, 1200);
  assert.deepEqual(after.live.done, { salmorreta: ["1|1:0"] }, "lo hecho no se pierde al recalcular");
});

test("un plan más largo que una propiedad se trocea y vuelve igual", () => {
  const api = loadBackend();
  const big = { v: 1, tasks: [] };
  for (let i = 0; i < 400; i++) big.tasks.push(["receta-larga-de-prueba|" + i + "|3", 480 + i, 500 + i, i % 3, null, "Vitro / fuego"]);
  assert.ok(JSON.stringify(big).length > 9000, "el plan de prueba debe pasar de 9 KB");
  newJornada(api, { plan: big });
  const chunks = api.store.get("j:1234:plan:n");
  assert.ok(Number(chunks) > 1, "se esperaban varios trozos");
  assert.deepEqual(api.get({ jornada: "1234" }).plan, big);
});

test("cerrar la jornada borra todos sus trozos", () => {
  const api = loadBackend();
  newJornada(api);
  api.post({ action: "close", code: "1234" });
  assert.equal(api.get({ jornada: "1234" }).missing, true);
  assert.deepEqual([...api.store.keys()].filter((k) => k.startsWith("j:1234")), []);
});

test("una acción desconocida no toca nada", () => {
  const api = loadBackend();
  newJornada(api);
  const res = api.post({ action: "borrar-todo" });
  assert.equal(res.ok, false);
  assert.equal(api.get({ jornada: "1234" }).ok, true);
});

test("una operación desconocida se rechaza y las buenas se aplican", () => {
  const api = loadBackend();
  newJornada(api);
  const res = api.post({ action: "patch", code: "1234", slot: 0, ops: [{ op: "done", key: "a|1|1" }, { op: "formatear" }] });
  assert.deepEqual(res.live.done, { a: ["1|1:0"] });
  assert.equal(res.rejected.length, 1);
});

test("el estado sondeado se mantiene pequeño con una jornada entera marcada", () => {
  const api = loadBackend();
  newJornada(api);
  const ops = [];
  for (let i = 0; i < 245; i++) ops.push({ op: "done", key: "croquetas-de-jamon|" + ((i % 9) + 1) + "|" + i, slot: i % 4 });
  const res = api.post({ action: "patch", code: "1234", slot: 0, ops });
  const size = JSON.stringify(res.live).length;
  assert.equal(Object.keys(res.live.done).reduce((n, k) => n + res.live.done[k].length, 0), 245);
  assert.ok(size < 9000, `el estado sondeado ocupa ${size} caracteres y el límite de una propiedad son 9 KB`);
});

test("las jornadas viejas se barren en la siguiente escritura", () => {
  const api = loadBackend();
  newJornada(api);
  // Una jornada de anteayer, con su plan y su estado.
  api.store.set("j:7777:meta", JSON.stringify({ createdAt: 1, at: 1, planRev: 1, people: 2 }));
  api.store.set("j:7777:live", JSON.stringify({ rev: 3, done: {}, owner: {}, timers: [], people: [] }));
  api.store.set("j:7777:plan:n", "1");
  api.store.set("j:7777:plan:0", "{}");
  const index = JSON.parse(api.store.get("j:index"));
  index["7777"] = Date.now() - 48 * 3600 * 1000;
  api.store.set("j:index", JSON.stringify(index));
  api.post({ action: "patch", code: "1234", slot: 0, ops: [{ op: "ping" }] });
  assert.deepEqual([...api.store.keys()].filter((k) => k.startsWith("j:7777")), [], "la vieja se borra");
  assert.equal(api.get({ jornada: "1234" }).ok, true, "la de hoy sigue");
});

/* ===================== El cliente (jornada.js) ===================== */

const planner = require("../planner.js");
const jornada = require("../jornada.js");
const recipes = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "plan-recipes.json"), "utf8")).recipes;

function renderable(plan) {
  return {
    start: plan.start,
    end: plan.end,
    days: plan.days,
    persons: plan.persons,
    busy: plan.busy,
    tasks: plan.tasks.map((t) => [
      jornada.taskKey(t), t.start, t.end, t.person, t.watcher === undefined ? null : t.watcher,
      t.equipment, t.type, t.personMin, t.totalMin, !!t.assumed, t.text, t.code, t.recipeName, t.block, t.step,
    ]),
    spans: plan.spans.map((s) => [s.key, s.label, s.code, s.recipeId, s.start, s.end]),
    recipes: plan.recipes.map((r) => [r.id, r.code, r.name, r.start, r.end, r.steps]),
  };
}

test("el plan compartido vuelve igual: nada de lo que se pinta se pierde", () => {
  const ids = recipes.map((r) => r.id);
  const plan = planner.buildPlan(recipes, ids, { people: 3 });
  const shared = jornada.packPlan(plan);
  const back = jornada.unpackPlan(shared, recipes);
  assert.deepEqual(renderable(back), renderable(plan));
  assert.deepEqual(back.missing, []);
  assert.equal(back.plannerVersion, planner.VERSION);
});

test("un paso que este dispositivo no tiene se pinta igual y queda avisado", () => {
  const plan = planner.buildPlan(recipes, recipes.map((r) => r.id), { people: 2 });
  const back = jornada.unpackPlan(jornada.packPlan(plan), []);
  assert.equal(back.missing.length, plan.tasks.length);
  assert.equal(back.tasks.length, plan.tasks.length);
  assert.equal(back.tasks[0].start, plan.tasks[0].start);
  assert.equal(back.tasks[0].equipment, plan.tasks[0].equipment);
  assert.equal(back.tasks[0].code, plan.tasks[0].code, "el código de la receta viaja en el plan");
  assert.match(back.tasks[0].text, /no está en los datos/);
});

test("el estado agrupado se deshace en claves de paso", () => {
  const live = { done: { croquetas: ["3|12:2", "4|1:0"] }, owner: { croquetas: ["3|12:2"] }, people: [{ s: 2, n: "Marta" }] };
  assert.deepEqual(jornada.doneMap(live), { "croquetas|3|12": 2, "croquetas|4|1": 0 });
  assert.deepEqual(jornada.ownerMap(live), { "croquetas|3|12": 2 });
  assert.deepEqual(jornada.peopleNames(live), { 2: "Marta" });
});

test("la cola guarda solo la última operación de cada tarea", () => {
  let q = [];
  q = jornada.queuePush(q, { op: "done", key: "a|1|1", slot: 1 });
  q = jornada.queuePush(q, { op: "undone", key: "a|1|1" });
  q = jornada.queuePush(q, { op: "done", key: "a|1|1", slot: 1 });
  q = jornada.queuePush(q, { op: "done", key: "a|1|2", slot: 1 });
  assert.equal(q.length, 2);
  assert.deepEqual(q[0], { op: "done", key: "a|1|1", slot: 1 });
});

test("lo que está en la cola se pinta ya, sin esperar al servidor", () => {
  const live = { done: { a: ["1|1:0"] }, owner: {} };
  const view = jornada.applyQueue(live, [{ op: "undone", key: "a|1|1" }, { op: "done", key: "a|1|2", slot: 3 }], 3);
  assert.deepEqual(view.done, { "a|1|2": 3 });
});

test("el desfase de reloj se mide descontando la mitad del viaje", () => {
  // El servidor va 10 min adelantado y la petición tardó 400 ms.
  const before = 1000000;
  const after = before + 400;
  const offset = jornada.offsetFrom(before + 600000 + 200, before, after);
  assert.ok(Math.abs(offset - 600000) <= 1, `desfase calculado: ${offset}`);
  assert.equal(jornada.remainingSeconds({ e: before + 600000 + 60000 }, offset, before), 60);
});

test("sin desfase guardado la cuenta atrás no se va a números negativos", () => {
  assert.equal(jornada.remainingSeconds({ e: 1000 }, 0, 99999), 0);
});

/* ---- Cliente y backend juntos: lo que haría una tablet ---- */

function wiredClient(api) {
  const calls = { get: 0, post: 0 };
  const fakeFetch = async (url, options) => {
    const text = (() => {
      if (options && options.method === "POST") {
        calls.post++;
        return JSON.stringify(api.post(JSON.parse(options.body)));
      }
      calls.get++;
      const parameter = {};
      String(url.split("?")[1] || "")
        .split("&")
        .filter(Boolean)
        .forEach((pair) => {
          const [k, v] = pair.split("=");
          parameter[decodeURIComponent(k)] = decodeURIComponent(v || "");
        });
      return JSON.stringify(api.get(parameter));
    })();
    return { ok: true, status: 200, text: async () => text };
  };
  return { client: jornada.createClient({ url: "https://script.example/exec", fetch: fakeFetch }), calls };
}

test("dos tablets sobre el mismo plan: una marca y la otra lo ve", async () => {
  const api = loadBackend();
  const a = wiredClient(api).client;
  const b = wiredClient(api).client;
  const plan = planner.buildPlan(recipes, recipes.map((r) => r.id), { people: 3 });
  const created = await a.create({ code: "4321", name: "Leo", inputs: { people: 3 }, plan: jornada.packPlan(plan), plannerVersion: planner.VERSION });
  assert.equal(created.code, "4321");
  assert.equal(typeof created.offset, "number", "cada respuesta trae el desfase de reloj");

  const joined = await b.join({ code: "4321", name: "Marta" });
  assert.notEqual(joined.slot, created.slot);
  const theirPlan = jornada.unpackPlan(joined.plan, recipes);
  assert.deepEqual(renderable(theirPlan), renderable(plan), "las dos tablets ven el mismo plan");

  const key = jornada.taskKey(plan.tasks[0]);
  await b.patch({ code: "4321", slot: joined.slot, ops: [{ op: "done", key }] });
  const polled = await a.read("4321", created.rev);
  assert.equal(polled.unchanged, undefined);
  assert.equal(jornada.doneMap(polled.live)[key], joined.slot);
  const again = await a.read("4321", polled.rev);
  assert.equal(again.unchanged, true, "sin cambios no se manda el estado otra vez");
});

test("un error del servidor llega como excepción con sus datos", async () => {
  const api = loadBackend();
  const { client } = wiredClient(api);
  await assert.rejects(() => client.join({ code: "0000", name: "Leo" }), (err) => err.data && err.data.missing === true);
});

test("el plan solo se vuelve a bajar cuando lo recalcula otro", async () => {
  const api = loadBackend();
  const a = wiredClient(api).client;
  const b = wiredClient(api);
  const plan = planner.buildPlan(recipes, recipes.map((r) => r.id), { people: 2 });
  await a.create({ code: "4321", name: "Leo", inputs: { people: 2 }, plan: jornada.packPlan(plan) });
  const joined = await b.client.join({ code: "4321", name: "Marta" });
  const plan2 = planner.buildPlan(recipes, recipes.map((r) => r.id), { people: 4 });
  await a.create({ code: "4321", replace: true, name: "Leo", inputs: { people: 4 }, plan: jornada.packPlan(plan2) });
  const poll = await b.client.read("4321", joined.rev);
  assert.ok(poll.planRev > joined.planRev, "el sondeo avisa de que el plan cambió");
  const fresh = await b.client.read("4321", poll.rev, true);
  assert.equal(jornada.unpackPlan(fresh.plan, recipes).persons, 4);
});

test("una tablet que lleva horas apagada deja su puesto libre", () => {
  const api = loadBackend();
  newJornada(api);
  const live = JSON.parse(api.store.get("j:1234:live"));
  live.people = [{ s: 0, n: "Leo", a: Date.now() - 8 * 3600 * 1000 }, { s: 1, n: "Marta", a: Date.now() }];
  api.store.set("j:1234:live", JSON.stringify(live));
  const res = api.post({ action: "join", code: "1234", name: "Nuevo" });
  assert.deepEqual(
    res.live.people.map((p) => p.n).sort(),
    ["Marta", "Nuevo"],
    "Leo, que no da señales desde hace 8 h, ya no ocupa sitio"
  );
  assert.equal(res.slot, 0, "el puesto que dejó libre se reutiliza");
});
