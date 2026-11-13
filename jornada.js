/* Jornada compartida · Rice to Meet You
 *
 * Varios dispositivos de la cocina sobre el MISMO plan: uno lo calcula y lo
 * sube, los demás se enganchan con un código de cuatro dígitos. Cada cocinero
 * ve lo suyo y lo que marca uno lo ven los demás.
 *
 * Aquí no se toca el DOM: solo el formato del plan compartido, la cola de
 * operaciones y las llamadas al Apps Script (ver JORNADA_COMPARTIDA.md).
 * La interfaz vive en app.js.
 */
(function (root) {
  "use strict";

  const SHARED_VERSION = 1;

  /* ---------------- El plan, empaquetado ----------------
     El texto de los pasos NO viaja: cada dispositivo lo saca de sus recetas
     por la clave `receta|bloque|paso`. Lo que viaja son las horas, la persona,
     el aparato y lo justo para pintar. */

  function taskKey(t) {
    return t.recipeId + "|" + t.block + "|" + t.step;
  }

  function packPlan(plan, extra) {
    return Object.assign(
      {
        v: SHARED_VERSION,
        plannerVersion: plan.version || (root.RTMYPlanner && root.RTMYPlanner.VERSION) || 0,
        start: plan.start,
        end: plan.end,
        days: plan.days,
        persons: plan.persons,
        busy: plan.busy || [],
        recipes: plan.recipes.map((r) => [r.id, r.code, r.name, r.start, r.end, r.steps]),
        spans: plan.spans.map((s) => [s.key, s.label, s.code, s.recipeId, s.start, s.end]),
        tasks: plan.tasks.map((t) => [
          taskKey(t), t.start, t.end,
          t.person === null || t.person === undefined ? -1 : t.person,
          t.watcher === null || t.watcher === undefined ? -1 : t.watcher,
          t.equipment || "", t.type || "", t.personMin, t.totalMin, t.assumed ? 1 : 0,
        ]),
      },
      extra || {}
    );
  }

  // `recipes` son las recetas que tiene este dispositivo: de ahí sale el texto.
  // Lo que no encuentre queda en `missing`: el plan se pinta igual (hora,
  // aparato) y la app avisa de que hay que actualizar los datos.
  function unpackPlan(shared, recipes) {
    const byId = new Map((recipes || []).map((r) => [r.id, r]));
    const missing = [];
    const tasks = (shared.tasks || []).map((row, i) => {
      const parts = String(row[0]).split("|");
      const recipeId = parts[0];
      const block = Number(parts[1]);
      const step = Number(parts[2]);
      const recipe = byId.get(recipeId);
      const stage = recipe && recipe.stages && recipe.stages[block - 1];
      const text = stage && stage.steps ? stage.steps[step - 1] : null;
      if (!text) missing.push(row[0]);
      const meta = (shared.recipes || []).find((r) => r[0] === recipeId);
      return {
        id: i,
        recipeId,
        code: (meta && meta[1]) || (recipe && recipe.code) || recipeId,
        recipeName: (meta && meta[2]) || (recipe && recipe.name) || recipeId,
        block,
        step,
        text: text ? String(text).trim() : "(este paso no está en los datos de este dispositivo)",
        start: row[1],
        end: row[2],
        person: row[3] < 0 ? null : row[3],
        watcher: row[4] < 0 ? null : row[4],
        equipment: row[5] || "",
        type: row[6] || "",
        personMin: row[7],
        totalMin: row[8],
        assumed: !!row[9],
      };
    });
    return {
      ok: tasks.length > 0,
      shared: true,
      plannerVersion: shared.plannerVersion || 0,
      start: shared.start,
      end: shared.end,
      days: shared.days,
      persons: shared.persons,
      busy: shared.busy || [],
      tasks,
      recipes: (shared.recipes || []).map((r) => ({ id: r[0], code: r[1], name: r[2], start: r[3], end: r[4], steps: r[5], requires: [] })),
      spans: (shared.spans || []).map((s) => ({ key: s[0], label: s[1], code: s[2], recipeId: s[3], start: s[4], end: s[5] })),
      warnings: [],
      assumptions: [],
      stock: [],
      skipped: [],
      blocked: {},
      missing,
    };
  }

  /* ---------------- Lo que se sondea ---------------- */

  // {"croquetas": ["3|12:2"]} -> {"croquetas|3|12": 2}
  function spread(map) {
    const out = {};
    Object.keys(map || {}).forEach((recipeId) => {
      (map[recipeId] || []).forEach((entry) => {
        const cut = String(entry).lastIndexOf(":");
        if (cut < 0) return;
        out[recipeId + "|" + String(entry).slice(0, cut)] = Number(String(entry).slice(cut + 1));
      });
    });
    return out;
  }

  function doneMap(live) {
    return spread(live && live.done);
  }

  function ownerMap(live) {
    return spread(live && live.owner);
  }

  function peopleNames(live) {
    const out = {};
    ((live && live.people) || []).forEach((p) => {
      if (p && typeof p.s === "number" && p.n) out[p.s] = p.n;
    });
    return out;
  }

  /* ---------------- La cola ----------------
     En una cocina el wifi se cae. Lo que marca el cocinero se pinta al
     instante, se guarda en la cola y se manda en cuanto haya red. Una tarea
     solo guarda su última operación: marcar y desmarcar cinco veces no manda
     cinco cosas. */

  function queueKeyOf(op) {
    if (op.op === "done" || op.op === "undone") return "done:" + op.key;
    if (op.op === "claim" || op.op === "release") return "owner:" + op.key;
    if (op.op === "timer" || op.op === "timerStop") return "timer:" + op.id;
    return null;
  }

  function queuePush(queue, op) {
    const key = queueKeyOf(op);
    const out = key ? queue.filter((x) => queueKeyOf(x) !== key) : queue.slice();
    out.push(op);
    return out;
  }

  // Lo que la cola cambia encima de lo que dice el servidor, para pintar sin
  // esperar la respuesta (interfaz optimista).
  function applyQueue(live, queue, slot) {
    const done = doneMap(live);
    const owner = ownerMap(live);
    (queue || []).forEach((op) => {
      if (op.op === "done") done[op.key] = op.slot === undefined ? slot : op.slot;
      else if (op.op === "undone") delete done[op.key];
      else if (op.op === "claim") {
        if (owner[op.key] === undefined) owner[op.key] = op.slot === undefined ? slot : op.slot;
      } else if (op.op === "release" && owner[op.key] === slot) delete owner[op.key];
    });
    return { done, owner };
  }

  /* ---------------- Relojes ----------------
     Los temporizadores compartidos guardan `endsAt` en hora del SERVIDOR. Si
     dos tablets van desfasadas, con el desfase guardado muestran lo mismo. */

  function offsetFrom(serverNow, before, after) {
    const latency = Math.max(0, (after - before) / 2);
    return Math.round(serverNow + latency - after);
  }

  function remainingSeconds(timer, offset, now) {
    return Math.max(0, Math.round((timer.e - (now + (offset || 0))) / 1000));
  }

  /* ---------------- Transporte ----------------
     Lecturas por GET y escrituras por POST con `text/plain`: ninguna de las
     dos provoca preflight, igual que la carga de recetas. */

  function createClient(config) {
    const cfg = config || {};
    const fetchFn = cfg.fetch || (typeof fetch === "function" ? fetch.bind(null) : null);
    const now = cfg.now || (() => Date.now());
    const timeout = cfg.timeoutMs || 12000;

    function url(params) {
      const q = Object.keys(params)
        .filter((k) => params[k] !== undefined && params[k] !== null)
        .map((k) => encodeURIComponent(k) + "=" + encodeURIComponent(params[k]))
        .join("&");
      return cfg.url + (cfg.url.indexOf("?") >= 0 ? "&" : "?") + q;
    }

    async function call(options) {
      if (!fetchFn) throw new Error("Sin fetch disponible.");
      const controller = typeof AbortController === "function" ? new AbortController() : null;
      const timer = controller ? setTimeout(() => controller.abort(), timeout) : null;
      const before = now();
      try {
        const res = await fetchFn(options.url, {
          method: options.method || "GET",
          // `text/plain` evita el preflight; Apps Script lee el cuerpo igual.
          headers: options.body ? { "Content-Type": "text/plain;charset=utf-8" } : undefined,
          body: options.body,
          redirect: "follow",
          signal: controller ? controller.signal : undefined,
        });
        if (!res.ok) throw new Error("HTTP " + res.status);
        const data = JSON.parse(await res.text());
        if (data && typeof data.now === "number") data.offset = offsetFrom(data.now, before, now());
        if (data && data.ok === false) {
          const err = new Error(data.error || "Error del servidor.");
          err.data = data;
          throw err;
        }
        return data;
      } finally {
        if (timer) clearTimeout(timer);
      }
    }

    function post(body) {
      return call({ url: cfg.url, method: "POST", body: JSON.stringify(body) });
    }

    return {
      create: (body) => post(Object.assign({ action: "create" }, body)),
      join: (body) => post(Object.assign({ action: "join" }, body)),
      close: (code) => post({ action: "close", code }),
      patch: (body) => post(Object.assign({ action: "patch" }, body)),
      // Sondeo: con el rev al día la respuesta son ~200 bytes.
      read: (code, rev, wantPlan) => call({ url: url({ jornada: code, rev: rev === undefined || rev === null ? undefined : rev, plan: wantPlan ? 1 : undefined }) }),
    };
  }

  const api = {
    SHARED_VERSION, taskKey, packPlan, unpackPlan,
    doneMap, ownerMap, peopleNames,
    queuePush, applyQueue, queueKeyOf,
    offsetFrom, remainingSeconds, createClient,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.RTMYJornada = api;
})(typeof window !== "undefined" ? window : globalThis);
