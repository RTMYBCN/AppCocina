/* Planificador de jornada · Rice to Meet You
 *
 * Lógica pura (sin DOM): recibe las recetas con su campo `plan` (que lee el
 * Apps Script del bloque M:Z de cada ficha) y devuelve un plan de trabajo:
 * qué paso hace cada persona, a qué hora y en qué equipo.
 *
 * Reglas (acordadas con Leo):
 *  - Solo entran los pasos de producción que traen datos de planificación.
 *  - Regla del stock: una ficha previa que no esté en la selección se da por
 *    hecha; si está en la selección, se prepara antes.
 *  - Cada ficha es una tanda. Un paso "POR TANDA" se repite según el
 *    rendimiento y las unidades por tanda.
 *  - Un equipo no puede atender dos pasos a la vez: se reserva desde el primer
 *    hasta el último paso que lo usa dentro de la misma etapa.
 *  - Las personas solo trabajan dentro de la jornada; las máquinas y las
 *    esperas siguen fuera de ella.
 *  - El día de cada etapa (campo Día) marca el primer momento en que puede
 *    empezar: el día siguiente arranca a la hora de inicio de la jornada.
 */
(function (root) {
  "use strict";

  const DAY = 1440;

  // Capacidad de cada equipo (unidades). Ajustar aquí si cambia el inventario.
  const EQUIPMENT = {
    "Thermomix grande": { key: "TMX_G", cap: 1 },
    "Thermomix pequeña": { key: "TMX_P", cap: 1 },
    "Thermomix (cualquiera)": { key: "TMX_ANY", cap: 1 },
    "Olla 25 L": { key: "OLLA25", cap: 2 },
    "Olla pequeña": { key: "OLLAP", cap: 3 },
    "Olla grande (sin presión)": { key: "OLLAG", cap: 2 },
    Vitro: { key: "VITRO", cap: 9 },
    "Fuego de gas": { key: "GAS", cap: 2 },
    "Vitro/fuego (cualquiera)": { key: "VITRO", cap: 9 },
    "Inducción IsEasy": { key: "IH", cap: 1 },
    Horno: { key: "HORNO", cap: 1 },
    Airfryer: { key: "AIR", cap: 1 },
    Freidora: { key: "FREID", cap: 1 },
    Panificadora: { key: "PAN", cap: 1 },
    Envasadora: { key: "ENV", cap: 1 },
    Microondas: { key: "MICRO", cap: 1 },
    Soplete: { key: "SOPLETE", cap: 1 },
  };
  const POOLS = { TMX_G: 1, TMX_P: 1, OLLA25: 2, OLLAP: 3, OLLAG: 2, VITRO: 9, GAS: 2, IH: 1, HORNO: 1, AIR: 1, FREID: 1, PAN: 1, ENV: 1, MICRO: 1, SOPLETE: 1 };
  const LABELS = {
    TMX_G: "Thermomix grande", TMX_P: "Thermomix pequeña", OLLA25: "Olla 25 L", OLLAP: "Olla pequeña",
    OLLAG: "Olla grande", VITRO: "Vitro / fuego", GAS: "Fuego de gas", IH: "Inducción IsEasy", HORNO: "Horno",
    AIR: "Airfryer", FREID: "Freidora", PAN: "Panificadora", ENV: "Envasadora", MICRO: "Microondas", SOPLETE: "Soplete",
  };

  // Equipos que la ficha deja a elección: el plan decide cuál se usa.
  const AMBIGUOUS = { TMX_ANY: true };

  const DEFAULTS = {
    people: 1,
    startMin: 8 * 60, // hora de inicio de la jornada (minutos desde las 00:00)
    shiftMin: 8 * 60, // duración de la jornada
    assumedPersonMin: 5, // paso con persona sin tiempo
    assumedMachineMin: 5, // paso de máquina sin tiempo
    assumedWaitMin: 30, // espera sin tiempo
    horizonDays: 14,
  };

  function num(v) {
    return typeof v === "number" && isFinite(v) ? v : null;
  }

  function hasPlan(recipe) {
    return !!(recipe && recipe.plan && Array.isArray(recipe.plan.steps) && recipe.plan.steps.length);
  }

  /* Interpola la espera según la cantidad (columnas "Otra cantidad"/"Min a esa
     cantidad"). Sin escala de cantidad en esta versión, devuelve el valor base. */
  function waitMinutes(step) {
    return num(step.mm);
  }

  function buildTasks(recipe, opts, assumptions) {
    const stepsOut = [];
    const plan = recipe.plan;
    const rows = plan.steps.slice().sort((a, b) => a.b - b.b || a.s - b.s);
    const yieldAmount = num(recipe.yieldAmount) || 1;
    rows.forEach((row) => {
      const stage = recipe.stages[row.b - 1];
      const text = stage && stage.steps[row.s - 1];
      if (!text) return;
      let mult = 1;
      if (row.rp && row.ub) {
        const n = Math.max(1, Math.ceil(yieldAmount / row.ub));
        mult = row.rp === "ENTRE TANDAS" ? Math.max(0, n - 1) : n;
      }
      let personMin = 0;
      let totalMin = 0;
      let assumed = false;
      switch (row.t) {
        case "Activo":
        case "Cocción atendida": {
          const pm = num(row.pm);
          personMin = (pm === null ? opts.assumedPersonMin : pm) * mult;
          totalMin = personMin;
          assumed = pm === null;
          break;
        }
        case "Máquina": {
          const mm = waitMinutes(row);
          totalMin = (mm === null ? opts.assumedMachineMin : mm) * mult;
          assumed = mm === null;
          break;
        }
        case "Espera": {
          const mm = waitMinutes(row);
          totalMin = (mm === null ? opts.assumedWaitMin : mm) * mult;
          assumed = mm === null;
          break;
        }
        default: {
          // Activo + espera: la persona hace algo y después queda una espera.
          const pm = num(row.pm);
          const mm = waitMinutes(row);
          personMin = (pm === null ? opts.assumedPersonMin : pm) * mult;
          totalMin = personMin + (mm === null ? opts.assumedWaitMin : mm) * mult;
          assumed = pm === null || mm === null;
        }
      }
      if (assumed) {
        assumptions.push({
          recipeId: recipe.id, code: recipe.code, name: recipe.name,
          block: row.b, step: row.s, type: row.t,
        });
      }
      stepsOut.push({
        recipeId: recipe.id,
        code: recipe.code,
        recipeName: recipe.name,
        block: row.b,
        step: row.s,
        text: String(text).trim(),
        day: num(row.d) || 1,
        equipment: row.eq,
        type: row.t,
        personMin,
        totalMin,
        scales: row.sc,
        pending: !!row.pend,
        note: row.n || null,
        assumed,
      });
    });
    return stepsOut;
  }

  const SEARCH_RUNS = 150;
  const EPS = 1e-9;

  function mulberry32(seed) {
    let a = seed >>> 0;
    return function () {
      a = (a + 0x6d2b79f5) >>> 0;
      let t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  // Una simulación completa con una prioridad y un máximo de personas.
  // Devuelve null si el trabajo no cabe en el horizonte.
  function simulate(ctx, prio, people, delays, anyStrict) {
    const { tasks, spans, firstSpan, opts, firstIds } = ctx;
    // Primer instante en que puede empezar un paso (día + retraso de arranque de la receta).
    const earliestOf = (task) =>
      (task.day - 1) * DAY + opts.startMin + (delays && firstIds.has(task.id) ? delays[task.recipeId] || 0 : 0);
    const n = tasks.length;
    const start = new Array(n).fill(null);
    const end = new Array(n).fill(null);
    const person = new Array(n).fill(null);
    const sp = spans.map(() => ({ unit: null, slot: 0, start: null, end: null, active: false }));
    const held = {};
    Object.keys(POOLS).forEach((k) => (held[k] = new Array(POOLS[k]).fill(false)));
    const freeSlot = (k) => held[k].indexOf(false);
    const busy = new Array(people).fill(0);
    const freeAt = new Array(people).fill(0);
    const lastPerson = new Map();
    const blocked = {};
    const pending = new Set(tasks.map((t) => t.id));
    const shiftEnd = opts.startMin + opts.shiftMin;
    const limit = opts.startMin + opts.horizonDays * DAY;

    let t = opts.startMin;
    while (pending.size && t < limit) {
      spans.forEach((s, i) => {
        const st = sp[i];
        if (st.active && end[s.last] !== null && end[s.last] <= t + EPS) {
          held[st.unit][st.slot] = false;
          st.active = false;
          st.end = end[s.last];
        }
      });
      const ready = [];
      pending.forEach((id) => {
        const task = tasks[id];
        if (task.preds.every((p) => end[p] !== null && end[p] <= t + EPS) && t >= earliestOf(task) - EPS) ready.push(task);
      });
      ready.sort((a, b) => prio[b.id] - prio[a.id] || a.day - b.day || a.id - b.id);
      for (const task of ready) {
        let who = null;
        if (task.personMin > 0) {
          const tm = t % DAY;
          const fits = tm >= opts.startMin && (tm + task.personMin <= shiftEnd || (task.personMin > opts.shiftMin && tm < shiftEnd));
          if (!fits) continue;
          const free = [];
          for (let p = 0; p < people; p++) if (freeAt[p] <= t + EPS) free.push(p);
          if (!free.length) {
            blocked["Personas"] = (blocked["Personas"] || 0) + 1;
            continue;
          }
          const prev = lastPerson.get(task.recipeId);
          who = free.includes(prev) ? prev : free.reduce((a, b) => (busy[b] < busy[a] ? b : a), free[0]);
        }
        const s = firstSpan.get(task.id);
        if (s) {
          const idx = ctx.spans.indexOf(s);
          let unit = s.key;
          if (s.key === "TMX_ANY") {
            unit = freeSlot("TMX_P") >= 0 ? "TMX_P" : anyStrict || freeSlot("TMX_G") < 0 ? null : "TMX_G";
          }
          const slot = unit === null ? -1 : freeSlot(unit);
          if (slot < 0) {
            const label = LABELS[s.key] || (s.key === "TMX_ANY" ? "Thermomix" : s.key);
            blocked[label] = (blocked[label] || 0) + 1;
            continue;
          }
          held[unit][slot] = true;
          sp[idx].unit = unit;
          sp[idx].slot = slot;
          sp[idx].active = true;
          sp[idx].start = t;
        }
        start[task.id] = t;
        end[task.id] = t + task.totalMin;
        person[task.id] = who;
        if (who !== null) {
          freeAt[who] = t + task.personMin;
          busy[who] += task.personMin;
          lastPerson.set(task.recipeId, who);
        }
        pending.delete(task.id);
      }
      if (!pending.size) break;
      // Siguiente instante en que algo puede cambiar.
      let next = Infinity;
      for (let i = 0; i < n; i++) if (end[i] !== null && end[i] > t + EPS && end[i] < next) next = end[i];
      for (let p = 0; p < people; p++) if (freeAt[p] > t + EPS && freeAt[p] < next) next = freeAt[p];
      pending.forEach((id) => {
        const e = earliestOf(tasks[id]);
        if (e > t + EPS && e < next) next = e;
      });
      const nextShift = (Math.floor(t / DAY) + (t % DAY >= opts.startMin ? 1 : 0)) * DAY + opts.startMin;
      if (nextShift > t + EPS && nextShift < next) next = nextShift;
      if (next === Infinity) return null;
      t = next;
    }
    if (pending.size) return null;
    let last = 0;
    for (let i = 0; i < n; i++) if (end[i] > last) last = end[i];
    return { start, end, person, spans: sp, busy, blocked, last };
  }

  function buildPlan(recipes, selectedIds, options) {
    const opts = Object.assign({}, DEFAULTS, options || {});
    opts.people = Math.max(1, Math.min(8, Math.round(opts.people)));
    opts.shiftMin = Math.max(60, Math.min(opts.shiftMin, DAY - opts.startMin));

    const byId = new Map(recipes.map((r) => [r.id, r]));
    const selected = [];
    const skipped = [];
    selectedIds.forEach((id) => {
      const r = byId.get(id);
      if (!r) return;
      (hasPlan(r) ? selected : skipped).push(r);
    });

    const result = {
      ok: selected.length > 0,
      options: opts,
      tasks: [],
      recipes: [],
      spans: [],
      persons: opts.people,
      warnings: [],
      assumptions: [],
      stock: [],
      skipped: skipped.map((r) => ({ id: r.id, code: r.code, name: r.name })),
      blocked: {},
      start: null,
      end: null,
    };
    if (!selected.length) return result;

    // ---- tareas
    const tasks = [];
    const firstOf = new Map();
    const lastOf = new Map();
    const selectedCodes = new Set(selected.map((r) => r.code));
    selected.forEach((r) => {
      const list = buildTasks(r, opts, result.assumptions);
      list.forEach((t) => {
        t.id = tasks.length;
        t.preds = [];
        t.succ = [];
        tasks.push(t);
      });
      if (list.length) {
        firstOf.set(r.id, list[0].id);
        lastOf.set(r.id, list[list.length - 1].id);
      }
      (r.plan.warnings || []).forEach((w) => result.warnings.push(`${r.code}: ${w}`));
    });

    // ---- precedencias dentro de cada receta ("Mientras…" va en paralelo al anterior)
    selected.forEach((r) => {
      const mine = tasks.filter((t) => t.recipeId === r.id);
      mine.forEach((t, i) => {
        if (i === 0) return;
        if (/^\s*mientras\b/i.test(t.text) && i >= 2) t.preds.push(mine[i - 2].id);
        else t.preds.push(mine[i - 1].id);
      });
    });
    // ---- precedencias entre recetas (regla del stock)
    const codeToRecipe = new Map(selected.map((r) => [r.code, r]));
    selected.forEach((r) => {
      (r.plan.requires || []).forEach((code) => {
        const prev = codeToRecipe.get(code);
        if (prev && prev.id !== r.id && firstOf.has(r.id) && lastOf.has(prev.id)) {
          tasks[firstOf.get(r.id)].preds.push(lastOf.get(prev.id));
        } else if (!selectedCodes.has(code)) {
          const name = (recipes.find((x) => x.code === code) || {}).name;
          result.stock.push({ code, name: name || code, neededBy: r.code });
        }
      });
    });
    tasks.forEach((t) => t.preds.forEach((p) => tasks[p].succ.push(t.id)));

    // ---- prioridad = camino crítico restante
    const cp = new Array(tasks.length).fill(null);
    const crit = (i) => {
      if (cp[i] !== null) return cp[i];
      cp[i] = tasks[i].totalMin + Math.max(0, ...tasks[i].succ.map(crit));
      return cp[i];
    };
    tasks.forEach((t) => crit(t.id));

    // ---- reservas de equipo: de su primer a su último paso en la misma etapa
    const spans = [];
    selected.forEach((r) => {
      const groups = new Map();
      tasks.filter((t) => t.recipeId === r.id).forEach((t) => {
        const eq = EQUIPMENT[t.equipment];
        if (!eq) return;
        const k = `${t.block}|${eq.key}`;
        if (!groups.has(k)) groups.set(k, { key: eq.key, ids: [] });
        groups.get(k).ids.push(t.id);
      });
      groups.forEach((g) => {
        spans.push({
          key: g.key, recipeId: r.id, code: r.code, ids: g.ids,
          first: g.ids[0], last: g.ids[g.ids.length - 1], unit: null, start: null, end: null, active: false,
        });
      });
    });
    const firstSpan = new Map(spans.map((s) => [s.first, s]));

    // ---- simulación por eventos. Un reparto voraz puede dar una máquina a la
    // receta equivocada (la que llega antes, no la que más tarda), así que se
    // prueban varias prioridades y se queda la que termina antes. Es determinista
    // (semillas fijas) y nunca empeora al añadir personas: siempre se prueba
    // también con menos personas.
    const ctx = { tasks, spans, firstSpan, opts, firstIds: new Set(firstOf.values()) };
    const DELAY_CHOICES = [0, 0, 2, 3, 5, 8, 12, 20, 30];
    const variants = [{ prio: cp, delays: null }];
    for (let k = 1; k <= SEARCH_RUNS; k++) {
      const rnd = mulberry32(k * 7919);
      const delays = {};
      selected.forEach((r) => {
        delays[r.id] = k % 3 === 0 ? 0 : DELAY_CHOICES[Math.floor(rnd() * DELAY_CHOICES.length)];
      });
      variants.push({ prio: cp.map((c) => c * (0.55 + 0.9 * rnd())), delays });
    }
    const sizes = [];
    for (let size = opts.people; size >= 1; size--) sizes.push(size);
    let best = null;
    const clock = Date.now();
    for (let i = 0; i < variants.length && (i < 4 || Date.now() - clock < 1200); i++) {
      for (const size of sizes) {
        for (const strict of [false, true]) {
          const run = simulate(ctx, variants[i].prio, size, variants[i].delays, strict);
          if (!run) continue;
          const spread = Math.max(...run.busy) - Math.min(...run.busy);
          const score = run.last * 1e4 + (size === opts.people ? 0 : 1) + spread / 1e3;
          if (!best || score < best.score) best = Object.assign(run, { score });
        }
      }
    }
    if (!best) {
      result.ok = false;
      result.warnings.push("No se pudo colocar todo el trabajo en el horizonte de planificación.");
      return result;
    }
    tasks.forEach((task) => {
      task.start = best.start[task.id];
      task.end = best.end[task.id];
      task.person = best.person[task.id];
    });
    spans.forEach((sp, i) => {
      sp.unit = best.spans[i].unit;
      sp.slot = best.spans[i].slot;
      sp.start = best.spans[i].start;
      sp.end = best.spans[i].end === null ? tasks[sp.last].end : best.spans[i].end;
    });
    // Nombre del aparato que de verdad le toca a cada paso. Si de un mismo
    // equipo hay varios y dos recetas lo usan a la vez, se numeran ("Olla 25 L
    // (1)" y "(2)"), para que nadie vaya a por el que ya está ocupado.
    const numbered = {};
    spans.forEach((a) => {
      if (!a.unit || POOLS[a.unit] < 2) return;
      spans.forEach((b) => {
        if (a === b || b.unit !== a.unit) return;
        if (a.start < b.end && b.start < a.end) numbered[a.unit] = true;
      });
    });
    spans.forEach((sp) => {
      if (!sp.unit) return;
      const base = LABELS[sp.unit];
      if (!base) return;
      sp.label = numbered[sp.unit] ? `${base} (${sp.slot + 1})` : base;
      if (AMBIGUOUS[sp.key] || numbered[sp.unit]) sp.ids.forEach((id) => (tasks[id].equipment = sp.label));
    });
    const busy = best.busy.concat(new Array(Math.max(0, opts.people - best.busy.length)).fill(0));
    const blocked = best.blocked;

    // ¿Se va trabajo al día siguiente solo porque no cabe en el turno? Si es así,
    // se calcula cuánto habría que alargarlo para terminarlo el mismo día.
    const pushed = tasks.filter((t) => Math.floor(best.start[t.id] / DAY) + 1 > t.day).length;
    if (pushed) {
      const wide = simulate(
        Object.assign({}, ctx, { opts: Object.assign({}, opts, { shiftMin: DAY - opts.startMin }) }),
        cp, opts.people, null, false
      );
      if (wide && wide.last < best.last) {
        result.overflow = {
          steps: pushed,
          endIfLonger: wide.last,
          shiftMinutes: Math.ceil((wide.last - opts.startMin) % DAY === 0 ? opts.shiftMin : wide.last - opts.startMin),
        };
      }
    }

    // Un paso de máquina o de espera no lo hace nadie, pero alguien lo deja en
    // marcha: se apunta a quien hizo el paso anterior de esa receta (o, si no
    // hay, a quien hace el siguiente) para que salga en su lista y sepa cuándo
    // volver. `watcher` no ocupa tiempo de esa persona.
    selected.forEach((r) => {
      const mine = tasks.filter((t) => t.recipeId === r.id && t.start !== null).sort((a, b) => a.start - b.start);
      mine.forEach((t, i) => {
        if (t.person !== null) return;
        let who = null;
        for (let j = i - 1; j >= 0 && who === null; j--) who = mine[j].person;
        for (let j = i + 1; j < mine.length && who === null; j++) who = mine[j].person;
        t.watcher = who === null ? 0 : who;
      });
    });

    // ---- resultado
    const done = tasks.filter((t) => t.start !== null);
    result.tasks = done.sort((a, b) => a.start - b.start || a.id - b.id);
    result.spans = spans
      .filter((s) => s.start !== null)
      .map((s) => ({ key: s.unit || s.key, label: s.label || LABELS[s.unit || s.key] || s.key, recipeId: s.recipeId, code: s.code, start: s.start, end: s.end }));
    result.blocked = blocked;
    result.busy = busy;
    result.start = Math.min(...done.map((t) => t.start));
    result.end = Math.max(...done.map((t) => t.end));
    result.days = Math.floor((result.end - 1) / DAY) + 1;
    result.recipes = selected.map((r) => {
      const mine = done.filter((t) => t.recipeId === r.id);
      return {
        id: r.id, code: r.code, name: r.name,
        start: mine.length ? Math.min(...mine.map((t) => t.start)) : null,
        end: mine.length ? Math.max(...mine.map((t) => t.end)) : null,
        steps: mine.length,
        requires: (r.plan.requires || []).filter((c) => selectedCodes.has(c)),
      };
    });
    result.personMinutes = busy.slice();
    return result;
  }

  // "08:05" / día 2 = "08:05 · día 2"
  function fmtClock(absMin) {
    const m = ((Math.round(absMin) % DAY) + DAY) % DAY;
    return String(Math.floor(m / 60)).padStart(2, "0") + ":" + String(m % 60).padStart(2, "0");
  }
  function dayOf(absMin) {
    return Math.floor(Math.round(absMin) / DAY) + 1;
  }
  function fmtDuration(min) {
    if (min < 1) return Math.max(1, Math.round(min * 60)) + " s";
    const m = Math.round(min);
    if (m < 60) return m + " min";
    const h = Math.floor(m / 60);
    const r = m % 60;
    if (h >= 24 && r === 0 && h % 24 === 0) return h / 24 + (h / 24 === 1 ? " día" : " días");
    return r ? `${h} h ${r} min` : `${h} h`;
  }

  const api = { buildPlan, hasPlan, fmtClock, fmtDuration, dayOf, DEFAULTS, EQUIPMENT, DAY };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.RTMYPlanner = api;
})(typeof window !== "undefined" ? window : globalThis);
