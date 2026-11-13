/* ============================================================
   RICE TO MEET YOU — Recetas del Taller
   App de una sola página: lista -> receta -> escalado -> pasos -> final.
   Fuente de datos: JSON servido por un Google Apps Script conectado
   en vivo al Excel "MANUAL COCINA" (ver CONFIG.DATA_URL).
   ============================================================ */

const CONFIG = {
  // Sustituir por la URL de despliegue del Apps Script (termina en /exec).
  // Mientras esté vacío, la app usa recipes.json local (modo desarrollo/demo).
  DATA_URL: "https://script.google.com/macros/s/AKfycbxldzxI4eqiW3i0nSrsj5JApxOKc2ujMVg6rixaRnkHqbzv_WTBkkrUsAAhq49nb01v/exec",
  LOCAL_FALLBACK: "recipes.json",
  // Última versión recibida del Apps Script (localStorage: sobrevive al cerrar
  // la app). Es el respaldo preferido si falla la carga: siempre más reciente
  // que recipes.json.
  LAST_GOOD_KEY: "rtmy_recipes_lastgood_v1",
  RUNTIME_KEY: "rtmy_runtime_v2",
  // Selección, personas y tareas hechas del planificador de jornada.
  PLAN_KEY: "rtmy_plan_v1",
  // Jornada compartida: el código, quién es este dispositivo y lo que queda
  // por mandar. Ver JORNADA_COMPARTIDA.md.
  JORNADA_KEY: "rtmy_jornada_v1",
  // Cada cuánto se pregunta al servidor por los cambios de los demás. Con
  // 4 tablets y 8 h de jornada son ~11.500 lecturas: muy por debajo de las
  // 50.000 diarias que admite PropertiesService.
  JORNADA_POLL_MS: 10000,
  // Si el Apps Script no responde a tiempo se usa el respaldo local en vez
  // de dejar la app en "Cargando recetas…" indefinidamente.
  FETCH_TIMEOUT_MS: 10000,
  // "Actualizar" pide al script que reconstruya desde el Sheet (sin su caché
  // de 5 min), lo que tarda más que una respuesta cacheada.
  FRESH_TIMEOUT_MS: 30000,
};

const MAX_SESSIONS = 4;

const state = {
  view: "loading", // loading | error | list | detail | steps | final | plan-setup | plan
  recipes: [],
  updatedAt: null, // última edición del Sheet (la informa el script)
  fetchedAt: null, // cuándo recibió la app estos datos
  refreshError: null, // aviso si la última actualización falló
  refreshingManual: false, // el botón "Actualizar" está esperando datos
  dataSource: null, // live | saved | fallback
  search: "",
  refreshing: false,
  timers: [],
  currentStepDurations: [],

  // Detail screen: a lightweight "draft" scale for whichever recipe is
  // currently being looked at, before (or while) it's being cooked.
  currentRecipe: null,
  scaleFactor: 1,
  scaleQty: null,
  scaleIngredientIndex: 0,
  sessionCapWarning: false,

  // Up to MAX_SESSIONS recipes can be "in preparation" at once — each one
  // keeps its own scale and step position so the cook can freely switch
  // between them (e.g. start a timer on one, work on another meanwhile).
  sessions: [],
  activeSessionId: null,

  // Snapshot used only by the final screen, captured at the moment a
  // session finishes (the session itself is removed right away so its
  // slot frees up for another recipe).
  finishedSession: null,

  // Planificador de jornada: qué recetas se preparan hoy, cuántas personas
  // y a qué hora empieza el turno. `result` es el plan calculado.
  // Jornada compartida: sin código, la app funciona como siempre, cada
  // dispositivo por su cuenta.
  jornada: {
    code: null, // código de cuatro dígitos
    slot: 0, // qué persona del plan es este dispositivo
    name: "", // su nombre (para reconocerlo al volver)
    rev: 0, // última versión del estado que se ha visto
    planRev: 0, // versión del plan compartido
    offset: 0, // desfase de este reloj contra el del servidor (ms)
    live: null, // último estado recibido
    queue: [], // operaciones pendientes de mandar (red caída)
    status: "off", // off | ok | offline | working
    error: null,
    joinCode: "", // lo que se está escribiendo en la caja del código
    inputs: null, // entradas con las que se calculó el plan compartido
    dataStamp: null,
    staleData: false, // el plan se calculó con otros datos de recetas
    panel: false, // el panel de compartir está abierto
    expired: false, // la jornada guardada ya no existía al volver a abrir
  },

  planUi: {
    selected: [],
    people: 2,
    start: "08:00",
    shiftH: 8,
    view: "persona", // persona | linea
    day: 1,
    done: {},
    result: null,
  },
};

const root = document.getElementById("app");
let lastFetchAt = 0;

function normalizeSearchText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function validateRecipesPayload(data) {
  if (!data || !Array.isArray(data.recipes)) {
    throw new Error("La respuesta no contiene una lista de recetas válida.");
  }
  const ids = new Set();
  const recipes = data.recipes.filter((recipe) => {
    if (!recipe || typeof recipe.id !== "string" || !recipe.id.trim()) return false;
    if (ids.has(recipe.id)) return false;
    if (typeof recipe.name !== "string" || !recipe.name.trim()) return false;
    if (!Array.isArray(recipe.stages) || !recipe.stages.length) return false;
    const validStages = recipe.stages.every((stage) => {
      if (!stage || typeof stage.title !== "string") return false;
      if (!Array.isArray(stage.ingredients) || !Array.isArray(stage.steps)) return false;
      const ingredientsOk = stage.ingredients.every(
        (ing) =>
          ing &&
          typeof ing.name === "string" &&
          Number.isFinite(ing.qty) &&
          ing.qty >= 0 &&
          typeof ing.unit === "string"
      );
      const stepsOk = stage.steps.every((step) => typeof step === "string" && step.trim());
      return ingredientsOk && stepsOk && (stage.steps.length || stage.description || stage.ingredients.length);
    });
    if (!validStages) return false;
    ids.add(recipe.id);
    return true;
  });
  if (!recipes.length) throw new Error("No se encontró ninguna receta válida.");
  return { ...data, recipes };
}

function saveRuntimeState() {
  try {
    localStorage.setItem(
      CONFIG.RUNTIME_KEY,
      JSON.stringify({
        sessions: state.sessions,
        timers: state.timers,
        activeSessionId: state.activeSessionId,
        finishedSession: state.finishedSession,
      })
    );
  } catch (e) {
    console.warn("No se pudo guardar el estado de cocina.", e);
  }
}

function restoreRuntimeState() {
  try {
    const saved = JSON.parse(localStorage.getItem(CONFIG.RUNTIME_KEY) || "null");
    if (!saved) return;
    state.sessions = Array.isArray(saved.sessions) ? saved.sessions : [];
    state.timers = Array.isArray(saved.timers) ? saved.timers : [];
    state.activeSessionId = saved.activeSessionId || null;
    state.finishedSession = saved.finishedSession || null;
    timerIdSeq = state.timers.reduce((max, timer) => Math.max(max, timer.id || 0), 0) + 1;
    state.timers.forEach(updateTimerRemaining);
  } catch (e) {
    console.warn("No se pudo restaurar el estado de cocina.", e);
  }
}

async function fetchRecipesFrom(url, source, opts) {
  opts = opts || {};
  // fresh: el script ignora su caché y relee el Sheet (un script antiguo
  // simplemente ignora el parámetro).
  const target = opts.fresh ? url + (url.includes("?") ? "&" : "?") + "fresh=" + Date.now() : url;
  const controller = new AbortController();
  const timer = setTimeout(
    () => controller.abort(),
    opts.fresh ? CONFIG.FRESH_TIMEOUT_MS : CONFIG.FETCH_TIMEOUT_MS
  );
  try {
    const res = await fetch(target, { cache: "no-store", signal: controller.signal });
    if (!res.ok) throw new Error("HTTP " + res.status);
    const data = validateRecipesPayload(await res.json());
    return { data, source };
  } finally {
    clearTimeout(timer);
  }
}

function readLastGood() {
  try {
    const saved = JSON.parse(localStorage.getItem(CONFIG.LAST_GOOD_KEY) || "null");
    if (!saved || !saved.data) return null;
    return { ts: saved.ts || null, data: validateRecipesPayload(saved.data) };
  } catch (e) {
    console.warn("La última versión guardada no era válida.", e);
    return null;
  }
}

function writeLastGood(data) {
  try {
    localStorage.setItem(CONFIG.LAST_GOOD_KEY, JSON.stringify({ ts: Date.now(), data }));
  } catch (e) {
    console.warn("No se pudo guardar la última versión de las recetas.", e);
  }
}

// Sustituye por completo las recetas en memoria por las recibidas.
function applyRecipes(data, source, fetchedAt) {
  state.recipes = sortRecipes(data.recipes || []);
  state.updatedAt = data.sourceUpdatedAt || data.fetchedAt || data.updatedAt || null;
  state.fetchedAt = fetchedAt || Date.now();
  state.dataSource = source;
  state.refreshError = null;
  relinkRecipeRefs();
}

/* Cada preparación abierta guarda SU PROPIA copia de la ficha, para que al
   cocinero no se le quede el paso en blanco si se cae la red. El problema:
   esa copia se guarda en localStorage y sobrevive a cerrar la app, así que
   una preparación empezada en octubre seguía mostrando la ficha de octubre
   para siempre, aunque el Sheet ya tuviera otra y se pulsara "Actualizar".
   Al llegar datos nuevos se vuelven a enganchar a la ficha de verdad. */
function relinkRecipeRefs() {
  let changed = false;
  state.sessions.forEach((session) => {
    const fresh = findRecipe(session.id);
    // Si la ficha ha desaparecido del Sheet se conserva la copia: mejor eso
    // que dejar al cocinero a medias sin pasos.
    if (!fresh || !fresh.stages || !fresh.stages.length) return;
    const same = JSON.stringify(fresh.stages) === JSON.stringify((session.recipe || {}).stages);
    session.recipe = fresh;
    if (same) return;
    changed = true;
    // La ficha ha cambiado: se reconstruye el recorrido y se intenta dejarle
    // en el mismo paso. Si ese paso ya no existe, en el más cercano que sí.
    const at = (session.flow || [])[session.flowPos] || null;
    const from = Math.min(session.startStageIndex || 0, fresh.stages.length - 1);
    session.startStageIndex = from;
    session.flow = buildFlow(fresh, from);
    let idx = -1;
    if (at) {
      idx = session.flow.findIndex(
        (x) => x.stageIndex === at.stageIndex && x.stepIndexInStage === at.stepIndexInStage
      );
      if (idx < 0) idx = session.flow.findIndex((x) => x.stageIndex === at.stageIndex);
    }
    if (idx < 0) idx = session.flowPos || 0;
    session.flowPos = Math.max(0, Math.min(idx, session.flow.length - 1));
    // Se avisa en pantalla: el cocinero tiene que saber que le han cambiado
    // la ficha debajo, no descubrirlo a medio paso.
    session.recipeChanged = true;
  });
  // La ficha abierta es OTRA referencia al objeto viejo: si no se reengancha,
  // el botón "Actualizar" trae los datos nuevos pero la pantalla sigue
  // pintando los de antes.
  if (state.currentRecipe) {
    const fresh = findRecipe(state.currentRecipe.id);
    if (fresh) state.currentRecipe = fresh;
  }
  if (state.finishedSession) {
    const fresh = findRecipe(state.finishedSession.recipe && state.finishedSession.recipe.id);
    if (fresh) state.finishedSession.recipe = fresh;
  }
  if (state.sessions.length) saveRuntimeState();
  return changed;
}

/* ---------------- Data loading ----------------
   The Apps Script always reads the Google Sheet live, so every fresh page
   load (or hard refresh) already shows the latest recipes automatically —
   no redeploy needed on the backend side.
   What it does NOT do on its own: refetch while the tab stays open, since
   the recipe list is only kept in memory for the session. To cover that
   (e.g. a tablet left open in the kitchen), we refetch when the tab
   regains focus, and expose a manual "Actualizar" button as an explicit
   fallback. Neither one interrupts an in-progress step-by-step flow —
   they only re-render when it's safe to do so (list view, or nothing
   currently open).
------------------------------------------------- */

function sortRecipes(recipes) {
  // Con secciones (B/S/C): por sección y número de ficha; sin ellas (datos
  // antiguos), alfabético como siempre.
  return recipes.slice().sort((a, b) => {
    const bySection = (a.sectionOrder || 0) - (b.sectionOrder || 0);
    if (bySection) return bySection;
    if (a.code && b.code && a.code !== b.code) return a.code < b.code ? -1 : 1;
    return a.name.localeCompare(b.name, "es", { sensitivity: "base" });
  });
}

let loadSeq = 0;

async function loadRecipes(opts) {
  opts = opts || {};
  const silent = !!opts.silent;
  const manual = !!opts.manual; // botón "Actualizar"

  // Un refresco automático (al volver a la pestaña) nunca pisa a uno manual
  // en curso: pide la versión con caché y podría devolver datos más viejos.
  if (!manual && silent && state.refreshingManual) return;

  // Solo la última petición lanzada puede cambiar el estado. Una pulsación de
  // "Actualizar" sustituye a un refresco automático que siga en vuelo, y la
  // respuesta tardía de este último se descarta.
  const seq = ++loadSeq;
  if (manual) state.refreshingManual = true;

  // 1) Primer pintado rápido con la última versión buena guardada (nunca con
  //    recipes.json mientras exista algo más reciente). La carga en vivo la
  //    sustituye en cuanto llega.
  if (!silent && !state.recipes.length) {
    const saved = readLastGood();
    if (saved) {
      applyRecipes(saved.data, "saved", saved.ts);
      state.view = "list";
      render();
    }
  }

  if (silent) {
    state.refreshing = true;
    state.refreshError = null;
    if (state.view === "list") render();
  }

  // Cierra esta carga (solo la última lanzada) y devuelve true si sigue vigente.
  const finish = () => {
    if (seq !== loadSeq) return false;
    state.refreshing = false;
    state.refreshingManual = false;
    return true;
  };

  try {
    let result;
    try {
      result = await fetchRecipesFrom(
        CONFIG.DATA_URL || CONFIG.LOCAL_FALLBACK,
        CONFIG.DATA_URL ? "live" : "fallback",
        { fresh: manual }
      );
    } catch (liveError) {
      if (!CONFIG.DATA_URL) throw liveError;
      if (seq !== loadSeq) return; // la sustituyó una petición más reciente
      if (state.recipes.length) {
        // Ya hay recetas en memoria (la última versión buena): se conservan.
        // Nunca se cambian por una copia más antigua.
        finish();
        state.refreshError = "No se pudo actualizar";
        if (state.view === "list") render();
        return;
      }
      result = await fetchRecipesFrom(CONFIG.LOCAL_FALLBACK, "fallback");
      state.refreshError = "No se pudo conectar con las recetas en vivo";
    }
    // Respuesta tardía de una petición ya sustituida (p. ej. el refresco
    // automático con caché, tras pulsar "Actualizar"): se descarta.
    if (!finish()) return;
    const { data, source } = result;
    const keepError = state.refreshError;
    applyRecipes(data, source);
    state.refreshError = keepError;
    if (source === "live") writeLastGood(data);
    state.view = state.view === "loading" ? "list" : state.view;
    lastFetchAt = Date.now();
    // Only force a re-render when it's safe: the list view, or first boot.
    // Detail/steps/final keep whatever was already on screen so an
    // in-progress recipe never changes under the cook's hands; the fresh
    // data is already in state.recipes for the next time they navigate.
    // Con "Actualizar" el cocinero ha pedido ver los datos nuevos: se repinta
    // también la ficha o el paso que tenga abierto (antes solo el listado, y
    // parecía que el botón no hacía nada).
    if (state.view === "list" || manual || root.dataset.rendered !== "1") render();
  } catch (err) {
    if (!finish()) return;
    if (!state.recipes.length) {
      state.view = "error";
      state.errorMsg = err.message;
      render();
    } else if (state.view === "list") {
      render();
    }
  }
}

// Botón "Actualizar": pide los datos más recientes del Sheet y reemplazan a
// los que hay en memoria.
function refreshRecipes() {
  loadRecipes({ silent: true, manual: true });
}

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState !== "visible") return;
  if (state.view === "loading") return;
  // Avoid refetching too often if the tab is switched back and forth.
  if (Date.now() - lastFetchAt < 30 * 1000) return;
  loadRecipes({ silent: true });
});

/* ---------------- Helpers ---------------- */

function findRecipe(id) {
  return state.recipes.find((r) => r.id === id) || null;
}

const fmtQty = RTMYText.fmtQty;

// "SALMORRETA RTMY" -> "Salmorreta RTMY". Las siglas van en una lista cerrada:
// una regla por longitud confundiría "DE" o "PAN" con siglas.
const ACRONYMS = new Set(["RTMY", "AOVE", "MCC", "APPCC", "PVP", "IVA", "IH"]);
function titleCase(s) {
  if (!s) return "";
  const out = s.replace(/[^\s]+/g, (w) => (ACRONYMS.has(w) ? w : w.toLowerCase()));
  return out.charAt(0).toUpperCase() + out.slice(1);
}

// Detect ALL-CAPS warning clauses and parenthetical notes inside a step.
const CAPS_RE = /\b(?:[A-ZÁÉÍÓÚÑÜ0-9]{2,}[.,:;]?(?:\s+|$)){2,}/g;
const UNIT_RE = /\b\d+(?:[.,]\d+)?\s?(?:g|gr|gramos?|ml|kg|cl|l|min|minutos?|seg|segundos?|horas?|h\b|grados?|°|nivel\s?\d+)\b/gi;

function escapeHtml(s) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function renderStepText(raw) {
  let text = escapeHtml(raw);
  text = text.replace(UNIT_RE, (m) => `<b>${m}</b>`);
  return text;
}

function extractWarnings(raw) {
  const warnings = [];
  let match;
  CAPS_RE.lastIndex = 0;
  while ((match = CAPS_RE.exec(raw))) {
    const phrase = match[0].trim();
    const words = phrase.split(/\s+/).filter((w) => /[A-ZÁÉÍÓÚÑÜ]/.test(w));
    if (words.length >= 2 && phrase.length > 4) {
      warnings.push(phrase.replace(/[.,:;]+$/, ""));
    }
  }
  return [...new Set(warnings)];
}

/* ---------------- Timers ----------------
   Steps often name a duration ("cocinar durante 8 minutos"). We detect
   those mentions and offer an optional timer button per step — starting
   one is always the cook's choice, never automatic. Timers live in a
   persistent bar outside #app (see index.html) so they keep counting
   down and stay visible no matter which screen/step the cook navigates
   to next; only #app gets replaced on render(), the timer bar never does.
------------------------------------------------- */

const DURATION_RE = /(\d+(?:[.,]\d+)?)(?:\s*[-–]\s*(\d+(?:[.,]\d+)?))?\s*(horas?|minutos?|segundos?)\b/gi;

function extractDurations(text) {
  const results = [];
  const seen = new Set();
  let m;
  DURATION_RE.lastIndex = 0;
  while ((m = DURATION_RE.exec(text))) {
    const n1 = parseFloat(m[1].replace(",", "."));
    const n2 = m[2] ? parseFloat(m[2].replace(",", ".")) : null;
    const unit = m[3].toLowerCase();
    // Ranges ("2-3 minutos") use the upper bound, so the alarm never
    // fires before the dish could realistically be ready.
    const value = n2 !== null ? n2 : n1;
    let seconds;
    if (unit.startsWith("hora")) seconds = value * 3600;
    else if (unit.startsWith("minuto")) seconds = value * 60;
    else seconds = value;
    seconds = Math.round(seconds);
    if (seconds <= 0) continue;
    const label = m[0].trim();
    const key = label.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    results.push({ label, seconds });
  }
  return results;
}

let timerIdSeq = 1;
let timerTickHandle = null;
let audioCtx = null;
const timerBarEl = document.getElementById("timer-bar");

function ensureAudioCtx() {
  try {
    if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    else if (audioCtx.state === "suspended") audioCtx.resume().catch(() => {});
  } catch (e) {}
  return audioCtx;
}

function playAlarmSound() {
  const ctx = audioCtx;
  if (!ctx) return;
  try {
    const now = ctx.currentTime;
    for (let i = 0; i < 3; i++) {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = "sine";
      osc.frequency.value = 880;
      const t0 = now + i * 0.35;
      gain.gain.setValueAtTime(0.0001, t0);
      gain.gain.exponentialRampToValueAtTime(0.35, t0 + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.0001, t0 + 0.25);
      osc.connect(gain).connect(ctx.destination);
      osc.start(t0);
      osc.stop(t0 + 0.3);
    }
  } catch (e) {}
}

function fmtClock(totalSeconds) {
  const m = Math.floor(totalSeconds / 60);
  const s = totalSeconds % 60;
  return `${m}:${s.toString().padStart(2, "0")}`;
}

function ensureTimerTick() {
  if (timerTickHandle) return;
  timerTickHandle = setInterval(() => {
    let changed = false;
    state.timers.forEach((t) => {
      if (t.running && !t.done) {
        updateTimerRemaining(t);
        changed = true;
        if (t.remaining <= 0) {
          t.remaining = 0;
          t.running = false;
          t.done = true;
          playAlarmSound();
          if (navigator.vibrate) navigator.vibrate([200, 100, 200, 100, 400]);
        }
      }
    });
    if (!state.timers.some((t) => t.running)) {
      clearInterval(timerTickHandle);
      timerTickHandle = null;
    }
    if (changed) {
      saveRuntimeState();
      renderTimerBar();
      renderPlanTimers();
    }
  }, 1000);
}

function updateTimerRemaining(timer) {
  if (!timer.running || timer.done) return;
  timer.remaining = Math.max(0, Math.ceil((timer.endsAt - Date.now()) / 1000));
}

function startTimer(label, seconds, sessionId, recipeName, person) {
  const dup = state.timers.find((t) => t.label === label && t.sessionId === (sessionId || null) && !t.done);
  if (dup) return; // already running/paused, avoid a duplicate of the same duration
  ensureAudioCtx(); // create/unlock audio now, while we still have a user gesture
  state.timers.push({
    id: timerIdSeq++,
    label,
    remaining: seconds,
    endsAt: Date.now() + seconds * 1000,
    running: true,
    done: false,
    sessionId: sessionId || null,
    recipeName: recipeName || null,
    // Quién lo puso en marcha (índice de persona del plan), para que aparezca
    // en su franja y no se mezcle con los de los demás.
    person: Number.isInteger(person) ? person : null,
  });
  saveRuntimeState();
  ensureTimerTick();
  renderTimerBar();
  renderPlanTimers();
  return state.timers[state.timers.length - 1];
}

window.startTimerByIndex = function (i) {
  const d = (state.currentStepDurations || [])[i];
  if (!d) return;
  const session = activeSession();
  startTimer(d.label, d.seconds, session ? session.id : null, session ? session.recipe.name : null);
};

window.toggleTimer = function (id) {
  const t = state.timers.find((x) => x.id === id);
  if (!t || t.done) return;
  t.running = !t.running;
  if (t.running) {
    t.endsAt = Date.now() + t.remaining * 1000;
    ensureTimerTick();
  }
  saveRuntimeState();
  renderTimerBar();
  renderPlanTimers();
};

window.cancelTimer = function (id) {
  const gone = state.timers.find((x) => x.id === id);
  if (gone && gone.shared && planShared()) {
    jornadaQueue({ op: "timerStop", id: gone.shared });
    jornadaFlush();
  }
  state.timers = state.timers.filter((x) => x.id !== id);
  saveRuntimeState();
  renderTimerBar();
  renderPlanTimers();
};

// Resuming a recipe from a finished timer's "Ir a la receta" button should
// dismiss that pill, not leave it sitting there for the cook to close by hand.
window.resumeFromTimer = function (sessionId, timerId) {
  cancelTimer(timerId);
  switchToSession(sessionId);
};

// Número de franjas del plan que hay ahora mismo en pantalla. Fuera del plan
// (o con una sola persona) la barra de abajo no se divide.
function planBarPersons() {
  if (state.view !== "plan") return 0;
  const plan = state.planUi.result;
  return plan && plan.ok && plan.persons > 1 ? plan.persons : 0;
}

function timerRowHtml(t, compact) {
  const sessionStillOpen = t.sessionId && state.sessions.some((s) => s.id === t.sessionId);
  // En la franja de su dueño sobra repetir "P1": lo dice la cabecera.
  const who = !compact && Number.isInteger(t.person) ? `<b class="rtmy-timer-who">P${t.person + 1}</b> ` : "";
  let info;
  if (t.done && compact) {
    // En franja estrecha basta con el aviso y el paso: de quién es lo dice la cabecera.
    info = `<span class="rtmy-timer-clock">⏰ ¡Listo!</span><span class="rtmy-timer-label">${escapeHtml(t.label)}</span>`;
  } else if (t.done) {
    info = t.recipeName
      ? `<span class="rtmy-timer-clock">⏰</span><span class="rtmy-timer-label">${who}<b>${escapeHtml(titleCase(t.recipeName))}</b> necesita tu atención — ${escapeHtml(t.label)}</span>`
      : `<span class="rtmy-timer-clock">¡Listo!</span><span class="rtmy-timer-label">${escapeHtml(t.label)}</span>`;
  } else {
    info = `<span class="rtmy-timer-clock">${fmtClock(t.remaining)}</span><span class="rtmy-timer-label">${who}${escapeHtml(t.label)}${t.recipeName && !compact ? " · " + escapeHtml(titleCase(t.recipeName)) : ""}</span>`;
  }
  return `
    <div class="rtmy-timer-row ${t.done ? "done" : ""}">
      <div class="rtmy-timer-info">${info}</div>
      <div class="rtmy-timer-controls">
        ${t.done && sessionStillOpen ? `<button class="rtmy-timer-goto" onclick="resumeFromTimer('${t.sessionId}', ${t.id})">Ir a la receta</button>` : ""}
        ${
          !t.done
            ? `<button onclick="toggleTimer(${t.id})" aria-label="${t.running ? "Pausar" : "Reanudar"}">${t.running ? "⏸" : "▶️"}</button>`
            : ""
        }
        <button onclick="cancelTimer(${t.id})" aria-label="Descartar">✕</button>
      </div>
    </div>`;
}

function renderTimerBar() {
  if (!timerBarEl) return;
  document.body.classList.toggle("has-timer", state.timers.length > 0);
  if (!state.timers.length) {
    timerBarEl.innerHTML = "";
    timerBarEl.classList.remove("visible", "cols");
    return;
  }
  timerBarEl.classList.add("visible");

  // Dentro del plan la barra se parte en las mismas franjas que la pantalla:
  // cada cuenta atrás cae debajo de quien la puso, no en el centro.
  const persons = planBarPersons();
  if (!persons) {
    timerBarEl.classList.remove("cols");
    timerBarEl.innerHTML = state.timers.map((t) => timerRowHtml(t, false)).join("");
    return;
  }
  timerBarEl.classList.add("cols");
  timerBarEl.style.setProperty("--cols", persons);
  let lanes = "";
  for (let p = 0; p < persons; p++) {
    const mine = state.timers.filter((t) => t.person === p);
    lanes += `
      <div class="rtmy-timer-lane${mine.some((t) => t.done) ? " alert" : ""}">
        <span class="rtmy-timer-lane-head">Persona ${p + 1}</span>
        ${mine.length ? mine.map((t) => timerRowHtml(t, true)).join("") : `<span class="rtmy-timer-lane-empty">Sin temporizadores</span>`}
      </div>`;
  }
  // Los de fuera del plan (una receta abierta a mano) no son de nadie.
  const loose = state.timers.filter((t) => !Number.isInteger(t.person) || t.person >= persons);
  timerBarEl.innerHTML = `<div class="rtmy-timer-grid">${lanes}</div>${loose.map((t) => timerRowHtml(t, false)).join("")}`;
}

/* Las cantidades dentro del texto de los pasos (actualizar las que ya
   vienen escritas en la ficha al escalar, y añadir las que faltan) viven
   en text.js, que se puede probar con node porque no toca el DOM. */
// Los pasos de una etapa, o su descripción si no los tiene desglosados.
function stageStepsOf(stage) {
  return stage.steps && stage.steps.length ? stage.steps : [stage.description];
}

/* ---------------- Router ----------------
   Every screen (list / detail / steps / final) gets its own hash so the
   header "Volver" button and the device's native back gesture always
   trigger a real hashchange — otherwise setting the hash to the value it
   already has is a no-op and the screen never updates (this was the bug:
   steps and final used to just flip state.view without touching the URL,
   so going "back" from them did nothing).

   Cooking a recipe now lives in a "session" (see Sessions below), so the
   steps screen is reached via #/sesion/<id> — a session must already
   exist (created by startFlow, or resumed by switchToSession) before that
   route resolves; the hash itself never encodes stage/step position.
------------------------------------------------- */

const ROUTE_SESION_RE = /^#\/sesion\/([^/]+)$/;
const ROUTE_RECETA_RE = /^#\/receta\/([^/]+)(?:\/(final))?$/;

function onHashChange() {
  const hash = location.hash || "#/";
  const sesionMatch = hash.match(ROUTE_SESION_RE);
  const recetaMatch = !sesionMatch ? hash.match(ROUTE_RECETA_RE) : null;

  if (hash === "#/" || hash === "") {
    state.view = "list";
  } else if (hash === "#/planificar") {
    state.view = "plan-setup";
  } else if (hash === "#/planificar/plan") {
    // Sin recetas elegidas no hay plan: se vuelve a la selección.
    // En una jornada compartida el plan no se recalcula: es el del servidor.
    if (!planShared()) state.planUi.result = null;
    // Una tablet que solo se engancha a una jornada no tiene selección propia:
    // su plan es el del servidor.
    state.view = planSelectedValid().length || state.jornada.code ? "plan" : "plan-setup";
  } else if (sesionMatch) {
    const id = decodeURIComponent(sesionMatch[1]);
    const s = findSession(id);
    if (s) {
      state.activeSessionId = id;
      state.view = "steps";
    } else {
      // No live session for this id (e.g. after a page reload — sessions
      // only live in memory). Fall back to the recipe's detail screen.
      const r = findRecipe(id);
      if (r) {
        state.currentRecipe = r;
        state.view = "detail";
        resetScale(r);
      } else {
        state.view = "list";
      }
    }
  } else if (recetaMatch) {
    const id = decodeURIComponent(recetaMatch[1]);
    const sub = recetaMatch[2];
    const r = findRecipe(id);
    if (!r) {
      state.view = "list";
    } else if (sub === "final") {
      state.view =
        state.finishedSession && state.finishedSession.recipe.id === id ? "final" : "detail";
      if (state.view === "detail") {
        state.currentRecipe = r;
        resetScale(r);
      }
    } else {
      state.currentRecipe = r;
      state.view = "detail";
      resetScale(r);
    }
  } else {
    state.view = "list";
  }
  render();
  renderSessionBar();
  window.scrollTo(0, 0);
}

function goto(hash) {
  location.hash = hash;
}

/* ---------------- Sessions ----------------
   A "session" is one recipe currently being cooked: its own scale and its
   own position in the step flow, independent from every other session.
   Up to MAX_SESSIONS can be open at once so the cook can start a timer on
   one dish and switch to another while it runs, then switch back.
------------------------------------------------- */

function findSession(recipeId) {
  return state.sessions.find((s) => s.id === recipeId) || null;
}

function activeSession() {
  return findSession(state.activeSessionId);
}

function switchToSession(id) {
  if (!findSession(id)) return;
  goto("#/sesion/" + encodeURIComponent(id));
}
window.switchToSession = switchToSession;

window.closeSession = function (id) {
  state.sessions = state.sessions.filter((s) => s.id !== id);
  if (state.activeSessionId === id) {
    state.activeSessionId = null;
    if (state.view === "steps") goto("#/");
  }
  if (state.sessions.length < MAX_SESSIONS) state.sessionCapWarning = false;
  saveRuntimeState();
  renderSessionBar();
  if (state.view === "detail") renderDetail();
};

/* ---------------- Scaling ---------------- */

function resetScale(recipe) {
  // If this recipe already has an open session, reflect its current
  // scale here instead of resetting to the base quantity.
  const existing = findSession(recipe.id);
  if (existing) {
    state.scaleFactor = existing.scaleFactor;
    state.scaleQty = existing.scaleQty;
    state.scaleIngredientIndex = existing.scaleIngredientIndex || 0;
    return;
  }
  state.scaleIngredientIndex = 0;
  const stage0 = recipe.stages[0];
  const main = stage0 && stage0.ingredients[0];
  state.scaleQty = main ? main.qty : null;
  state.scaleFactor = 1;
}

// The "main" ingredient is whichever one the cook has chosen to scale the
// recipe by (defaults to the first ingredient in the table).
function mainIngredient(recipe) {
  const stage0 = recipe.stages[0];
  if (!stage0 || !stage0.ingredients.length) return null;
  const idx = state.scaleIngredientIndex || 0;
  return stage0.ingredients[idx] || stage0.ingredients[0];
}

window.onSelectMainIngredient = function (idx) {
  const r = state.currentRecipe;
  const stage0 = r.stages[0];
  if (!stage0 || !stage0.ingredients[idx]) return;
  const newMain = stage0.ingredients[idx];
  state.scaleIngredientIndex = idx;
  // Keep the current scale ratio, just re-anchor the editable quantity
  // to whichever ingredient is now the scaling reference.
  state.scaleQty = newMain.qty * state.scaleFactor;
  renderDetail();
};

function applyScale(ingredients, ratio) {
  // Las unidades contables se redondean hacia arriba: 2,8 huevos son 3.
  return ingredients.map((ing) => ({ ...ing, qty: RTMYText.scaleQty(ing.qty, ing.unit, ratio) }));
}

function scaledIngredients(recipe) {
  const stage0 = recipe.stages[0];
  if (!stage0) return [];
  return applyScale(stage0.ingredients, state.scaleFactor);
}

function sliderRange(mainQty) {
  const min = Math.max(1, Math.round(mainQty * 0.25));
  const max = Math.round(mainQty * 3);
  const step = mainQty >= 200 ? 10 : mainQty >= 20 ? 5 : 1;
  return { min, max, step };
}

/* ---------------- Flow (stage + step navigation) ---------------- */

function buildFlow(recipe, startStageIndex) {
  const flow = [];
  for (let si = startStageIndex; si < recipe.stages.length; si++) {
    const stage = recipe.stages[si];
    const steps = stage.steps && stage.steps.length ? stage.steps : (stage.description ? [stage.description] : []);
    steps.forEach((stepText, idx) => {
      flow.push({
        stageIndex: si,
        stepIndexInStage: idx,
        totalInStage: steps.length,
        stageTitle: stage.title,
      });
    });
  }
  return flow;
}

function startFlow(recipe, startStageIndex) {
  let session = findSession(recipe.id);
  if (!session && state.sessions.length >= MAX_SESSIONS) {
    // Already cooking the maximum of MAX_SESSIONS different recipes at
    // once — show the limit inline instead of silently doing nothing.
    state.sessionCapWarning = true;
    renderDetail();
    return;
  }
  if (!session) {
    session = { id: recipe.id, recipe, scaleFactor: 1, scaleQty: null, scaleIngredientIndex: 0, startStageIndex: 0, flow: [], flowPos: 0 };
    state.sessions.push(session);
  }
  session.recipe = recipe;
  session.scaleFactor = state.scaleFactor;
  session.scaleQty = state.scaleQty;
  session.scaleIngredientIndex = state.scaleIngredientIndex || 0;
  session.startStageIndex = startStageIndex;
  session.flow = buildFlow(recipe, startStageIndex);
  if (!session.flow.length) {
    state.sessions = state.sessions.filter((item) => item !== session);
    state.errorMsg = "Esta etapa no contiene pasos ejecutables.";
    state.view = "error";
    render();
    return;
  }
  session.flowPos = 0;
  state.activeSessionId = recipe.id;
  state.sessionCapWarning = false;
  saveRuntimeState();
  goto("#/sesion/" + encodeURIComponent(recipe.id));
}

/* ---------------- Views ---------------- */

function render() {
  root.dataset.rendered = "1";
  // El plan se ve mejor ancho en tablet/ordenador: las franjas caben a la vez.
  document.body.classList.toggle("view-plan", state.view === "plan");
  if (state.view === "loading") renderLoading();
  else if (state.view === "error") renderError();
  else if (state.view === "list") renderList();
  else if (state.view === "detail") renderDetail();
  else if (state.view === "steps") renderSteps();
  else if (state.view === "final") renderFinal();
  else if (state.view === "plan-setup") renderPlanSetup();
  else if (state.view === "plan") renderPlan();
  renderTimerBar(); // al entrar o salir del plan la barra cambia de forma
}

function header(showBack, backHash, title) {
  return `
    <div class="rtmy-header">
      <div class="rtmy-badge"><span>RTMY</span></div>
      <div class="rtmy-title">
        <h1>Rice to Meet You</h1>
        <p>${title || "Recetas del taller"}</p>
      </div>
    </div>
    ${
      showBack
        ? `<button class="rtmy-back" onclick="goto('${backHash}')">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 18l-6-6 6-6"/></svg>
            Volver
          </button>`
        : ""
    }
  `;
}

function renderLoading() {
  root.innerHTML = `
    ${header(false)}
    <div class="rtmy-empty">Cargando recetas…</div>
  `;
}

function renderError() {
  root.innerHTML = `
    ${header(false)}
    <div class="rtmy-panel">
      <p><strong>No se pudieron cargar las recetas.</strong></p>
      <p style="font-size:13px;opacity:.75">${escapeHtml(state.errorMsg || "")}</p>
      <button class="rtmy-cta" onclick="loadRecipes()">Reintentar</button>
    </div>
  `;
}

// "hoy 18:55" / "6 nov 18:55": con fecha, para saber de qué día es cada dato.
function fmtWhen(ts) {
  if (!ts) return "";
  try {
    const d = new Date(ts);
    if (isNaN(d)) return "";
    const time = d.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" });
    if (d.toDateString() === new Date().toDateString()) return "hoy " + time;
    const day = d.toLocaleDateString("es-ES", { day: "numeric", month: "short" }).replace(".", "");
    return day + " " + time;
  } catch (e) {
    return "";
  }
}

function recipesStatusText() {
  if (state.refreshing) return "Actualizando…";
  if (state.dataSource === "fallback") return "Copia local · puede estar desactualizada";
  const edited = state.updatedAt ? " · hoja editada " + fmtWhen(state.updatedAt) : "";
  if (state.dataSource === "saved") return "Última versión guardada de " + fmtWhen(state.fetchedAt) + edited;
  return "Actualizado " + fmtWhen(state.fetchedAt) + edited;
}

function recipesErrorText() {
  if (!state.refreshError || state.refreshing) return "";
  if (state.dataSource === "fallback") return state.refreshError + ".";
  return state.refreshError + ": se mantiene la versión de " + fmtWhen(state.fetchedAt) + ".";
}

function buildRecipeCardHtml(r) {
  const inProgress = !!findSession(r.id);
  return `
        <button class="rtmy-card ${inProgress ? "in-progress" : ""}" onclick="goto('#/receta/${encodeURIComponent(r.id)}')">
          <div>
            <div class="rtmy-card-name">${escapeHtml(titleCase(r.name))}</div>
            <div class="rtmy-card-meta">${r.yieldAmount ?? ""} ${escapeHtml(
    r.yieldUnit || ""
  )} · ${r.stages.length} etapa${r.stages.length === 1 ? "" : "s"}</div>
            ${inProgress ? '<div class="rtmy-card-status">👩‍🍳 En preparación</div>' : ""}
          </div>
          <svg class="rtmy-card-arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 6l6 6-6 6"/></svg>
        </button>
      `;
}

function buildListItemsHtml() {
  const q = normalizeSearchText(state.search.trim());
  const list = state.recipes.filter((r) => normalizeSearchText(r.name).includes(q));
  if (!list.length) return `<div class="rtmy-empty">No se encontraron recetas.</div>`;

  // Recetas con sección (preparaciones base, sub-preparaciones, platos…):
  // se agrupan bajo su título. Sin sección, lista plana como antes.
  if (!list.some((r) => r.section)) return list.map(buildRecipeCardHtml).join("");

  const groups = [];
  list.forEach((r) => {
    const title = r.section || "Otras recetas";
    let g = groups[groups.length - 1];
    if (!g || g.title !== title) {
      g = { title, items: [] };
      groups.push(g);
    }
    g.items.push(r);
  });
  return groups
    .map(
      (g) => `
        <section class="rtmy-section" aria-label="${escapeHtml(g.title)}">
          <h2 class="rtmy-section-title">${escapeHtml(g.title)}<span class="rtmy-section-count">${g.items.length}</span></h2>
          <div class="rtmy-section-items">${g.items.map(buildRecipeCardHtml).join("")}</div>
        </section>`
    )
    .join("");
}

function renderList() {
  const errorText = recipesErrorText();
  root.innerHTML = `
    ${header(false)}
    <div class="rtmy-refresh-row">
      <span>${escapeHtml(recipesStatusText())}</span>
      <button class="rtmy-refresh-btn" onclick="refreshRecipes()" ${state.refreshingManual ? "disabled" : ""}>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="${state.refreshing ? "spin" : ""}"><path d="M4 4v5h5M20 20v-5h-5"/><path d="M4 9a8 8 0 0 1 14-4.5M20 15a8 8 0 0 1-14 4.5"/></svg>
        Actualizar
      </button>
    </div>
    ${errorText ? `<div class="rtmy-refresh-error" role="status">${escapeHtml(errorText)}</div>` : ""}
    ${planEntryHtml()}
    <label class="rtmy-sr-only" for="recipe-search">Buscar receta</label>
    <input id="recipe-search" class="rtmy-search" placeholder="Buscar receta…" value="${escapeHtml(
      state.search
    )}" oninput="onSearch(this.value)">
    <div class="rtmy-list" id="rtmy-list-container">
      ${buildListItemsHtml()}
    </div>
    <p class="rtmy-footer-note">Rice to Meet You · El Raval, Barcelona</p>
  `;
}

window.onSearch = function (val) {
  state.search = val;
  // Only refresh the results container — never touch the <input> itself,
  // otherwise re-rendering it destroys and recreates the DOM node, which
  // drops focus (and on mobile, closes the keyboard) after every keystroke.
  const container = document.getElementById("rtmy-list-container");
  if (container) container.innerHTML = buildListItemsHtml();
};

/* Cuántos ingredientes hay que alistar para empezar por esa etapa. Solo se
   dice cuando la ficha reparte la tabla por fases; si no, no aporta nada. */
function stageIngredientNote(recipe, stageIndex) {
  if (!RTMYText.recipeUsesPhases(recipe)) return "";
  const n = RTMYText.ingredientsForStage(recipe, stageIndex).length;
  if (!n) return " · sin ingredientes propios";
  return ` · ${n} ingrediente${n === 1 ? "" : "s"}`;
}

/* La tabla de ingredientes se agrupa por fases (Masa, Empanado, Fritura…)
   para que se vea de un golpe qué es de cada etapa. El índice que se le pasa
   a onSelectMainIngredient tiene que seguir siendo el de la tabla sin agrupar:
   es el que escala la receta. */
function renderIngredientPicker(recipe) {
  const groups = RTMYText.groupByPhase(scaledIngredients(recipe));
  const showHeads = groups.length > 1;
  return groups
    .map((group) => {
      const head =
        showHeads && group.phase
          ? `<li class="rtmy-ingredient-group">${escapeHtml(titleCase(group.phase))}</li>`
          : "";
      const rows = group.items
        .map(({ ingredient, index }) => {
          const selected = index === (state.scaleIngredientIndex || 0);
          return `
          <li>
            <button class="rtmy-ingredient-select ${selected ? "selected" : ""}" onclick="onSelectMainIngredient(${index})" aria-pressed="${selected}">
              <span class="rtmy-ingredient-radio" aria-hidden="true"></span>
              <span class="name">${escapeHtml(titleCase(ingredient.name))}</span>
              <span class="qty">${fmtQty(ingredient.qty)} ${escapeHtml(ingredient.unit)}</span>
            </button>
          </li>`;
        })
        .join("");
      return head + rows;
    })
    .join("");
}

function renderDetail() {
  const r = state.currentRecipe;
  const stage0 = r.stages[0];
  const main = mainIngredient(r);
  const hasMultipleStages = r.stages.length > 1;
  const session = findSession(r.id);

  root.innerHTML = `
    ${header(true, "#/", "Detalle de receta")}
    ${
      state.sessionCapWarning
        ? `
    <div class="rtmy-panel rtmy-cap-warning">
      <p><strong>Ya tienes ${MAX_SESSIONS} recetas en preparación a la vez.</strong></p>
      <p style="font-size:13px;opacity:.8;margin-top:4px;">Cierra una para poder empezar esta.</p>
      <div class="rtmy-session-list">
        ${state.sessions
          .map((s) => {
            const item = s.flow[s.flowPos];
            const step = item ? `Paso ${item.stepIndexInStage + 1} de ${item.totalInStage}` : "";
            return `<div class="rtmy-session-list-row">
              <span>${escapeHtml(titleCase(s.recipe.name))} <span style="opacity:.6;">· ${step}</span></span>
              <button class="rtmy-cta secondary" style="padding:6px 12px;margin:0;" onclick="closeSession('${s.id}')">Cerrar</button>
            </div>`;
          })
          .join("")}
      </div>
    </div>`
        : ""
    }
    <div class="rtmy-panel">
      <div class="rtmy-recipe-yield">${r.yieldAmount ?? ""} ${escapeHtml(r.yieldUnit || "")}</div>
      <h2 class="rtmy-recipe-name">${escapeHtml(titleCase(r.name))}</h2>
      ${session ? `<span class="rtmy-inprogress-badge">👩‍🍳 En preparación</span>` : ""}
      ${stage0 && stage0.description ? `<p class="rtmy-recipe-desc">${escapeHtml(stage0.description)}</p>` : ""}

      ${
        main
          ? `
      <div class="rtmy-scale-block">
        <div class="rtmy-scale-label">Escalar receta</div>
        <div class="rtmy-scale-row">
          <span class="rtmy-scale-ingredient">${escapeHtml(titleCase(main.name))}</span>
          <input class="rtmy-scale-input" id="scaleInput" type="number" min="0" step="any"
            value="${fmtQty(state.scaleQty)}" onchange="onScaleInput(this.value)">
          <span class="rtmy-scale-unit">${escapeHtml(main.unit)}</span>
        </div>
        ${renderSlider(main)}
        <span class="rtmy-ratio-pill">${state.scaleFactor === 1 ? "Cantidad original" : "×" + (Math.round(state.scaleFactor * 100) / 100).toString().replace(".", ",")}</span>
      </div>

      <div class="rtmy-scale-label" style="margin-top:14px;">Toca un ingrediente para escalar por él</div>
      <ul class="rtmy-ingredients rtmy-ingredients-selectable">
        ${renderIngredientPicker(r)}
      </ul>
      `
          : `<p style="opacity:.7;font-size:14px;">Esta preparación no tiene una tabla de ingredientes propia (usa componentes de otra receta).</p>`
      }
    </div>

    ${
      session
        ? `<button class="rtmy-cta" onclick="switchToSession('${r.id}')">Continuar donde lo dejaste</button>`
        : ""
    }

    ${
      hasMultipleStages
        ? `
    <div class="rtmy-panel">
      <div class="rtmy-scale-label">${session ? "O empezar de nuevo desde…" : "¿Por dónde empezamos?"}</div>
      <div class="rtmy-stage-picker">
        ${r.stages
          .map(
            (s, i) => `
          <button class="rtmy-stage-btn" onclick="startFlow(state.currentRecipe, ${i})">
            <span class="rtmy-stage-num">${i + 1}</span>
            <span>
              <div class="stage-title">${escapeHtml(titleCase(s.title))}</div>
              <div class="stage-sub">${(s.steps && s.steps.length) || 0} paso${(s.steps || []).length === 1 ? "" : "s"}${stageIngredientNote(r, i)}</div>
            </span>
          </button>
        `
          )
          .join("")}
      </div>
    </div>
    `
        : !session
          ? `<button class="rtmy-cta" onclick="startFlow(state.currentRecipe, 0)">Empezar receta</button>`
          : ""
    }

    <p class="rtmy-footer-note">Rice to Meet You · El Raval, Barcelona</p>
  `;
}

function renderSlider(main) {
  const { min, max, step } = sliderRange(main.qty);
  const pct = Math.max(
    0,
    Math.min(100, ((state.scaleQty - min) / (max - min)) * 100)
  );
  return `
    <input type="range" class="rtmy-slider" id="scaleSlider"
      aria-label="Cantidad de ${escapeHtml(main.name)}"
      min="${min}" max="${max}" step="${step}" value="${state.scaleQty}"
      style="--pct:${pct}%"
      oninput="onScaleSlider(this.value)">
    <div class="rtmy-slider-minmax"><span>${fmtQty(min)} ${escapeHtml(main.unit)}</span><span>${fmtQty(max)} ${escapeHtml(main.unit)}</span></div>
  `;
}

window.onScaleSlider = function (val) {
  const r = state.currentRecipe;
  const main = mainIngredient(r);
  state.scaleQty = parseFloat(val);
  state.scaleFactor = state.scaleQty / main.qty;
  renderDetail();
};

window.onScaleInput = function (val) {
  const r = state.currentRecipe;
  const main = mainIngredient(r);
  const v = parseFloat(val.replace(",", "."));
  if (isNaN(v) || v <= 0) return;
  state.scaleQty = v;
  state.scaleFactor = v / main.qty;
  renderDetail();
};

window.dismissRecipeChanged = function () {
  const session = activeSession();
  if (!session) return;
  session.recipeChanged = false;
  saveRuntimeState();
  renderSteps();
};

function renderSteps() {
  const session = activeSession();
  if (!session) {
    goto("#/");
    return;
  }
  const r = session.recipe;
  const flow = session.flow;
  const pos = session.flowPos;
  const item = flow[pos];
  if (!item) {
    state.sessions = state.sessions.filter((itemSession) => itemSession.id !== session.id);
    saveRuntimeState();
    goto("#/receta/" + encodeURIComponent(session.recipe.id));
    return;
  }
  const stage = r.stages[item.stageIndex];
  const rawStepText = stageStepsOf(stage)[item.stepIndexInStage];
  const warnings = extractWarnings(rawStepText);
  const durations = extractDurations(rawStepText);
  state.currentStepDurations = durations;

  // Las cantidades se miden contra la tabla de la etapa 0 (la que el
  // cocinero escala) o, en recetas compuestas, contra la tabla propia de esa
  // etapa, que no se escala. Se pasa la tabla SIN escalar más la proporción:
  // así se reescriben también las cantidades que la ficha ya trae escritas.
  // Se procesa la etapa entera, no solo este paso, porque la cantidad de
  // cada ingrediente se dice una sola vez: la primera que aparece.
  // Con fases declaradas, cada etapa saca sus ingredientes de la tabla
  // escalada: así los pasos del boleado también dicen "los 6 ud de huevos",
  // que antes salían sin cantidad porque esa etapa no tenía tabla propia.
  const usesPhases = RTMYText.recipeUsesPhases(r);
  const isMainStage = item.stageIndex === 0;
  const stageIngredients = RTMYText.ingredientsForStage(r, item.stageIndex);
  const stageRatio = usesPhases || isMainStage ? session.scaleFactor : 1;
  const stepText = RTMYText.stepTexts(stageStepsOf(stage), stageIngredients, stageRatio)[
    item.stepIndexInStage
  ];

  // progress dots: one per stage in active range
  const activeStages = r.stages.slice(session.startStageIndex);
  const dots = activeStages
    .map((s, idx) => {
      const stageIdx = session.startStageIndex + idx;
      let cls = "dot";
      if (stageIdx < item.stageIndex) cls += " done";
      else if (stageIdx === item.stageIndex) cls += " current";
      return `<span class="${cls}"></span>`;
    })
    .join("");

  const isLast = pos === flow.length - 1;
  const isFirst = pos === 0;

  root.innerHTML = `
    ${header(true, "#/receta/" + encodeURIComponent(r.id), escapeHtml(titleCase(r.name)))}
    <div class="rtmy-progress">${dots}</div>
    <div class="rtmy-stage-heading">${escapeHtml(titleCase(stage.title))}</div>
    <div class="rtmy-step-count">Paso ${item.stepIndexInStage + 1} de ${item.totalInStage}</div>
    ${
      session.recipeChanged
        ? `<div class="rtmy-warning" style="display:flex;width:100%;box-sizing:border-box;justify-content:space-between;">
      <span>La ficha se ha actualizado. Revisa el paso antes de seguir.</span>
      <button class="chip-close" onclick="dismissRecipeChanged()" aria-label="Entendido">✕</button>
    </div>`
        : ""
    }
    ${
      item.stepIndexInStage === 0 && (usesPhases || item.stageIndex !== 0) && stageIngredients.length
        ? `
    <div class="rtmy-panel" style="padding:14px 16px;">
      <div class="rtmy-scale-label" style="margin-bottom:2px;">Ingredientes de esta etapa</div>
      <ul class="rtmy-ingredients">
        ${applyScale(stageIngredients, stageRatio)
          .map(
            (i) => `
          <li><span class="name">${escapeHtml(titleCase(i.name))}</span><span class="qty">${fmtQty(i.qty)} ${escapeHtml(i.unit)}</span></li>
        `
          )
          .join("")}
      </ul>
    </div>`
        : ""
    }
    <div class="rtmy-step-card">
      <div class="rtmy-step-number">${item.stepIndexInStage + 1}</div>
      <div>
        <div class="rtmy-step-text">${renderStepText(stepText)}</div>
        ${warnings
          .map((w) => {
            const danger = /^NO\b|^SIN\b/.test(w);
            return `<div class="rtmy-warning ${danger ? "danger" : ""}">${danger ? "🚫" : "⚠️"} ${escapeHtml(w)}</div>`;
          })
          .join("")}
        ${
          durations.length
            ? `<div class="rtmy-timer-chips">
                ${durations
                  .map(
                    (d, i) =>
                      `<button class="rtmy-timer-chip" onclick="startTimerByIndex(${i})">⏱️ Iniciar ${escapeHtml(d.label)}</button>`
                  )
                  .join("")}
              </div>`
            : ""
        }
      </div>
    </div>
    <div class="rtmy-step-nav">
      ${
        !isFirst
          ? `<button class="rtmy-cta secondary" style="flex:1" onclick="stepBack()">Atrás</button>`
          : ""
      }
      <button class="rtmy-cta" style="flex:2" onclick="${
        isLast ? "finishFlow()" : "stepNext()"
      }">${isLast ? "Finalizar" : "Siguiente"}</button>
    </div>
  `;
}

window.stepNext = function () {
  const s = activeSession();
  if (!s) return;
  if (s.flowPos < s.flow.length - 1) {
    s.flowPos++;
    saveRuntimeState();
    render();
    renderSessionBar();
    window.scrollTo(0, 0);
  }
};
window.stepBack = function () {
  const s = activeSession();
  if (!s) return;
  if (s.flowPos > 0) {
    s.flowPos--;
    saveRuntimeState();
    render();
    renderSessionBar();
    window.scrollTo(0, 0);
  }
};
window.finishFlow = function () {
  const s = activeSession();
  if (!s) return;
  state.finishedSession = { recipe: s.recipe, scaleFactor: s.scaleFactor || 1 };
  state.sessions = state.sessions.filter((x) => x.id !== s.id);
  if (state.activeSessionId === s.id) state.activeSessionId = null;
  saveRuntimeState();
  goto("#/receta/" + encodeURIComponent(s.recipe.id) + "/final");
};

function renderFinal() {
  const r = state.finishedSession.recipe;
  const lastStage = r.stages[r.stages.length - 1];
  // Las últimas indicaciones también llevan cantidades: se escalan igual
  // que en la pantalla de pasos (solo si esa etapa es la que se escala).
  const ratio = r.stages.length === 1 ? state.finishedSession.scaleFactor || 1 : 1;
  const steps = lastStage.steps && lastStage.steps.length
    ? RTMYText.stepTexts(lastStage.steps, lastStage.ingredients, ratio)
    : [];

  root.innerHTML = `
    ${header(true, "#/receta/" + encodeURIComponent(r.id), escapeHtml(titleCase(r.name)))}
    <div class="rtmy-final-icon">✅</div>
    <h2 class="rtmy-final-title">¡Receta terminada!</h2>
    <div class="rtmy-panel">
      <div class="rtmy-scale-label">Racionado, envasado y conservación</div>
      <p class="rtmy-recipe-desc" style="margin-top:6px;">Últimas indicaciones de "${escapeHtml(
        titleCase(lastStage.title)
      )}":</p>
      <ul class="rtmy-ingredients" style="border-top:none;">
        ${steps
          .map(
            (s) => `<li style="display:block;"><span style="display:block;">${renderStepText(s)}</span></li>`
          )
          .join("")}
      </ul>
    </div>
    <button class="rtmy-cta" onclick="goto('#/')">Volver al listado de recetas</button>
    <p class="rtmy-footer-note">Rice to Meet You · El Raval, Barcelona</p>
  `;
}

/* ---------------- Session switcher bar ----------------
   Persistent, outside #app (like the timer bar) — shows every recipe
   currently being cooked so the cook can jump between them freely.
------------------------------------------------- */

const sessionBarEl = document.getElementById("session-bar");

function renderSessionBar() {
  if (!sessionBarEl) return;
  document.body.classList.toggle("has-session", state.sessions.length > 0);
  if (!state.sessions.length) {
    sessionBarEl.innerHTML = "";
    sessionBarEl.classList.remove("visible");
    return;
  }
  sessionBarEl.classList.add("visible");
  const sessionChips = state.sessions
    .map((s) => {
      const active = state.view === "steps" && s.id === state.activeSessionId;
      const item = s.flow[s.flowPos];
      const stepInfo = item ? `Paso ${item.stepIndexInStage + 1}/${item.totalInStage}` : "";
      return `
      <div class="rtmy-session-chip ${active ? "active" : ""}">
        <button class="chip-open" onclick="switchToSession('${s.id}')" ${active ? 'aria-current="step"' : ""}>
          <span class="chip-name">${escapeHtml(titleCase(s.recipe.name))}</span>
          <span class="chip-step">${stepInfo}</span>
        </button>
        <button class="chip-close" onclick="closeSession('${s.id}')" aria-label="Cerrar ${escapeHtml(titleCase(s.recipe.name))}">✕</button>
      </div>`;
    })
    .join("");
  sessionBarEl.innerHTML = `
    <div class="rtmy-session-bar-content">
      <button class="rtmy-new-prep-btn" onclick="goto('#/')" aria-label="Iniciar una nueva preparación">
        <span aria-hidden="true">＋</span>
        Nueva preparación
      </button>
      <div class="rtmy-session-chips">${sessionChips}</div>
    </div>
  `;
}

/* ---------------- Planificador de jornada ----------------
   Dos pantallas: elegir qué se prepara (y cuántas personas hay) y ver el
   plan resultante, por persona o en línea de tiempo. La lógica de reparto
   vive en planner.js; aquí solo se pinta y se guardan las preferencias.
------------------------------------------------- */

const PLAN_COLORS = ["#7a2426", "#5b3a2c", "#5a8a52", "#b5593a", "#c98a1b", "#7d8452", "#8a5a7a", "#3f6f7a"];
const planSoloCache = new WeakMap();

function planRecipes() {
  return window.RTMYPlanner ? state.recipes.filter(RTMYPlanner.hasPlan) : [];
}

function savePlanState() {
  try {
    const ui = state.planUi;
    localStorage.setItem(
      CONFIG.PLAN_KEY,
      JSON.stringify({ selected: ui.selected, people: ui.people, start: ui.start, shiftH: ui.shiftH, view: ui.view, done: ui.done })
    );
  } catch (e) {
    console.warn("No se pudo guardar el planificador.", e);
  }
}

function restorePlanState() {
  try {
    const saved = JSON.parse(localStorage.getItem(CONFIG.PLAN_KEY) || "null");
    if (!saved) return;
    const ui = state.planUi;
    if (Array.isArray(saved.selected)) ui.selected = saved.selected.filter((id) => typeof id === "string");
    if (Number.isFinite(saved.people)) ui.people = Math.max(1, Math.min(8, Math.round(saved.people)));
    if (/^\d{2}:\d{2}$/.test(saved.start || "")) ui.start = saved.start;
    if (Number.isFinite(saved.shiftH)) ui.shiftH = Math.max(2, Math.min(14, Math.round(saved.shiftH)));
    if (saved.view === "linea" || saved.view === "persona") ui.view = saved.view;
    if (saved.done && typeof saved.done === "object") ui.done = saved.done;
  } catch (e) {
    console.warn("No se pudo restaurar el planificador.", e);
  }
}

function planOptions() {
  const [h, m] = state.planUi.start.split(":").map(Number);
  return { people: state.planUi.people, startMin: h * 60 + m, shiftMin: state.planUi.shiftH * 60 };
}

// La selección guardada nunca se recorta aquí: si por un momento las recetas no
// traen datos de planificación (p. ej. durante una recarga), se conserva y el
// plan se calcula solo con las que sí los traen.
function computePlan() {
  const ui = state.planUi;
  ui.result = RTMYPlanner.buildPlan(state.recipes, ui.selected, planOptions());
  return ui.result;
}

function planSelectedValid() {
  const valid = new Set(planRecipes().map((r) => r.id));
  return state.planUi.selected.filter((id) => valid.has(id));
}

// Duración de una receta trabajando una sola persona (para orientar al elegir).
function planSoloText(recipe) {
  let text = planSoloCache.get(recipe);
  if (text === undefined) {
    const p = RTMYPlanner.buildPlan([recipe], [recipe.id], { people: 1 });
    text = p.ok ? (p.days > 1 ? p.days + " días" : RTMYPlanner.fmtDuration(p.end - p.start)) : "";
    planSoloCache.set(recipe, text);
  }
  return text;
}

function planEntryHtml() {
  if (!planRecipes().length) return "";
  const n = planSelectedValid().length;
  return `
    <button class="rtmy-plan-entry" onclick="goto('#/planificar')">
      <span class="rtmy-plan-entry-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>
      </span>
      <span class="rtmy-plan-entry-text">
        <b>Planificar jornada</b>
        <small>${n ? n + " receta" + (n === 1 ? "" : "s") + " elegida" + (n === 1 ? "" : "s") : "Reparte el trabajo entre personas y equipos"}</small>
      </span>
      <svg class="rtmy-card-arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 6l6 6-6 6"/></svg>
    </button>`;
}

function planGroupBySection(list) {
  const groups = [];
  list.forEach((r) => {
    const title = r.section || "Otras recetas";
    let g = groups[groups.length - 1];
    if (!g || g.title !== title) {
      g = { title, items: [] };
      groups.push(g);
    }
    g.items.push(r);
  });
  return groups;
}

/* ----- Pantalla 1: elegir recetas y jornada ----- */

function renderPlanSetup() {
  const ui = state.planUi;
  const withPlan = planRecipes();
  const missing = state.recipes.filter((r) => !withPlan.includes(r));
  const selected = new Set(planSelectedValid());
  const n = selected.size;

  const stockNames = [];
  withPlan.forEach((r) => {
    if (!selected.has(r.id)) return;
    (r.plan.requires || []).forEach((code) => {
      const dep = state.recipes.find((x) => x.code === code);
      if (dep && !selected.has(dep.id) && !stockNames.includes(titleCase(dep.name))) stockNames.push(titleCase(dep.name));
    });
  });

  const pick = (r) => {
    const on = selected.has(r.id);
    const solo = planSoloText(r);
    return `
      <button type="button" class="rtmy-plan-pick ${on ? "selected" : ""}" role="checkbox" aria-checked="${on}" onclick="onPlanToggle('${escapeHtml(r.id)}')">
        <span class="rtmy-plan-check" aria-hidden="true"></span>
        <span class="rtmy-plan-pick-body">
          <span class="rtmy-card-name">${escapeHtml(titleCase(r.name))}</span>
          <span class="rtmy-card-meta">${r.yieldAmount ?? ""} ${escapeHtml(r.yieldUnit || "")} · ${r.plan.steps.length} pasos${solo ? " · " + solo + " en solitario" : ""}</span>
        </span>
      </button>`;
  };

  const groups = planGroupBySection(withPlan)
    .map(
      (g) => `
      <section class="rtmy-section" aria-label="${escapeHtml(g.title)}">
        <h2 class="rtmy-section-title">${escapeHtml(g.title)}<span class="rtmy-section-count">${g.items.length}</span></h2>
        <div class="rtmy-section-items">${g.items.map(pick).join("")}</div>
      </section>`
    )
    .join("");

  const hours = [4, 5, 6, 7, 8, 9, 10, 12]
    .map((h) => `<option value="${h}" ${h === ui.shiftH ? "selected" : ""}>${h} h</option>`)
    .join("");

  root.innerHTML = `
    ${header(true, "#/", "Planificar jornada")}
    <p class="rtmy-plan-intro">Elige lo que se prepara hoy. Lo que no marques se da por hecho que ya está en stock.</p>

    <div class="rtmy-panel rtmy-plan-config">
      <h2 class="rtmy-plan-h">Jornada</h2>
      <div class="rtmy-plan-row">
        <span class="rtmy-plan-label" id="lbl-people">Personas cocinando</span>
        <div class="rtmy-stepper" role="group" aria-labelledby="lbl-people">
          <button type="button" onclick="onPlanPeople(-1)" aria-label="Una persona menos" ${ui.people <= 1 ? "disabled" : ""}>−</button>
          <output aria-live="polite">${ui.people}</output>
          <button type="button" onclick="onPlanPeople(1)" aria-label="Una persona más" ${ui.people >= 8 ? "disabled" : ""}>+</button>
        </div>
      </div>
      <div class="rtmy-plan-row">
        <label class="rtmy-plan-label" for="plan-start">Empieza a las</label>
        <input id="plan-start" class="rtmy-plan-input" type="time" value="${escapeHtml(ui.start)}" onchange="onPlanStart(this.value)">
      </div>
      <div class="rtmy-plan-row">
        <label class="rtmy-plan-label" for="plan-shift">Duración del turno</label>
        <select id="plan-shift" class="rtmy-plan-input" onchange="onPlanShift(this.value)">${hours}</select>
      </div>
    </div>

    ${
      withPlan.length
        ? groups
        : `<div class="rtmy-empty">Aún no hay recetas con datos de planificación.</div>`
    }

    ${
      missing.length
        ? `<details class="rtmy-plan-missing">
            <summary>Sin datos de planificación todavía (${missing.length})</summary>
            <p>Estas recetas se irán sumando cuando tengan su bloque de planificación en el Excel.</p>
            <ul>${missing.map((r) => `<li>${escapeHtml(titleCase(r.name))}</li>`).join("")}</ul>
          </details>`
        : ""
    }

    ${
      stockNames.length
        ? `<div class="rtmy-plan-note stock" role="status"><b>Se da por hecho que hay en stock:</b> ${stockNames.map(escapeHtml).join(", ")}. Si hoy hay que prepararlo, márcalo arriba.</div>`
        : ""
    }

    ${planJoinHtml()}

    <button class="rtmy-cta" onclick="goPlan()" ${n ? "" : "disabled"}>
      ${n ? `Ver el plan · ${n} receta${n === 1 ? "" : "s"} · ${ui.people} persona${ui.people === 1 ? "" : "s"}` : "Elige al menos una receta"}
    </button>
    ${n ? `<button class="rtmy-cta secondary" onclick="onPlanClear()">Quitar selección</button>` : ""}
  `;
}

window.onPlanToggle = function (id) {
  const ui = state.planUi;
  ui.selected = ui.selected.includes(id) ? ui.selected.filter((x) => x !== id) : ui.selected.concat(id);
  savePlanState();
  renderPlanSetup();
};
window.onPlanClear = function () {
  state.planUi.selected = [];
  savePlanState();
  renderPlanSetup();
};
window.onPlanPeople = function (delta) {
  const ui = state.planUi;
  ui.people = Math.max(1, Math.min(8, ui.people + delta));
  savePlanState();
  renderPlanSetup();
};
// Hora y duración se guardan sin repintar: así no se pierde el foco del campo.
window.onPlanStart = function (value) {
  if (!/^\d{2}:\d{2}$/.test(value || "")) return;
  state.planUi.start = value;
  savePlanState();
};
window.onPlanShift = function (value) {
  state.planUi.shiftH = Number(value) || 8;
  savePlanState();
};
window.goPlan = function () {
  if (!planSelectedValid().length) return;
  state.planUi.day = 1;
  goto("#/planificar/plan");
};

/* ----- Pantalla 2: el plan ----- */

function planTaskKey(t) {
  return t.recipeId + "|" + t.block + "|" + t.step;
}

function planColorMap(plan) {
  const map = {};
  plan.recipes.forEach((r, i) => (map[r.id] = PLAN_COLORS[i % PLAN_COLORS.length]));
  return map;
}

function planClock(plan, abs) {
  return RTMYPlanner.fmtClock(abs);
}

// "08:00–09:42"; si cruza de día: "día 1 08:00 → día 2 08:05".
function planRangeText(start, end) {
  const d1 = RTMYPlanner.dayOf(start);
  const d2 = RTMYPlanner.dayOf(end - 1);
  const a = RTMYPlanner.fmtClock(start);
  const b = RTMYPlanner.fmtClock(end);
  if (d1 === d2) return d1 > 1 ? `día ${d1} · ${a}–${b}` : `${a}–${b}`;
  return `día ${d1} ${a} → día ${d2} ${b}`;
}

function planWaitMin(t) {
  return t.type === "Activo + espera" ? t.totalMin - t.personMin : 0;
}

// Una tarjeta de paso dentro de la franja de una persona (o de "sin nadie").
function planCardHtml(t, plan, colors, view) {
  const key = planTaskKey(t);
  const doneBy = view.done[key];
  const done = doneBy !== undefined;
  // En una jornada compartida se dice quién lo marcó y quién cogió la tarea:
  // "hecho" a secas no vale cuando son cuatro en la cocina.
  const by = view.shared && done && doneBy !== view.slot ? view.names[doneBy] || "Persona " + (doneBy + 1) : "";
  const owner = view.shared ? view.owner[key] : undefined;
  const ownerText =
    owner === undefined || done ? "" : owner === view.slot ? "la tienes tú" : "la tiene " + (view.names[owner] || "Persona " + (owner + 1));
  const machine = t.person === null;
  const wait = planWaitMin(t);
  const dur = (t.assumed ? "~" : "") + RTMYPlanner.fmtDuration(machine ? t.totalMin : t.personMin);
  return `
    <li class="rtmy-plan-task ${machine ? "machine" : ""} ${done ? "done" : ""}" style="--rc:${colors[t.recipeId]}">
      <div class="rtmy-plan-head">
        <span class="rtmy-plan-hour">${planClock(plan, t.start)}</span>
        <span class="rtmy-plan-dur">${machine ? "hasta " + planClock(plan, t.end) + " · " + dur : dur}</span>
        <button class="rtmy-plan-tick" onclick="onPlanDone('${escapeHtml(key)}')" aria-pressed="${done}" aria-label="${done ? "Marcar como pendiente" : "Marcar como hecho"}">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>
        </button>
      </div>
      <div class="rtmy-plan-recipe-tag"><b>${escapeHtml(t.code)}</b> ${escapeHtml(titleCase(t.recipeName))} <span>· paso ${t.step}</span></div>
      <div class="rtmy-plan-tags">
        <span class="rtmy-plan-eq">${escapeHtml(t.equipment)}</span>
        ${machine && t.totalMin >= 5 ? `<span class="rtmy-plan-alone">en marcha · no hace falta estar</span>` : ""}
        ${by ? `<span class="rtmy-plan-by">hecho por ${escapeHtml(by)}</span>` : ""}
        ${ownerText ? `<span class="rtmy-plan-by">${escapeHtml(ownerText)}</span>` : ""}
        ${
          // Coger un paso que le toca a otro: el servidor se la da al primero
          // que la pide y al segundo le avisa en vez de quitarsela.
          !view.shared || done
            ? ""
            : owner === view.slot
              ? `<button class="rtmy-plan-claim" onclick="onPlanClaim('${escapeHtml(key)}', false)">Soltarla</button>`
              : owner === undefined && t.person !== null && t.person !== view.slot
                ? `<button class="rtmy-plan-claim" onclick="onPlanClaim('${escapeHtml(key)}', true)">La cojo yo</button>`
                : ""
        }
      </div>
      <p>${escapeHtml(t.text)}</p>
      ${
        machine
          ? // Un paso de menos de 2 min no necesita temporizador: se hace al momento.
            t.totalMin >= 2
            ? `<div class="rtmy-plan-wait"><button class="rtmy-plan-timer" onclick="planTimer(${t.id})" aria-label="Iniciar temporizador de ${dur}"><span aria-hidden="true">⏱</span> Temporizador de ${dur}</button></div>`
            : ""
          : wait
            ? `<div class="rtmy-plan-wait">Después, ${t.assumed ? "~" : ""}${RTMYPlanner.fmtDuration(wait)} de espera · listo a las ${planClock(plan, t.end)}
                 <button class="rtmy-plan-timer" onclick="planTimer(${t.id})" aria-label="Iniciar temporizador de ${RTMYPlanner.fmtDuration(wait)}"><span aria-hidden="true">⏱</span> Temporizador</button></div>`
            : ""
      }
    </li>`;
}

// Las recetas que toca una persona ese día, como chips de color.
function planWhoChips(tasks, plan, colors) {
  const seen = [];
  tasks.forEach((t) => {
    if (!seen.some((x) => x.id === t.recipeId)) seen.push({ id: t.recipeId, code: t.code, name: t.recipeName });
  });
  if (!seen.length) return "";
  // Con más de dos recetas solo cabe el código; el nombre va en cada tarjeta.
  const short = seen.length > 2;
  return `<div class="rtmy-plan-who ${short ? "short" : ""}">${seen
    .map(
      (r) =>
        `<span class="rtmy-plan-who-chip" style="--rc:${colors[r.id]}" title="${escapeHtml(titleCase(r.name))}"><i></i>${escapeHtml(r.code)}${short ? "" : " " + escapeHtml(titleCase(r.name))}</span>`
    )
    .join("")}</div>`;
}

/* Vista "por persona": una franja por cocinero, en horizontal. En el móvil se
   desliza de una a otra; en pantalla ancha se ven todas a la vez. Los pasos que
   hace sola la máquina salen en la franja de quien la pone en marcha, marcados
   aparte, para que sepa cuándo volver sin buscarlos en otro sitio. */
function planByPersonHtml(plan, colors) {
  const ui = state.planUi;
  const j = state.jornada;
  const view = { shared: planShared(), slot: j.slot, done: planDoneMap(), owner: planOwnerMap(), names: planPeopleNames() };
  const inDay = (t) => RTMYPlanner.dayOf(t.start) === ui.day;
  const cols = [];

  for (let p = 0; p < plan.persons; p++) {
    const mine = plan.tasks
      .filter((t) => (t.person === p || (t.person === null && t.watcher === p)) && inDay(t))
      .sort((x, y) => x.start - y.start || x.id - y.id);
    const total = plan.tasks.filter((t) => t.person === p).reduce((s, t) => s + t.personMin, 0);
    const waiting = mine.filter((t) => t.person === null).length;
    cols.push({
      person: p,
      mine: view.shared && p === j.slot,
      title: (view.names[p] ? view.names[p] : "Persona " + (p + 1)) + (view.shared && p === j.slot ? " · tú" : ""),
      sub: (total ? RTMYPlanner.fmtDuration(total) + " de trabajo" : "sin tareas") + (waiting ? " · " + waiting + " en marcha" : ""),
      chips: planWhoChips(mine, plan, colors),
      body: mine.length
        ? `<ol class="rtmy-plan-tasks">${mine.map((t) => planCardHtml(t, plan, colors, view)).join("")}</ol>`
        : `<p class="rtmy-plan-none">${total ? "Nada que hacer este día." : "Con este lote no hace falta una persona más."}</p>`,
    });
  }

  return `
    <div class="rtmy-plan-cols" style="--cols:${cols.length}">
      ${cols
        .map(
          (c) => `
        <section class="rtmy-plan-col ${c.mine ? "mine" : ""}" aria-label="${escapeHtml(c.title)}">
          <header class="rtmy-plan-col-head">
            <h2>${escapeHtml(c.title)}</h2>
            <span class="rtmy-plan-col-sub">${escapeHtml(c.sub)}</span>
            ${c.chips}
            <div class="rtmy-plan-timers" data-person="${c.person}" aria-live="polite"></div>
          </header>
          ${c.body}
        </section>`
        )
        .join("")}
    </div>
    ${cols.length > 1 ? `<p class="rtmy-plan-hint swipe">Desliza a los lados para ver cada franja.</p>` : ""}`;
}

function planTimelineHtml(plan, colors) {
  // Eje comprimido: los huecos largos sin actividad (la noche entre dos días)
  // se pliegan en una marca, y las esperas de horas se dibujan cruzándolos.
  const GAP_MIN = 180; // hueco a partir del cual se pliega
  const GAP_PX = 46; // ancho de la marca plegada
  const LONG = 240; // una espera más larga que esto no cuenta como actividad
  const items = [];
  plan.tasks.forEach((t) => items.push([t.start, t.end]));
  plan.spans.forEach((sp) => items.push([sp.start, sp.end]));
  const keep = [];
  items.forEach(([s, e]) => {
    if (e - s <= LONG) keep.push([s, e]);
    else keep.push([s, s + 5], [e - 5, e]);
  });
  keep.sort((x, y) => x[0] - y[0]);
  const segs = [];
  keep.forEach(([s, e]) => {
    const last = segs[segs.length - 1];
    if (last && s - last.b <= GAP_MIN) last.b = Math.max(last.b, e);
    else segs.push({ a: s, b: e });
  });
  const active = segs.reduce((sum, g) => sum + (g.b - g.a), 0);
  const px = Math.max(0.45, Math.min(4, 760 / Math.max(60, active)));
  let x = 0;
  segs.forEach((g) => {
    g.x0 = x;
    g.w = Math.round((g.b - g.a) * px);
    x += g.w + GAP_PX;
  });
  const W = x - GAP_PX;
  const pos = (m) => {
    let g = segs.find((sg) => m >= sg.a && m <= sg.b);
    if (!g) g = segs.filter((sg) => sg.b < m).pop() || segs[0];
    return Math.round(g.x0 + (Math.max(g.a, Math.min(m, g.b)) - g.a) * px);
  };
  const step = px * 60 >= 44 ? 60 : px * 120 >= 44 ? 120 : 360;

  let ticks = "";
  let grid = "";
  segs.forEach((g, i) => {
    for (let m = Math.ceil(g.a / step) * step; m <= g.b; m += step) {
      ticks += `<span class="rtmy-tl-tick" style="left:${pos(m)}px">${RTMYPlanner.fmtClock(m)}</span>`;
    }
    const d = RTMYPlanner.dayOf(g.a);
    const prevDay = i ? RTMYPlanner.dayOf(segs[i - 1].b) : 1;
    if (i && d !== prevDay) {
      grid += `<span class="rtmy-tl-day" style="left:${g.x0 - Math.round(GAP_PX / 2)}px"><i>Día ${d}</i></span>`;
    }
  });
  grid = `<div class="rtmy-tl-grid" style="width:${W}px">${grid}</div>`;

  const bar = (start, end, cls, color, label, title) => {
    const l = pos(start);
    const w = Math.max(5, pos(end) - l);
    return `<span class="rtmy-tl-bar ${cls}" style="left:${l}px;width:${w}px;--rc:${color}" title="${escapeHtml(title)}">${w >= 34 ? escapeHtml(label) : ""}</span>`;
  };

  const lanes = [];
  for (let p = 0; p < plan.persons; p++) {
    const bars = plan.tasks
      .filter((t) => t.person === p)
      .map((t) => {
        const title = `${RTMYPlanner.fmtClock(t.start)} · ${t.code} paso ${t.step}: ${t.text}`;
        let b = bar(t.start, t.start + t.personMin, "", colors[t.recipeId], `${t.code}.${t.step}`, title);
        if (planWaitMin(t)) b += bar(t.start + t.personMin, t.end, "wait", colors[t.recipeId], "", `Espera tras ${t.code}.${t.step}`);
        return b;
      })
      .join("");
    lanes.push({ label: "Persona " + (p + 1), bars });
  }
  const eqLabels = [...new Set(plan.spans.map((sp) => sp.label))].sort((a, b) => a.localeCompare(b, "es"));
  eqLabels.forEach((label) => {
    const bars = plan.spans
      .filter((sp) => sp.label === label)
      .map((sp) => bar(sp.start, sp.end, "equip", colors[sp.recipeId], sp.code, `${label}: ${sp.code} de ${RTMYPlanner.fmtClock(sp.start)} a ${RTMYPlanner.fmtClock(sp.end)}`))
      .join("");
    lanes.push({ label, bars, equip: true });
  });
  const free = plan.tasks.filter((t) => t.person === null && !RTMYPlanner.EQUIPMENT[t.equipment]);
  if (free.length) {
    lanes.push({
      label: "Reposos",
      equip: true,
      bars: free.map((t) => bar(t.start, t.end, "wait", colors[t.recipeId], `${t.code}.${t.step}`, `${t.code} paso ${t.step}: ${t.text}`)).join(""),
    });
  }

  const legend = plan.recipes
    .map((r) => `<span class="rtmy-tl-key"><i style="background:${colors[r.id]}"></i>${escapeHtml(r.code)} ${escapeHtml(titleCase(r.name))}</span>`)
    .join("");

  return `
    <div class="rtmy-tl-legend">${legend}</div>
    <div class="rtmy-tl-scroll" tabindex="0" aria-label="Línea de tiempo del plan">
      <div class="rtmy-tl" style="width:${W + 112}px">
        <div class="rtmy-tl-axis"><span class="rtmy-tl-label"></span><div class="rtmy-tl-track" style="width:${W}px">${ticks}</div></div>
        ${lanes
          .map(
            (l) => `
          <div class="rtmy-tl-lane ${l.equip ? "equip" : ""}">
            <span class="rtmy-tl-label">${escapeHtml(l.label)}</span>
            <div class="rtmy-tl-track" style="width:${W}px">${l.bars}</div>
          </div>`
          )
          .join("")}
        ${grid}
      </div>
    </div>
    <p class="rtmy-plan-hint">Desliza hacia los lados para recorrer la jornada. Barra llena: trabajo de una persona. Barra clara: espera sin nadie al lado. Las filas de equipo muestran cuándo está ocupado.${segs.length > 1 ? " Los huecos sin actividad entre días se han plegado." : ""}</p>`;
}

function planNotesHtml(plan) {
  const notes = [];
  if (plan.assumptions.length) {
    const codes = [...new Set(plan.assumptions.map((x) => x.code))];
    const list = plan.assumptions
      .map((x) => `<li>${escapeHtml(x.code)} · bloque ${x.block}, paso ${x.step} (${escapeHtml(x.type)})</li>`)
      .join("");
    notes.push(
      `<div class="rtmy-plan-note warn">
        <b>${plan.assumptions.length} paso${plan.assumptions.length === 1 ? "" : "s"} sin tiempo en la ficha</b> (${codes.map(escapeHtml).join(", ")}).
        Se han supuesto ${plan.options.assumedWaitMin} min de espera o ${plan.options.assumedMachineMin} min de máquina, así que las horas de esas recetas son orientativas (se marcan con ~).
        <details><summary>Ver cuáles</summary><ul>${list}</ul></details>
      </div>`
    );
  }
  if (plan.overflow) {
    const h = Math.floor(plan.overflow.shiftMinutes / 60);
    const m = plan.overflow.shiftMinutes % 60;
    notes.push(
      `<div class="rtmy-plan-note info"><b>${plan.overflow.steps} paso${plan.overflow.steps === 1 ? "" : "s"} no cabe${plan.overflow.steps === 1 ? "" : "n"} en el turno</b> y pasa${plan.overflow.steps === 1 ? "" : "n"} al día siguiente. Alargando el turno a ${h} h${m ? " " + m + " min" : ""} todo terminaría a las ${RTMYPlanner.fmtClock(plan.overflow.endIfLonger)} del mismo día.</div>`
    );
  }
  if (plan.stock.length) {
    const names = [...new Set(plan.stock.map((x) => titleCase(x.name)))];
    notes.push(`<div class="rtmy-plan-note stock"><b>En stock (no se prepara hoy):</b> ${names.map(escapeHtml).join(", ")}.</div>`);
  }
  const queues = Object.entries(plan.blocked).filter(([, v]) => v > 0);
  if (queues.some(([k]) => k === "Personas")) {
    notes.push(`<div class="rtmy-plan-note info">Hay pasos esperando a que quede una persona libre. Con una persona más se adelantaría el final.</div>`);
  }
  const eq = queues.filter(([k]) => k !== "Personas").sort((x, y) => y[1] - x[1]);
  if (eq.length) {
    notes.push(`<div class="rtmy-plan-note info"><b>Cola por equipo:</b> ${eq.slice(0, 3).map(([k]) => escapeHtml(k)).join(", ")}. Dos pasos necesitaban la misma máquina a la vez, así que uno espera.</div>`);
  }
  // Avisos del Excel. Los de "falta tiempo" ya están resumidos arriba, así que
  // no se repiten; el resto se pliega si son muchos.
  const warns = plan.warnings.filter((w) => !/falta tiempo para el tipo/i.test(w));
  if (warns.length) {
    const body = warns.map((w) => `<li>${escapeHtml(w)}</li>`).join("");
    notes.push(
      warns.length > 2
        ? `<div class="rtmy-plan-note warn"><b>${warns.length} avisos en los datos del Excel.</b><details><summary>Ver detalle</summary><ul>${body}</ul></details></div>`
        : `<div class="rtmy-plan-note warn"><ul class="bare">${body}</ul></div>`
    );
  }
  return notes.join("");
}

function renderPlan() {
  const ui = state.planUi;
  const j = state.jornada;
  // Enganchada a una jornada pero sin el plan todavía (al abrir la app).
  if (j.code && !ui.result) {
    root.innerHTML = `
      ${header(true, "#/planificar", "Plan de la jornada")}
      <div class="rtmy-panel rtmy-plan-share joined ${j.status === "offline" ? "offline" : ""}">
        <div class="rtmy-plan-kicker">Jornada</div>
        <div class="rtmy-plan-code">${escapeHtml(j.code)}</div>
        <p class="rtmy-plan-share-help">${j.status === "offline" ? "Sin conexión con el servidor. Se reintenta solo." : "Bajando el plan de la jornada…"}</p>
        ${j.error ? `<p class="rtmy-plan-share-error">${escapeHtml(j.error)}</p>` : ""}
        <button class="rtmy-plan-reset" onclick="onJornadaLeave()">Salir de la jornada (este dispositivo)</button>
      </div>`;
    return;
  }
  const plan = planCurrent();
  if (!plan.ok) {
    root.innerHTML = `
      ${header(true, "#/planificar", "Plan de la jornada")}
      ${j.error ? `<p class="rtmy-plan-share-error">${escapeHtml(j.error)}</p>` : ""}
      <div class="rtmy-panel"><p><strong>No se pudo montar el plan.</strong></p>
      <p style="font-size:13px;opacity:.75">${escapeHtml(plan.warnings.join(" ") || "Elige al menos una receta con datos de planificación.")}</p></div>
      ${planJoinHtml()}`;
    return;
  }
  const colors = planColorMap(plan);
  if (ui.day > plan.days) ui.day = 1;
  const endDay = RTMYPlanner.dayOf(plan.end - 1);
  const done = planDoneMap();
  const doneCount = plan.tasks.filter((t) => done[planTaskKey(t)] !== undefined).length;

  const recipesHtml = plan.recipes
    .map(
      (r) => `
      <li style="--rc:${colors[r.id]}">
        <span class="rtmy-plan-dot" aria-hidden="true"></span>
        <span class="rtmy-plan-recipe-name"><b>${escapeHtml(r.code)}</b> ${escapeHtml(titleCase(r.name))}</span>
        <span class="rtmy-plan-recipe-time">${planRangeText(r.start, r.end)}</span>
        <a class="rtmy-plan-open" href="#/receta/${encodeURIComponent(r.id)}">Abrir receta</a>
      </li>`
    )
    .join("");

  const dayTabs =
    plan.days > 1
      ? `<div class="rtmy-plan-tabs days" role="tablist" aria-label="Día">${Array.from({ length: plan.days }, (_, i) => i + 1)
          .map((d) => `<button role="tab" aria-selected="${d === ui.day}" class="${d === ui.day ? "active" : ""}" onclick="onPlanDay(${d})">Día ${d}</button>`)
          .join("")}</div>`
      : "";

  root.innerHTML = `
    ${header(true, "#/planificar", "Plan de la jornada")}

    <div class="rtmy-panel rtmy-plan-summary">
      <div class="rtmy-plan-kicker">${endDay > 1 ? "Termina el día " + endDay + " a las" : "Todo listo a las"}</div>
      <div class="rtmy-plan-endtime">${RTMYPlanner.fmtClock(plan.end)}</div>
      <div class="rtmy-plan-meta">${
        // Con varios días, la duración total incluiría la noche: engaña más que informa.
        endDay > 1
          ? `Empieza a las ${RTMYPlanner.fmtClock(plan.start)}`
          : `${RTMYPlanner.fmtDuration(plan.end - plan.start)} desde las ${RTMYPlanner.fmtClock(plan.start)}`
      } · ${plan.persons} persona${plan.persons === 1 ? "" : "s"} · ${plan.tasks.length} pasos</div>
      <ul class="rtmy-plan-recipes">${recipesHtml}</ul>
      <div class="rtmy-plan-progress" role="progressbar" aria-valuemin="0" aria-valuemax="${plan.tasks.length}" aria-valuenow="${doneCount}"><span style="width:${Math.round((doneCount / plan.tasks.length) * 100)}%"></span></div>
      <div class="rtmy-plan-progress-text">${doneCount} de ${plan.tasks.length} pasos hechos</div>
    </div>

    ${planShareHtml(plan)}
    ${planNotesHtml(plan)}

    <div class="rtmy-plan-tabs" role="tablist" aria-label="Vista del plan">
      <button role="tab" aria-selected="${ui.view === "persona"}" class="${ui.view === "persona" ? "active" : ""}" onclick="onPlanView('persona')">Por persona</button>
      <button role="tab" aria-selected="${ui.view === "linea"}" class="${ui.view === "linea" ? "active" : ""}" onclick="onPlanView('linea')">Línea de tiempo</button>
    </div>
    ${ui.view === "persona" ? dayTabs : ""}
    ${ui.view === "persona" ? planByPersonHtml(plan, colors) : planTimelineHtml(plan, colors)}

    <button class="rtmy-cta secondary" onclick="goto('#/planificar')">Cambiar recetas o personas</button>
    <button class="rtmy-plan-reset" onclick="onPlanResetDone()">Borrar lo marcado como hecho</button>
  `;
  syncPlanSwipeHint();
  renderPlanTimers();
}

/* Los temporizadores del plan se muestran en la franja de quien los puso, para
   no confundirlos con los de otra persona. Se repintan cada segundo sin tocar
   el resto de la vista, para no perder el desplazamiento de la franja. */
function renderPlanTimers() {
  const slots = root.querySelectorAll(".rtmy-plan-timers[data-person]");
  if (!slots.length) return;
  // Los temporizadores de máquina que han puesto los demás: se ven, con su
  // cuenta atrás en hora del servidor, pero no se tocan desde aquí.
  const shared = planSharedTimers();
  const names = planPeopleNames();
  const now = Date.now();
  slots.forEach((slot) => {
    const person = Number(slot.dataset.person);
    const mine = state.timers.filter((t) => t.person === person);
    const theirs = shared.filter((t) => t.b === person);
    slot.innerHTML =
      mine
        .map(
          (t) => `
        <div class="rtmy-plan-timer-chip ${t.done ? "done" : ""}">
          <span class="clock">${t.done ? "¡Listo!" : fmtClock(t.remaining)}</span>
          <span class="what">${escapeHtml(t.label)}</span>
          ${t.done ? "" : `<button onclick="toggleTimer(${t.id})" aria-label="${t.running ? "Pausar" : "Reanudar"}">${t.running ? "⏸" : "▶️"}</button>`}
          <button onclick="cancelTimer(${t.id})" aria-label="Descartar temporizador">✕</button>
        </div>`
        )
        .join("") +
      theirs
        .map((t) => {
          const left = RTMYJornada.remainingSeconds(t, state.jornada.offset, now);
          return `
        <div class="rtmy-plan-timer-chip remote ${left ? "" : "done"}">
          <span class="clock">${left ? fmtClock(left) : "¡Listo!"}</span>
          <span class="what">${escapeHtml(t.l || "")} · ${escapeHtml(names[t.b] || "Persona " + (t.b + 1))}</span>
        </div>`;
        })
        .join("");
    slot.hidden = !mine.length && !theirs.length;
  });
}

// La pista "desliza a los lados" solo tiene sentido si alguna franja se sale.
// Se vigila con ResizeObserver: al girar el móvil o cambiar el ancho, el
// evento "resize" llega antes de que el navegador recalcule el ancho real.
let planColsObserver = null;
function syncPlanSwipeHint() {
  const cols = root.querySelector(".rtmy-plan-cols");
  const hint = root.querySelector(".rtmy-plan-hint.swipe");
  if (planColsObserver) {
    planColsObserver.disconnect();
    planColsObserver = null;
  }
  if (!hint || !cols) return;
  const update = () => {
    hint.hidden = cols.scrollWidth <= cols.clientWidth + 2;
  };
  update();
  requestAnimationFrame(update); // tras aplicar estilos y media queries
  if (typeof ResizeObserver === "function") {
    planColsObserver = new ResizeObserver(() => requestAnimationFrame(update));
    planColsObserver.observe(cols);
    planColsObserver.observe(document.documentElement);
  }
  // Red de seguridad al girar el móvil, por si el observador no llega.
  window.addEventListener("resize", () => requestAnimationFrame(update), { passive: true });
  window.addEventListener("orientationchange", () => setTimeout(update, 180), { passive: true });
}

window.onPlanView = function (view) {
  state.planUi.view = view;
  savePlanState();
  renderPlan();
};
window.onPlanDay = function (day) {
  state.planUi.day = day;
  renderPlan();
};
window.onPlanDone = function (key) {
  if (planShared()) {
    // Interfaz optimista: se pinta ya y se manda por detrás. Con 1-3 s de
    // latencia, esperar la respuesta se nota y molesta.
    const wasDone = planDoneMap()[key] !== undefined;
    jornadaQueue({ op: wasDone ? "undone" : "done", key, slot: state.jornada.slot });
    renderPlan();
    jornadaFlush();
    return;
  }
  const done = state.planUi.done;
  if (done[key]) delete done[key];
  else done[key] = 1;
  savePlanState();
  renderPlan();
};
// Coger (o soltar) un paso que el plan le ha dado a otra persona.
window.onPlanClaim = function (key, take) {
  if (!planShared()) return;
  jornadaQueue({ op: take ? "claim" : "release", key, slot: state.jornada.slot });
  renderPlan();
  jornadaFlush();
};
window.onPlanResetDone = function () {
  if (planShared()) {
    const done = planDoneMap();
    Object.keys(done).forEach((key) => jornadaQueue({ op: "undone", key }, true));
    saveJornadaState();
    renderPlan();
    jornadaFlush();
    return;
  }
  state.planUi.done = {};
  savePlanState();
  renderPlan();
};
// Lanza el temporizador de un paso del plan en la barra de abajo.
window.planTimer = function (taskId) {
  const plan = state.planUi.result;
  const t = plan && plan.tasks.find((x) => x.id === taskId);
  if (!t) return;
  const minutes = t.type === "Activo + espera" ? t.totalMin - t.personMin : t.totalMin;
  const short = t.text.length > 34 ? t.text.slice(0, 33).trimEnd() + "…" : t.text;
  const who = t.person !== null && t.person !== undefined ? t.person : t.watcher;
  const label = `${t.code}.${t.step} · ${short}`;
  const seconds = Math.max(1, Math.round(minutes * 60));
  const timer = startTimer(label, seconds, null, t.recipeName, who);
  // Los temporizadores de máquina son de la cocina, no de la tablet: al arroz
  // le da igual quién lo programó, así que se comparten.
  if (timer && planShared()) {
    timer.shared = state.jornada.slot + "-" + timer.id;
    jornadaQueue({ op: "timer", id: timer.shared, key: planTaskKey(t), label, seconds });
    saveRuntimeState();
    jornadaFlush();
  }
};

/* ---------------- Jornada compartida (interfaz) ----------------
   Varios dispositivos sobre el mismo plan. Sin código de jornada todo
   funciona como antes: el plan se calcula aquí y lo marcado se guarda en este
   dispositivo. Con código, el plan y lo marcado vienen del servidor.
   El contrato está en JORNADA_COMPARTIDA.md; el transporte, en jornada.js.
------------------------------------------------- */

function jornadaAvailable() {
  return !!(CONFIG.DATA_URL && window.RTMYJornada);
}

// ¿Se está trabajando sobre un plan compartido?
function planShared() {
  return !!(state.jornada.code && state.planUi.result && state.planUi.result.shared);
}

// El plan que hay que pintar: el compartido si lo hay, o el que calcula este
// dispositivo. Al ser determinista, los dos salen iguales con los mismos datos.
function planCurrent() {
  return state.planUi.result || computePlan();
}

// Lo marcado como hecho: {clave: slot de quien lo marcó}. En solitario el slot
// no significa nada, pero así la vista es la misma en los dos modos.
function planDoneMap() {
  const j = state.jornada;
  if (planShared()) return RTMYJornada.applyQueue(j.live, j.queue, j.slot).done;
  const out = {};
  Object.keys(state.planUi.done).forEach((key) => (out[key] = 0));
  return out;
}

function planOwnerMap() {
  const j = state.jornada;
  return planShared() ? RTMYJornada.applyQueue(j.live, j.queue, j.slot).owner : {};
}

function planPeopleNames() {
  return planShared() ? RTMYJornada.peopleNames(state.jornada.live) : {};
}

// Los temporizadores que han puesto los demás (los míos ya están en la barra).
function planSharedTimers() {
  const j = state.jornada;
  if (!planShared() || !j.live) return [];
  return (j.live.timers || []).filter((t) => t.b !== j.slot);
}

let jornadaClientRef = null;
function jornadaClient() {
  if (!jornadaClientRef) jornadaClientRef = RTMYJornada.createClient({ url: CONFIG.DATA_URL });
  return jornadaClientRef;
}

function saveJornadaState() {
  try {
    const j = state.jornada;
    localStorage.setItem(
      CONFIG.JORNADA_KEY,
      JSON.stringify({ code: j.code, slot: j.slot, name: j.name, rev: j.rev, planRev: j.planRev, offset: j.offset, queue: j.queue, inputs: j.inputs, dataStamp: j.dataStamp })
    );
  } catch (e) {
    console.warn("No se pudo guardar la jornada compartida.", e);
  }
}

function restoreJornadaState() {
  try {
    const saved = JSON.parse(localStorage.getItem(CONFIG.JORNADA_KEY) || "null");
    if (!saved) return;
    const j = state.jornada;
    // El nombre se guarda aunque no haya jornada: el cocinero no tiene que
    // volver a escribirlo cada mañana.
    if (typeof saved.name === "string") j.name = saved.name;
    if (!/^\d{4}$/.test(saved.code || "")) return;
    j.code = saved.code;
    j.slot = Number.isFinite(saved.slot) ? saved.slot : 0;
    j.name = typeof saved.name === "string" ? saved.name : "";
    j.rev = Number(saved.rev) || 0;
    j.planRev = Number(saved.planRev) || 0;
    j.offset = Number(saved.offset) || 0;
    j.queue = Array.isArray(saved.queue) ? saved.queue : [];
    j.inputs = saved.inputs || null;
    j.dataStamp = saved.dataStamp || null;
  } catch (e) {
    console.warn("No se pudo restaurar la jornada compartida.", e);
  }
}

function jornadaQueue(op, silent) {
  const j = state.jornada;
  j.queue = RTMYJornada.queuePush(j.queue, op);
  if (!silent) saveJornadaState();
}

// Lo que trae cualquier respuesta del servidor: su hora, el estado y si el
// plan ha cambiado porque alguien lo ha recalculado.
function jornadaApply(res) {
  const j = state.jornada;
  if (typeof res.offset === "number") j.offset = res.offset;
  if (typeof res.rev === "number") j.rev = res.rev;
  if (res.live) j.live = res.live;
  j.status = "ok";
  j.error = null;
  if (res.meta) {
    j.inputs = res.meta.inputs || null;
    j.dataStamp = res.meta.dataStamp || null;
    j.staleData = !!(j.dataStamp && state.updatedAt && j.dataStamp !== state.updatedAt);
  }
  if (res.plan) {
    state.planUi.result = RTMYJornada.unpackPlan(res.plan, state.recipes);
    if (state.planUi.day > (state.planUi.result.days || 1)) state.planUi.day = 1;
  }
  if (typeof res.planRev === "number") {
    const changed = j.planRev && res.planRev !== j.planRev && !res.plan;
    j.planRev = res.planRev;
    if (changed) jornadaReloadPlan();
  }
}

function jornadaFail(err) {
  const j = state.jornada;
  if (err && err.data && err.data.missing) {
    // Una jornada se cierra sola cuando pasa día y medio sin usarse: la tablet
    // que vuelve al día siguiente con el código de ayer tiene que enterarse,
    // no quedarse en silencio con su plan de siempre.
    jornadaLeaveLocal("La jornada " + j.code + " ya se cerró: se cierran solas cuando pasa un día y medio sin usarse.", true);
    return;
  }
  j.status = "offline";
  // Al cocinero no se le enseña "Failed to fetch": en una cocina el wifi se cae
  // y lo único que importa es que no se ha perdido nada.
  const raw = (err && (err.name === "AbortError" ? "timeout" : err.message)) || "";
  j.error = /timeout|fetch|network|load failed/i.test(raw)
    ? "Sin conexión con el servidor. Lo marcado se guarda y se manda en cuanto vuelva."
    : raw || "Error al hablar con el servidor.";
}

let jornadaBusy = false;

// Manda lo que haya en la cola. Si falla, se queda en la cola y se reintenta
// en el siguiente sondeo: no se pierde lo que marcó el cocinero.
async function jornadaFlush() {
  const j = state.jornada;
  if (jornadaBusy || !j.code || !j.queue.length || !jornadaAvailable()) return;
  jornadaBusy = true;
  const sending = j.queue.slice();
  try {
    const res = await jornadaClient().patch({ code: j.code, slot: j.slot, name: j.name, ops: sending });
    j.queue = j.queue.filter((op) => sending.indexOf(op) < 0);
    jornadaApply(res);
    jornadaRejected(res.rejected);
  } catch (err) {
    jornadaFail(err);
  } finally {
    jornadaBusy = false;
    saveJornadaState();
    renderPlanKeepingScroll();
  }
}

// Coger una tarea que ya tiene dueño se rechaza en el servidor: se avisa, no
// se roba.
function jornadaRejected(list) {
  if (!list || !list.length) return;
  const names = planPeopleNames();
  const clash = list.filter((r) => r.op === "claim");
  if (clash.length) {
    const who = names[clash[0].owner] || "Persona " + (clash[0].owner + 1);
    state.jornada.error = "Esa tarea ya la tiene " + who + ".";
  }
}

async function jornadaPoll() {
  const j = state.jornada;
  if (jornadaBusy || !j.code || !jornadaAvailable()) return;
  if (document.hidden) return; // tablet en reposo: ni se sondea ni se gasta cuota
  if (j.queue.length) return jornadaFlush();
  jornadaBusy = true;
  try {
    const res = await jornadaClient().read(j.code, j.rev);
    const changed = !res.unchanged;
    jornadaApply(res);
    if (changed) renderPlanKeepingScroll();
    else if (state.view === "plan") renderPlanTimers();
  } catch (err) {
    jornadaFail(err);
    renderPlanKeepingScroll();
  } finally {
    jornadaBusy = false;
  }
}

async function jornadaReloadPlan() {
  const j = state.jornada;
  if (!j.code) return;
  try {
    const res = await jornadaClient().read(j.code, null, true);
    jornadaApply(res);
    renderPlanKeepingScroll();
  } catch (err) {
    jornadaFail(err);
  }
}

let jornadaTimer = null;
function jornadaWatch() {
  if (jornadaTimer || !state.jornada.code) return;
  jornadaTimer = setInterval(jornadaPoll, CONFIG.JORNADA_POLL_MS);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) jornadaPoll();
  });
}

function jornadaUnwatch() {
  if (jornadaTimer) clearInterval(jornadaTimer);
  jornadaTimer = null;
}

function jornadaLeaveLocal(message, expired) {
  const j = state.jornada;
  j.expired = !!expired;
  jornadaUnwatch();
  j.code = null;
  j.live = null;
  j.queue = [];
  j.rev = 0;
  j.planRev = 0;
  j.inputs = null;
  j.staleData = false;
  j.status = "off";
  j.error = message || null;
  state.planUi.result = null;
  saveJornadaState();
}

/* ----- Acciones ----- */

// Crea la jornada con el plan que hay en pantalla y reparte el código.
window.onJornadaCreate = async function (replace) {
  const j = state.jornada;
  if (!jornadaAvailable()) return;
  const plan = replace ? computePlan() : planCurrent();
  if (!plan.ok) return;
  j.status = "working";
  j.error = null;
  renderPlan();
  try {
    const body = {
      code: replace ? j.code : null,
      replace: !!replace,
      name: j.name,
      slot: replace ? j.slot : undefined,
      inputs: { selected: planSelectedValid(), people: state.planUi.people, start: state.planUi.start, shiftH: state.planUi.shiftH },
      dataStamp: state.updatedAt || null,
      plannerVersion: RTMYPlanner.VERSION,
      plan: RTMYJornada.packPlan(plan),
    };
    const res = await jornadaClient().create(body);
    j.code = res.code;
    j.slot = res.slot;
    j.planRev = res.planRev;
    j.inputs = body.inputs;
    j.dataStamp = body.dataStamp;
    j.staleData = false;
    jornadaApply(res);
    // El que crea la jornada ya tiene el plan calculado: se queda con el suyo
    // en el formato compartido, para que las dos pantallas sean idénticas.
    state.planUi.result = RTMYJornada.unpackPlan(RTMYJornada.packPlan(plan), state.recipes);
    saveJornadaState();
    jornadaWatch();
    jornadaPoll();
  } catch (err) {
    jornadaFail(err);
  }
  renderPlan();
};

window.onJornadaJoin = async function () {
  const j = state.jornada;
  const code = String(j.joinCode || "").replace(/\D/g, "");
  if (!jornadaAvailable() || !/^\d{4}$/.test(code)) {
    j.error = "El código son cuatro números.";
    render();
    return;
  }
  j.status = "working";
  j.error = null;
  render();
  try {
    const res = await jornadaClient().join({ code, name: j.name });
    j.code = res.code;
    j.slot = res.slot;
    j.planRev = res.planRev || 0;
    j.joinCode = "";
    j.queue = [];
    jornadaApply(res);
    saveJornadaState();
    jornadaWatch();
    if (state.view === "plan-setup") {
      goto("#/planificar/plan");
      return;
    }
  } catch (err) {
    j.code = null;
    jornadaFail(err);
  }
  render();
};

// Al volver a abrir la app se vuelve a entrar en la jornada guardada.
async function jornadaResume() {
  const j = state.jornada;
  if (!j.code || !jornadaAvailable()) return;
  j.status = "working";
  try {
    const res = await jornadaClient().join({ code: j.code, name: j.name, slot: j.slot });
    j.slot = res.slot;
    j.planRev = res.planRev || 0;
    jornadaApply(res);
    saveJornadaState();
    jornadaWatch();
    jornadaFlush();
  } catch (err) {
    jornadaFail(err);
  }
  if (state.view === "plan") renderPlan();
}

window.onJornadaLeave = function () {
  jornadaLeaveLocal(null, false);
  render();
};

window.onJornadaName = function (value) {
  const j = state.jornada;
  j.name = String(value || "").trim().slice(0, 24);
  saveJornadaState();
  if (j.code) {
    jornadaQueue({ op: "ping", slot: j.slot });
    jornadaFlush();
  }
};

window.onJornadaCodeInput = function (value) {
  state.jornada.joinCode = String(value || "").replace(/\D/g, "").slice(0, 4);
};

window.onJornadaSlot = async function (value) {
  const j = state.jornada;
  const slot = Math.max(0, Math.min(7, Number(value) || 0));
  if (!j.code || slot === j.slot) return;
  j.slot = slot;
  saveJornadaState();
  try {
    jornadaApply(await jornadaClient().join({ code: j.code, name: j.name, slot }));
  } catch (err) {
    jornadaFail(err);
  }
  renderPlan();
};

window.onJornadaPanel = function (open) {
  state.jornada.panel = !!open;
  state.jornada.error = null;
  state.jornada.expired = false;
  renderPlan();
};

/* ----- La franja de la jornada, arriba del plan ----- */

function jornadaStatusText() {
  const j = state.jornada;
  if (j.status === "working") return "Conectando…";
  if (j.queue.length) return j.queue.length + " cambio" + (j.queue.length === 1 ? "" : "s") + " por enviar";
  if (j.status === "offline") return "Sin conexión · se reintenta solo";
  return "Al día";
}

// ¿La selección de pantalla 1 ya no es la del plan compartido?
function jornadaInputsChanged() {
  const j = state.jornada;
  // Una tablet que solo se enganchó no tiene selección propia: a ella no se le
  // propone recalcular nada.
  if (!j.inputs || !planSelectedValid().length) return false;
  const now = { selected: planSelectedValid(), people: state.planUi.people, start: state.planUi.start, shiftH: state.planUi.shiftH };
  return JSON.stringify(now) !== JSON.stringify({ selected: j.inputs.selected || [], people: j.inputs.people, start: j.inputs.start, shiftH: j.inputs.shiftH });
}

/* En la pantalla de elegir recetas: una tablet que llega a media mañana solo
   necesita el código, no elegir nada. */
function planJoinHtml() {
  const j = state.jornada;
  if (!jornadaAvailable()) return "";
  if (j.code) {
    return `
      <div class="rtmy-plan-note share">
        <b>Estás en la jornada ${escapeHtml(j.code)}.</b> El plan es el de todos: si cambias la selección, habrá que recalcularlo para los demás.
        <button class="rtmy-plan-timer" onclick="goto('#/planificar/plan')">Ver el plan de la jornada</button>
      </div>`;
  }
  return `
    ${j.expired ? `<div class="rtmy-plan-note share gone"><b>${escapeHtml(j.error || "La jornada anterior ya se cerró.")}</b> Esta tablet ha vuelto a trabajar por su cuenta.</div>` : ""}
    <div class="rtmy-panel rtmy-plan-share">
      <h2 class="rtmy-plan-h">¿Ya hay una jornada en marcha?</h2>
      <p class="rtmy-plan-share-help">Escribe su código y verás el mismo plan que las demás tablets, sin elegir nada aquí.</p>
      <div class="rtmy-plan-row">
        <label class="rtmy-plan-label" for="j-name-setup">Tu nombre</label>
        <input id="j-name-setup" class="rtmy-plan-input" type="text" maxlength="24" placeholder="Leo" value="${escapeHtml(j.name)}" onchange="onJornadaName(this.value)">
      </div>
      <div class="rtmy-plan-row">
        <label class="rtmy-plan-label" for="j-code-setup">Código de jornada</label>
        <input id="j-code-setup" class="rtmy-plan-input code" type="text" inputmode="numeric" maxlength="4" placeholder="0000" value="${escapeHtml(j.joinCode)}" oninput="onJornadaCodeInput(this.value)">
      </div>
      <button class="rtmy-cta secondary" onclick="onJornadaJoin()" ${j.status === "working" ? "disabled" : ""}>Entrar</button>
      ${j.error ? `<p class="rtmy-plan-share-error">${escapeHtml(j.error)}</p>` : ""}
    </div>`;
}

function planShareHtml(plan) {
  const j = state.jornada;
  if (!jornadaAvailable()) return "";
  const nameField = `
    <label class="rtmy-plan-label" for="j-name">Tu nombre</label>
    <input id="j-name" class="rtmy-plan-input" type="text" maxlength="24" placeholder="Leo" value="${escapeHtml(j.name)}" onchange="onJornadaName(this.value)">`;

  if (!j.code) {
    if (!j.panel) {
      // Si la jornada de ayer ya no existe, el plan que se está viendo es solo
      // de esta tablet aunque parezca el mismo: hay que decirlo.
      if (j.expired) {
        return `
          <div class="rtmy-plan-note share gone">
            <b>${escapeHtml(j.error || "La jornada anterior ya se cerró.")}</b>
            Este plan vuelve a ser solo de esta tablet: lo que marques aquí no lo ven las demás.
            <button class="rtmy-plan-timer" onclick="onJornadaPanel(true)">Crear una jornada nueva</button>
          </div>`;
      }
      return `
        <div class="rtmy-plan-note share">
          <b>Cada tablet va por su cuenta.</b> Para que varios trabajéis sobre este mismo plan, crea un código de jornada.
          <button class="rtmy-plan-timer" onclick="onJornadaPanel(true)">Trabajar en equipo</button>
        </div>`;
    }
    return `
      <div class="rtmy-panel rtmy-plan-share">
        <h2 class="rtmy-plan-h">Trabajar en equipo</h2>
        <div class="rtmy-plan-row">${nameField}</div>
        <button class="rtmy-cta" onclick="onJornadaCreate(false)" ${j.status === "working" ? "disabled" : ""}>Crear la jornada con este plan</button>
        <p class="rtmy-plan-share-or">o entra en una que ya existe</p>
        <div class="rtmy-plan-row">
          <label class="rtmy-plan-label" for="j-code">Código de jornada</label>
          <input id="j-code" class="rtmy-plan-input code" type="text" inputmode="numeric" maxlength="4" placeholder="0000" value="${escapeHtml(j.joinCode)}" oninput="onJornadaCodeInput(this.value)">
        </div>
        <button class="rtmy-cta secondary" onclick="onJornadaJoin()" ${j.status === "working" ? "disabled" : ""}>Entrar</button>
        ${j.error ? `<p class="rtmy-plan-share-error">${escapeHtml(j.error)}</p>` : ""}
        <button class="rtmy-plan-reset" onclick="onJornadaPanel(false)">Dejarlo para después</button>
      </div>`;
  }

  const names = planPeopleNames();
  const people = ((j.live && j.live.people) || [])
    .slice()
    .sort((a, b) => a.s - b.s)
    .map((p) => `<span class="rtmy-plan-who-chip ${p.s === j.slot ? "me" : ""}"><i></i>${escapeHtml(p.n || "Persona " + (p.s + 1))}</span>`)
    .join("");
  const slots = Array.from({ length: plan.persons || 1 }, (_, i) => i)
    .map((i) => `<option value="${i}" ${i === j.slot ? "selected" : ""}>${escapeHtml(names[i] || "Persona " + (i + 1))}</option>`)
    .join("");

  return `
    <div class="rtmy-panel rtmy-plan-share joined ${j.status === "offline" ? "offline" : ""}">
      <div class="rtmy-plan-share-head">
        <div>
          <div class="rtmy-plan-kicker">Jornada</div>
          <div class="rtmy-plan-code">${escapeHtml(j.code)}</div>
        </div>
        <div class="rtmy-plan-share-state">
          <span class="dot" aria-hidden="true"></span>${escapeHtml(jornadaStatusText())}
        </div>
      </div>
      <p class="rtmy-plan-share-help">Dicta este código en las demás tablets para que vean el mismo plan.</p>
      <div class="rtmy-plan-row">${nameField}</div>
      <div class="rtmy-plan-row">
        <label class="rtmy-plan-label" for="j-slot">Eres</label>
        <select id="j-slot" class="rtmy-plan-input" onchange="onJornadaSlot(this.value)">${slots}</select>
      </div>
      ${people ? `<div class="rtmy-plan-who">${people}</div>` : ""}
      ${j.staleData ? `<p class="rtmy-plan-share-error">Este plan se calculó con otra versión de las recetas. Dale a "Actualizar" en la lista para igualar los datos.</p>` : ""}
      ${plan.missing && plan.missing.length ? `<p class="rtmy-plan-share-error">${plan.missing.length} paso${plan.missing.length === 1 ? "" : "s"} del plan no está${plan.missing.length === 1 ? "" : "n"} en los datos de esta tablet: se ven la hora y el aparato, pero no el texto.</p>` : ""}
      ${jornadaInputsChanged() ? `<button class="rtmy-cta" onclick="onJornadaCreate(true)">Has cambiado la selección · recalcular el plan para todos</button>` : ""}
      ${j.error ? `<p class="rtmy-plan-share-error">${escapeHtml(j.error)}</p>` : ""}
      <button class="rtmy-plan-reset" onclick="onJornadaLeave()">Salir de la jornada (este dispositivo)</button>
    </div>`;
}

// Repinta el plan sin perder dónde estaba mirando el cocinero: el sondeo llega
// cada 10 s y no puede mover la pantalla a nadie.
function renderPlanKeepingScroll() {
  if (state.view !== "plan") return;
  const cols = root.querySelector(".rtmy-plan-cols");
  const left = cols ? cols.scrollLeft : 0;
  const top = window.scrollY;
  const active = document.activeElement;
  const focusId = active && active.id ? active.id : null;
  renderPlan();
  const after = root.querySelector(".rtmy-plan-cols");
  if (after) after.scrollLeft = left;
  window.scrollTo(0, top);
  if (focusId) {
    const back = document.getElementById(focusId);
    if (back) back.focus();
  }
}

/* ---------------- Boot ---------------- */

window.goto = goto;
window.startFlow = startFlow;
window.loadRecipes = loadRecipes;
window.refreshRecipes = refreshRecipes;

window.addEventListener("hashchange", onHashChange);

loadRecipes().then(() => {
  restoreRuntimeState();
  // Las preparaciones abiertas se restauran DESPUÉS de cargar las recetas,
  // así que aquí es donde hay que engancharlas a la ficha de verdad.
  relinkRecipeRefs();
  restorePlanState();
  restoreJornadaState();
  if (state.timers.some((timer) => timer.running && !timer.done)) ensureTimerTick();
  renderTimerBar();
  renderSessionBar();
  onHashChange();
  // Si este dispositivo estaba en una jornada, se vuelve a enganchar y manda
  // lo que quedara en la cola de cuando se cayó la red.
  jornadaResume();
});
