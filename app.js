/* ============================================================
   RICE TO MEET YOU — Recetas del Taller
   App de una sola página: lista -> receta -> escalado -> pasos -> final.
   Fuente de datos: JSON servido por un Google Apps Script conectado
   en vivo al Excel "MANUAL COCINA" (ver CONFIG.DATA_URL).
   ============================================================ */

const CONFIG = {
  // Sustituir por la URL de despliegue del Apps Script (termina en /exec).
  // Mientras esté vacío, la app usa recipes.json local (modo desarrollo/demo).
  DATA_URL: "https://script.google.com/macros/s/AKfycbylwL_HoE5bLKaAnKPMILZOD96h_hel1b5Uhuu5ht0JzvbRisBGgUQy4OWg7tiMACpY/exec",
  LOCAL_FALLBACK: "recipes.json",
  CACHE_KEY: "rtmy_recipes_cache_v2",
  RUNTIME_KEY: "rtmy_runtime_v2",
  CACHE_TTL_MS: 1000 * 60 * 10, // 10 minutos
};

const MAX_SESSIONS = 4;

const state = {
  view: "loading", // loading | error | list | detail | steps | final
  recipes: [],
  updatedAt: null,
  dataSource: null,
  search: "",
  refreshing: false,
  timers: [],
  currentStepDurations: [],

  // Detail screen: a lightweight "draft" scale for whichever recipe is
  // currently being looked at, before (or while) it's being cooked.
  currentRecipe: null,
  scaleFactor: 1,
  scaleQty: null,
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

async function fetchRecipesFrom(url, source) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error("HTTP " + res.status);
  const data = validateRecipesPayload(await res.json());
  return { data, source };
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

function sortRecipesByName(recipes) {
  return recipes
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name, "es", { sensitivity: "base" }));
}

async function loadRecipes(opts) {
  opts = opts || {};
  const silent = !!opts.silent;

  // 1) try cache for a fast first paint (only relevant on first boot)
  if (!silent) {
    try {
      const cached = JSON.parse(sessionStorage.getItem(CONFIG.CACHE_KEY) || "null");
      if (cached && Date.now() - cached.ts < CONFIG.CACHE_TTL_MS) {
        const data = validateRecipesPayload(cached.data);
        state.recipes = sortRecipesByName(data.recipes);
        state.updatedAt =
          cached.data.sourceUpdatedAt || cached.data.fetchedAt || cached.data.updatedAt || null;
        state.dataSource = "cache";
        state.view = "list";
        render();
      }
    } catch (e) {
      console.warn("La caché de recetas no era válida.", e);
    }
  }

  if (silent) {
    state.refreshing = true;
    if (state.view === "list") render();
  }

  try {
    let result;
    try {
      result = await fetchRecipesFrom(
        CONFIG.DATA_URL || CONFIG.LOCAL_FALLBACK,
        CONFIG.DATA_URL ? "live" : "fallback"
      );
    } catch (liveError) {
      if (!CONFIG.DATA_URL) throw liveError;
      result = await fetchRecipesFrom(CONFIG.LOCAL_FALLBACK, "fallback");
    }
    const { data, source } = result;
    state.recipes = sortRecipesByName(data.recipes || []);
    state.updatedAt = data.sourceUpdatedAt || data.fetchedAt || data.updatedAt || null;
    state.dataSource = source;
    state.view = state.view === "loading" ? "list" : state.view;
    lastFetchAt = Date.now();
    try {
      sessionStorage.setItem(
        CONFIG.CACHE_KEY,
        JSON.stringify({ ts: Date.now(), data })
      );
    } catch (e) {
      console.warn("No se pudo actualizar la caché de recetas.", e);
    }
    state.refreshing = false;
    // Only force a re-render when it's safe: the list view, or first boot.
    // Detail/steps/final keep whatever was already on screen so an
    // in-progress recipe never changes under the cook's hands; the fresh
    // data is already in state.recipes for the next time they navigate.
    if (state.view === "list" || root.dataset.rendered !== "1") render();
  } catch (err) {
    state.refreshing = false;
    if (!state.recipes.length) {
      state.view = "error";
      state.errorMsg = err.message;
      render();
    } else if (state.view === "list") {
      render();
    }
  }
}

function refreshRecipes() {
  loadRecipes({ silent: true });
}

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState !== "visible") return;
  if (state.view === "loading") return;
  // Avoid refetching too often if the tab is switched back and forth.
  if (Date.now() - lastFetchAt < 30 * 1000) return;
  refreshRecipes();
});

/* ---------------- Helpers ---------------- */

function findRecipe(id) {
  return state.recipes.find((r) => r.id === id) || null;
}

function fmtQty(n) {
  if (n === null || n === undefined) return "";
  const rounded = n >= 10 ? Math.round(n) : Math.round(n * 10) / 10;
  return rounded.toString().replace(".", ",");
}

function titleCase(s) {
  if (!s) return "";
  return s.charAt(0) + s.slice(1).toLowerCase();
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
    }
  }, 1000);
}

function updateTimerRemaining(timer) {
  if (!timer.running || timer.done) return;
  timer.remaining = Math.max(0, Math.ceil((timer.endsAt - Date.now()) / 1000));
}

function startTimer(label, seconds, sessionId, recipeName) {
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
  });
  saveRuntimeState();
  ensureTimerTick();
  renderTimerBar();
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
};

window.cancelTimer = function (id) {
  state.timers = state.timers.filter((x) => x.id !== id);
  saveRuntimeState();
  renderTimerBar();
};

function renderTimerBar() {
  if (!timerBarEl) return;
  document.body.classList.toggle("has-timer", state.timers.length > 0);
  if (!state.timers.length) {
    timerBarEl.innerHTML = "";
    timerBarEl.classList.remove("visible");
    return;
  }
  timerBarEl.classList.add("visible");
  timerBarEl.innerHTML = state.timers
    .map((t) => {
      const sessionStillOpen = t.sessionId && state.sessions.some((s) => s.id === t.sessionId);
      const info = t.done
        ? t.recipeName
          ? `<span class="rtmy-timer-clock">⏰</span><span class="rtmy-timer-label"><b>${escapeHtml(titleCase(t.recipeName))}</b> necesita tu atención — ${escapeHtml(t.label)}</span>`
          : `<span class="rtmy-timer-clock">¡Listo!</span><span class="rtmy-timer-label">${escapeHtml(t.label)}</span>`
        : `<span class="rtmy-timer-clock">${fmtClock(t.remaining)}</span><span class="rtmy-timer-label">${escapeHtml(t.label)}${t.recipeName ? " · " + escapeHtml(titleCase(t.recipeName)) : ""}</span>`;
      return `
    <div class="rtmy-timer-row ${t.done ? "done" : ""}">
      <div class="rtmy-timer-info">${info}</div>
      <div class="rtmy-timer-controls">
        ${t.done && sessionStillOpen ? `<button class="rtmy-timer-goto" onclick="switchToSession('${t.sessionId}')">Ir a la receta</button>` : ""}
        ${
          !t.done
            ? `<button onclick="toggleTimer(${t.id})" aria-label="${t.running ? "Pausar" : "Reanudar"}">${t.running ? "⏸" : "▶️"}</button>`
            : ""
        }
        <button onclick="cancelTimer(${t.id})" aria-label="Descartar">✕</button>
      </div>
    </div>
  `;
    })
    .join("");
}

/* ---------------- Ingredient quantities in step text ----------------
   "saltear la cebolla" -> "saltear los 210 g de cebolla", pulling the
   number straight from the (already scaled) ingredient list — never a
   guess, always the real figure for that batch.

   Matching strategy: use each ingredient's first word as its "key" (the
   table often reads "Cebolla blanca" while the step just says "la
   cebolla"). If two ingredients in the same list share a first word
   (e.g. two "Aceite ..."), that key is skipped entirely rather than
   guessing which one a step means. Only the first mention of each
   ingredient in a step gets annotated.
------------------------------------------------- */

const UNIT_GENDER = {
  g: "m", gramo: "m", gramos: "m",
  kg: "m", kilo: "m", kilos: "m",
  ml: "m", mililitro: "m", mililitros: "m",
  cl: "m", l: "m", litro: "m", litros: "m",
  ud: "f", "u.": "f", unidad: "f", unidades: "f",
  loncha: "f", lonchas: "f",
};

function articleFor(unit, plural) {
  const gender = UNIT_GENDER[(unit || "").toLowerCase()] || "m";
  if (gender === "f") return plural ? "las" : "la";
  return plural ? "los" : "el";
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function ingredientKeyTerm(name) {
  const word = name.trim().split(/\s+/)[0];
  return word.toLowerCase();
}

// Generic culinary nouns that regularly show up in prose meaning "the
// mix/prep I already made" rather than the specific named ingredient
// (e.g. "Base de tortilla RTMY" the ingredient vs. "la base de patata"
// meaning the potato mixture). Too ambiguous to auto-annotate safely.
const GENERIC_KEY_BLOCKLIST = new Set([
  "base", "salsa", "mezcla", "masa", "relleno", "caldo", "fondo", "reduccion", "reducción",
]);

function injectIngredientQuantities(text, ingredients) {
  if (!ingredients || !ingredients.length) return text;

  const byKey = new Map();
  ingredients.forEach((ing) => {
    const key = ingredientKeyTerm(ing.name);
    if (key.length < 3 || GENERIC_KEY_BLOCKLIST.has(key)) return;
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(ing);
  });

  let result = text;
  byKey.forEach((list, key) => {
    if (list.length > 1) return; // ambiguous shared first word — don't guess
    const ing = list[0];
    // \b on both ends: without the trailing boundary, a short key like
    // "sal" would also match inside an unrelated word like "saltear".
    // The optional (es|s) only covers simple Spanish pluralization. The
    // lookbehind skips mentions that already carry their own explicit
    // quantity (e.g. "embolsar en raciones de 150 gramos de carrilleras")
    // so we don't tack the full batch weight onto a portion size.
    // Trailing \s* (not \s+) on the lookbehind: the regex engine tries a
    // match starting right at the space before the key as well as right
    // after it, so the exclusion has to fire at either boundary or the
    // "at the space" attempt slips through as the leftmost match.
    const re = new RegExp(
      `(?<!\\d+\\s*(?:g|gramos?|kg|kilos?|ml|mililitros?|ud|unidades?)\\s+de\\s*)(el|la|los|las)?(\\s*)\\b(${escapeRegex(key)}(?:es|s)?)\\b`,
      "i"
    );
    const m = result.match(re);
    if (!m) return;
    const qtyLabel = `${fmtQty(ing.qty)} ${ing.unit}`;
    const plural = Math.abs(ing.qty - 1) > 0.001;
    const matchedWord = m[3];
    // No article to swap out (e.g. "con sal", "por harina"): the leading
    // whitespace (m[2]) was consumed by the match too, so it must be put
    // back or the words on either side run together ("consal").
    const replacement = m[1]
      ? `${articleFor(ing.unit, plural)} ${qtyLabel} de ${matchedWord}`
      : `${m[2]}${matchedWord} (${qtyLabel})`;
    result = result.replace(m[0], replacement);
  });
  return result;
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
    return;
  }
  const stage0 = recipe.stages[0];
  const main = stage0 && stage0.ingredients[0];
  state.scaleQty = main ? main.qty : null;
  state.scaleFactor = 1;
}

function mainIngredient(recipe) {
  const stage0 = recipe.stages[0];
  return stage0 && stage0.ingredients.length ? stage0.ingredients[0] : null;
}

function applyScale(ingredients, ratio) {
  return ingredients.map((ing) => ({ ...ing, qty: ing.qty * ratio }));
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
    session = { id: recipe.id, recipe, scaleFactor: 1, scaleQty: null, startStageIndex: 0, flow: [], flowPos: 0 };
    state.sessions.push(session);
  }
  session.recipe = recipe;
  session.scaleFactor = state.scaleFactor;
  session.scaleQty = state.scaleQty;
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
  if (state.view === "loading") return renderLoading();
  if (state.view === "error") return renderError();
  if (state.view === "list") return renderList();
  if (state.view === "detail") return renderDetail();
  if (state.view === "steps") return renderSteps();
  if (state.view === "final") return renderFinal();
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

function fmtUpdatedAt() {
  if (!state.updatedAt) return "";
  try {
    const d = new Date(state.updatedAt);
    return d.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" });
  } catch (e) {
    return "";
  }
}

function buildListItemsHtml() {
  const q = normalizeSearchText(state.search.trim());
  const list = state.recipes.filter((r) => normalizeSearchText(r.name).includes(q));
  return list.length
    ? list
        .map((r) => {
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
        })
        .join("")
    : `<div class="rtmy-empty">No se encontraron recetas.</div>`;
}

function renderList() {
  const sourceLabel =
    state.dataSource === "fallback"
      ? " · copia local"
      : state.dataSource === "cache"
        ? " · caché"
        : "";
  root.innerHTML = `
    ${header(false)}
    <div class="rtmy-refresh-row">
      <span>${state.refreshing ? "Actualizando…" : state.updatedAt ? "Datos consultados a las " + fmtUpdatedAt() + sourceLabel : sourceLabel.replace(" · ", "")}</span>
      <button class="rtmy-refresh-btn" onclick="refreshRecipes()" ${state.refreshing ? "disabled" : ""}>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="${state.refreshing ? "spin" : ""}"><path d="M4 4v5h5M20 20v-5h-5"/><path d="M4 9a8 8 0 0 1 14-4.5M20 15a8 8 0 0 1-14 4.5"/></svg>
        Actualizar
      </button>
    </div>
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

      <ul class="rtmy-ingredients">
        ${scaledIngredients(r)
          .map(
            (i) => `
          <li><span class="name">${escapeHtml(titleCase(i.name))}</span><span class="qty">${fmtQty(i.qty)} ${escapeHtml(i.unit)}</span></li>
        `
          )
          .join("")}
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
              <div class="stage-sub">${(s.steps && s.steps.length) || 0} paso${(s.steps || []).length === 1 ? "" : "s"}</div>
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
  const rawStepText =
    (stage.steps && stage.steps.length ? stage.steps : [stage.description])[
      item.stepIndexInStage
    ];
  const warnings = extractWarnings(rawStepText);
  const durations = extractDurations(rawStepText);
  state.currentStepDurations = durations;

  // Ingredients available to name-check quantities against: the scaled
  // base batch for stage 0, or that stage's own (unscaled) table for
  // composite recipes with sub-recipes further down.
  const stageIngredients =
    item.stageIndex === 0
      ? applyScale(r.stages[0].ingredients, session.scaleFactor)
      : stage.ingredients;
  const stepText = injectIngredientQuantities(rawStepText, stageIngredients);

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
      item.stepIndexInStage === 0 && item.stageIndex !== 0 && stage.ingredients.length
        ? `
    <div class="rtmy-panel" style="padding:14px 16px;">
      <div class="rtmy-scale-label" style="margin-bottom:2px;">Ingredientes de esta etapa</div>
      <ul class="rtmy-ingredients">
        ${stage.ingredients
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
  state.finishedSession = { recipe: s.recipe };
  state.sessions = state.sessions.filter((x) => x.id !== s.id);
  if (state.activeSessionId === s.id) state.activeSessionId = null;
  saveRuntimeState();
  goto("#/receta/" + encodeURIComponent(s.recipe.id) + "/final");
};

function renderFinal() {
  const r = state.finishedSession.recipe;
  const lastStage = r.stages[r.stages.length - 1];
  const steps = lastStage.steps && lastStage.steps.length ? lastStage.steps : [];

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

/* ---------------- Boot ---------------- */

window.goto = goto;
window.startFlow = startFlow;
window.loadRecipes = loadRecipes;
window.refreshRecipes = refreshRecipes;

window.addEventListener("hashchange", onHashChange);
loadRecipes().then(() => {
  restoreRuntimeState();
  if (state.timers.some((timer) => timer.running && !timer.done)) ensureTimerTick();
  renderTimerBar();
  renderSessionBar();
  onHashChange();
});
