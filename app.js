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
  CACHE_KEY: "rtmy_recipes_cache_v1",
  CACHE_TTL_MS: 1000 * 60 * 10, // 10 minutos
};

const state = {
  view: "loading", // loading | error | list | detail | steps | final
  recipes: [],
  updatedAt: null,
  currentRecipe: null,
  search: "",
  scaleFactor: 1,
  scaleQty: null,
  startStageIndex: 0,
  flow: [],
  flowPos: 0,
  refreshing: false,
};

const root = document.getElementById("app");
let lastFetchAt = 0;

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

async function loadRecipes(opts) {
  opts = opts || {};
  const silent = !!opts.silent;

  // 1) try cache for a fast first paint (only relevant on first boot)
  if (!silent) {
    try {
      const cached = JSON.parse(sessionStorage.getItem(CONFIG.CACHE_KEY) || "null");
      if (cached && Date.now() - cached.ts < CONFIG.CACHE_TTL_MS) {
        state.recipes = cached.data.recipes;
        state.updatedAt = cached.data.updatedAt || null;
        state.view = "list";
        render();
      }
    } catch (e) {}
  }

  if (silent) {
    state.refreshing = true;
    if (state.view === "list") render();
  }

  const url = CONFIG.DATA_URL || CONFIG.LOCAL_FALLBACK;
  try {
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error("HTTP " + res.status);
    const data = await res.json();
    state.recipes = data.recipes || [];
    state.updatedAt = data.updatedAt || null;
    state.view = state.view === "loading" ? "list" : state.view;
    lastFetchAt = Date.now();
    try {
      sessionStorage.setItem(
        CONFIG.CACHE_KEY,
        JSON.stringify({ ts: Date.now(), data })
      );
    } catch (e) {}
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

function slugParam() {
  const m = location.hash.match(/^#\/receta\/([^/]+)/);
  return m ? decodeURIComponent(m[1]) : null;
}

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

/* ---------------- Router ---------------- */

function onHashChange() {
  const hash = location.hash || "#/";
  if (hash === "#/" || hash === "") {
    state.view = "list";
  } else if (hash.startsWith("#/receta/")) {
    const id = slugParam();
    const r = findRecipe(id);
    if (r) {
      state.currentRecipe = r;
      state.view = "detail";
      resetScale(r);
    } else {
      state.view = "list";
    }
  }
  render();
  window.scrollTo(0, 0);
}

function goto(hash) {
  location.hash = hash;
}

/* ---------------- Scaling ---------------- */

function resetScale(recipe) {
  const stage0 = recipe.stages[0];
  const main = stage0 && stage0.ingredients[0];
  state.scaleQty = main ? main.qty : null;
  state.scaleFactor = 1;
  state.startStageIndex = 0;
}

function mainIngredient(recipe) {
  const stage0 = recipe.stages[0];
  return stage0 && stage0.ingredients.length ? stage0.ingredients[0] : null;
}

function scaledIngredients(recipe) {
  const stage0 = recipe.stages[0];
  if (!stage0) return [];
  const ratio = state.scaleFactor;
  return stage0.ingredients.map((ing) => ({
    ...ing,
    qty: ing.qty * ratio,
  }));
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
  state.startStageIndex = startStageIndex;
  state.flow = buildFlow(recipe, startStageIndex);
  state.flowPos = 0;
  state.view = "steps";
  render();
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

function renderList() {
  const q = state.search.trim().toLowerCase();
  const list = state.recipes.filter((r) => r.name.toLowerCase().includes(q));
  root.innerHTML = `
    ${header(false)}
    <div class="rtmy-refresh-row">
      <span>${state.refreshing ? "Actualizando…" : state.updatedAt ? "Datos de las " + fmtUpdatedAt() : ""}</span>
      <button class="rtmy-refresh-btn" onclick="refreshRecipes()" ${state.refreshing ? "disabled" : ""}>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="${state.refreshing ? "spin" : ""}"><path d="M4 4v5h5M20 20v-5h-5"/><path d="M4 9a8 8 0 0 1 14-4.5M20 15a8 8 0 0 1-14 4.5"/></svg>
        Actualizar
      </button>
    </div>
    <input class="rtmy-search" placeholder="Buscar receta…" value="${escapeHtml(
      state.search
    )}" oninput="onSearch(this.value)">
    <div class="rtmy-list">
      ${
        list.length
          ? list
              .map(
                (r) => `
        <button class="rtmy-card" onclick="goto('#/receta/${encodeURIComponent(r.id)}')">
          <div>
            <div class="rtmy-card-name">${escapeHtml(titleCase(r.name))}</div>
            <div class="rtmy-card-meta">${r.yieldAmount ?? ""} ${escapeHtml(
                  r.yieldUnit || ""
                )} · ${r.stages.length} etapa${r.stages.length === 1 ? "" : "s"}</div>
          </div>
          <svg class="rtmy-card-arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 6l6 6-6 6"/></svg>
        </button>
      `
              )
              .join("")
          : `<div class="rtmy-empty">No se encontraron recetas.</div>`
      }
    </div>
    <p class="rtmy-footer-note">Rice to Meet You · El Raval, Barcelona</p>
  `;
}

window.onSearch = function (val) {
  state.search = val;
  renderList();
};

function renderDetail() {
  const r = state.currentRecipe;
  const stage0 = r.stages[0];
  const main = mainIngredient(r);
  const hasMultipleStages = r.stages.length > 1;

  root.innerHTML = `
    ${header(true, "#/", "Detalle de receta")}
    <div class="rtmy-panel">
      <div class="rtmy-recipe-yield">${r.yieldAmount ?? ""} ${escapeHtml(r.yieldUnit || "")}</div>
      <h2 class="rtmy-recipe-name">${escapeHtml(titleCase(r.name))}</h2>
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
      hasMultipleStages
        ? `
    <div class="rtmy-panel">
      <div class="rtmy-scale-label">¿Por dónde empezamos?</div>
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
        : `<button class="rtmy-cta" onclick="startFlow(state.currentRecipe, 0)">Empezar receta</button>`
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
  const r = state.currentRecipe;
  const flow = state.flow;
  const pos = state.flowPos;
  const item = flow[pos];
  const stage = r.stages[item.stageIndex];
  const stepText =
    (stage.steps && stage.steps.length ? stage.steps : [stage.description])[
      item.stepIndexInStage
    ];
  const warnings = extractWarnings(stepText);

  // progress dots: one per stage in active range
  const activeStages = r.stages.slice(state.startStageIndex);
  const dots = activeStages
    .map((s, idx) => {
      const stageIdx = state.startStageIndex + idx;
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
  if (state.flowPos < state.flow.length - 1) {
    state.flowPos++;
    render();
    window.scrollTo(0, 0);
  }
};
window.stepBack = function () {
  if (state.flowPos > 0) {
    state.flowPos--;
    render();
    window.scrollTo(0, 0);
  }
};
window.finishFlow = function () {
  state.view = "final";
  render();
  window.scrollTo(0, 0);
};

function renderFinal() {
  const r = state.currentRecipe;
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

/* ---------------- Boot ---------------- */

window.goto = goto;
window.startFlow = startFlow;
window.loadRecipes = loadRecipes;
window.refreshRecipes = refreshRecipes;

window.addEventListener("hashchange", onHashChange);
loadRecipes().then(() => {
  onHashChange();
});
