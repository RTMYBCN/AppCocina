const assert = require("assert").strict;
const fs = require("fs");
const path = require("path");
const root = path.resolve(__dirname, "..");

function test(name, run) {
  try {
    run();
    console.log(`✓ ${name}`);
  } catch (error) {
    console.error(`✗ ${name}`);
    throw error;
  }
}

test("recipes.json has valid unique recipes and executable stages", () => {
  const payload = JSON.parse(fs.readFileSync(path.join(root, "recipes.json"), "utf8"));
  assert.ok(Array.isArray(payload.recipes));
  assert.ok(payload.recipes.length > 0);

  const ids = new Set();
  payload.recipes.forEach((recipe) => {
    assert.equal(typeof recipe.id, "string");
    assert.ok(recipe.id.length > 0);
    assert.ok(!ids.has(recipe.id), `duplicate recipe id: ${recipe.id}`);
    ids.add(recipe.id);
    assert.equal(typeof recipe.name, "string");
    assert.ok(Array.isArray(recipe.stages) && recipe.stages.length > 0);

    recipe.stages.forEach((stage) => {
      assert.ok(Array.isArray(stage.ingredients));
      assert.ok(Array.isArray(stage.steps));
      assert.ok(stage.ingredients.length || stage.steps.length || stage.description);
      stage.ingredients.forEach((ingredient) => {
        assert.equal(typeof ingredient.name, "string");
        assert.ok(Number.isFinite(ingredient.qty));
        assert.ok(ingredient.qty >= 0);
      });
      stage.steps.forEach((step) => assert.ok(typeof step === "string" && step.trim()));
    });
  });
});

test("HTML references the current cache-busted assets and accessibility regions", () => {
  const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
  assert.match(html, /styles\.css\?v=18/);
  assert.match(html, /planner\.js\?v=18/);
  assert.match(html, /app\.js\?v=18/);
  assert.match(html, /text\.js\?v=18/);
  assert.match(html, /id="timer-bar"[^>]+aria-live="assertive"/);
});

test("Apps Script reads sheets in bulk and exposes distinct timestamps", () => {
  const source = fs.readFileSync(path.join(root, "Code.gs"), "utf8");
  assert.match(source, /getDataRange\(\)\.getValues\(\)/);
  assert.doesNotMatch(source, /getRange\(row,\s*col\)\.getValue\(\)/);
  assert.match(source, /fetchedAt:/);
  assert.match(source, /sourceUpdatedAt:/);
  assert.match(source, /CacheService\.getScriptCache\(\)/);
});

test("frontend contains resilient fallback, persisted sessions and deadline timers", () => {
  const source = fs.readFileSync(path.join(root, "app.js"), "utf8");
  assert.match(source, /fetchRecipesFrom\(CONFIG\.LOCAL_FALLBACK,\s*"fallback"\)/);
  assert.match(source, /localStorage\.setItem/);
  assert.match(source, /endsAt:\s*Date\.now\(\)\s*\+\s*seconds\s*\*\s*1000/);
  assert.match(source, /validateRecipesPayload/);
});

test("active preparations are highlighted and can start a new preparation", () => {
  const source = fs.readFileSync(path.join(root, "app.js"), "utf8");
  const css = fs.readFileSync(path.join(root, "styles.css"), "utf8");
  assert.match(source, /Nueva preparación/);
  assert.match(source, /En preparación/);
  assert.match(source, /rtmy-card \$\{inProgress \? "in-progress"/);
  assert.match(css, /\.rtmy-card\.in-progress/);
  assert.match(css, /\.rtmy-new-prep-btn/);
});

test("the planner shows one column per cook and names the recipe on each card", () => {
  const source = fs.readFileSync(path.join(root, "app.js"), "utf8");
  const css = fs.readFileSync(path.join(root, "styles.css"), "utf8");
  assert.match(source, /rtmy-plan-cols/);
  assert.match(source, /class="rtmy-plan-col"/);
  assert.match(source, /rtmy-plan-recipe-tag/);
  // los pasos de máquina van en la franja de quien la pone en marcha
  assert.match(source, /t\.watcher === p/);
  assert.match(source, /en marcha · no hace falta estar/);
  assert.match(css, /\.rtmy-plan-cols\{[^}]*overflow-x:auto/s);
  assert.match(css, /\.rtmy-plan-col-head\{[^}]*position:sticky/s);
  // "SALMORRETA RTMY" debe quedar como "Salmorreta RTMY", no "Salmorreta Rtmy"
  assert.match(source, /const ACRONYMS = new Set\(\["RTMY"/);
});

test("each timer belongs to the cook who started it", () => {
  const source = fs.readFileSync(path.join(root, "app.js"), "utf8");
  const css = fs.readFileSync(path.join(root, "styles.css"), "utf8");
  // el temporizador guarda la persona y se pinta en su franja
  assert.match(source, /person: Number\.isInteger\(person\) \? person : null/);
  assert.match(source, /function renderPlanTimers\(\)/);
  assert.match(source, /rtmy-plan-timers\[data-person\]/);
  assert.match(source, /rtmy-timer-who/); // la barra de abajo marca P1 / P2
  assert.match(css, /\.rtmy-plan-timer-chip/);
});

test("an open preparation follows the recipe when the sheet changes", () => {
  const source = fs.readFileSync(path.join(root, "app.js"), "utf8");
  // la copia congelada de la ficha que guarda cada sesión se vuelve a enganchar
  assert.match(source, /function relinkRecipeRefs\(\)/);
  assert.match(source, /session\.recipe = fresh;/);
  assert.match(source, /session\.recipeChanged = true;/);
  // también la ficha que está en pantalla y la sesión terminada
  assert.match(source, /state\.currentRecipe = fresh;/);
  // y se reengancha al restaurar, que ocurre DESPUÉS de cargar las recetas
  assert.match(source, /restoreRuntimeState\(\);[\s\S]{0,240}relinkRecipeRefs\(\);/);
  // "Actualizar" repinta también la ficha abierta
  assert.match(source, /state\.view === "list" \|\| manual \|\|/);
  assert.match(source, /La ficha se ha actualizado/);
});

test("the bottom bar splits into one lane per cook while the plan is open", () => {
  const source = fs.readFileSync(path.join(root, "app.js"), "utf8");
  const css = fs.readFileSync(path.join(root, "styles.css"), "utf8");
  assert.match(source, /function planBarPersons\(\)/);
  assert.match(source, /rtmy-timer-grid/);
  assert.match(source, /rtmy-timer-lane-head/);
  // los temporizadores sin dueño siguen a lo ancho
  assert.match(source, /const loose = state\.timers\.filter/);
  assert.match(css, /\.rtmy-timer-grid\{/);
  assert.match(css, /grid-template-columns:repeat\(var\(--cols,2\)/);
});

test("each stage declares which ingredient phases it needs", () => {
  const code = fs.readFileSync(path.join(root, "Code.gs"), "utf8");
  // la celda de al lado del título de la etapa (columna D del getValues = col. C)
  assert.match(code, /function splitPhases\(text\)/);
  assert.match(code, /phases: splitPhases\(cleanText\(valueAt\(values, r, 3\)\)\)/);
  assert.match(code, /phases: b\.phases \|\| \[\]/);

  const source = fs.readFileSync(path.join(root, "app.js"), "utf8");
  const css = fs.readFileSync(path.join(root, "styles.css"), "utf8");
  // la tabla se agrupa por fases y el índice que escala sigue siendo el real
  assert.match(source, /function renderIngredientPicker\(recipe\)/);
  assert.match(source, /onSelectMainIngredient\(\$\{index\}\)/);
  assert.match(css, /\.rtmy-ingredients li\.rtmy-ingredient-group\{/);
  // el selector de etapa dice cuántos ingredientes hay que alistar
  assert.match(source, /function stageIngredientNote\(recipe, stageIndex\)/);
  // los pasos de las etapas 2ª y siguientes ya reciben su tabla escalada
  assert.match(source, /RTMYText\.ingredientsForStage\(r, item\.stageIndex\)/);
  assert.match(source, /usesPhases \|\| isMainStage \? session\.scaleFactor : 1/);
  assert.match(source, /applyScale\(stageIngredients, stageRatio\)/);
});
