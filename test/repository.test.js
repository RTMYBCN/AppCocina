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
  assert.match(html, /styles\.css\?v=5/);
  assert.match(html, /app\.js\?v=5/);
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
