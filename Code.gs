/**
 * RTMY — API de recetas para la app de preparaciones del taller.
 *
 * Lee en vivo el Excel "MANUAL COCINA" y devuelve JSON con las recetas
 * que siguen la plantilla "FICHA TÉCNICA · ESCANDALLO DE COCINA".
 * Nunca expone columnas de coste, merma, subtotal ni precio de compra:
 * solo cantidad, unidad, nombre de ingrediente y pasos de preparación.
 *
 * Despliegue: Extensiones > Apps Script > pegar este código > Desplegar
 * > Nueva implementación > Aplicación web > Ejecutar como: Yo
 * > Quién tiene acceso: Cualquier usuario > Desplegar.
 * Copiar la URL que termina en /exec y pegarla en CONFIG.DATA_URL de app.js.
 */

var SPREADSHEET_ID = '1wcm_-ASHlkXjzOTXGyIl2lcbOfi1GFWDBZHRleVPPwo';
var MARKER = 'FICHA TÉCNICA · ESCANDALLO DE COCINA';

function doGet(e) {
  var recipes = buildRecipes();
  var payload = {
    updatedAt: new Date().toISOString(),
    recipes: recipes,
  };
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildRecipes() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var recipes = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName().trim();
    if (sheetName.toUpperCase() === 'PLANTILLA') continue;

    var marker = safeGet(sheet, 1, 5); // E1
    if (!marker || String(marker).toUpperCase().indexOf(MARKER.toUpperCase()) === -1) {
      continue;
    }

    var recipe = parseSheet(sheet);
    if (recipe) {
      recipe.sourceSheet = sheetName;
      recipes.push(recipe);
    }
  }
  return recipes;
}

function safeGet(sheet, row, col) {
  if (row > sheet.getMaxRows() || col > sheet.getMaxColumns()) return null;
  var v = sheet.getRange(row, col).getValue();
  return v === '' ? null : v;
}

function isNumber(v) {
  return typeof v === 'number' && !isNaN(v);
}

function cleanText(v) {
  if (v === null || v === undefined) return null;
  var s = String(v).trim();
  return s === '' ? null : s;
}

function looksNumbered(text) {
  return /^\s*\d+\.\s/.test(text);
}

function splitSteps(text) {
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  var parts = text.split(/\n(?=\s*\d+\.\s?)/);
  var steps = [];
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i].trim();
    if (!p) continue;
    p = p.replace(/^\d+\.\s*/, '').trim();
    if (p) steps.push(p);
  }
  if (steps.length === 0 && text.trim()) steps.push(text.trim());
  return steps;
}

function slugify(s) {
  var out = s
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .trim()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
  return out;
}

function parseSheet(sheet) {
  var name = cleanText(safeGet(sheet, 2, 5)); // E2
  if (!name) return null;
  var yieldAmountRaw = safeGet(sheet, 2, 6); // F2
  var yieldAmount = isNumber(yieldAmountRaw) ? yieldAmountRaw : null;
  var yieldUnit = cleanText(safeGet(sheet, 3, 6)); // F3

  var lastRow = sheet.getLastRow();
  var stages = [];
  var current = null;

  for (var r = 5; r <= lastRow; r++) {
    var bVal = cleanText(safeGet(sheet, r, 2)); // B
    var cVal = safeGet(sheet, r, 3); // C
    var dVal = cleanText(safeGet(sheet, r, 4)); // D
    var eVal = cleanText(safeGet(sheet, r, 5)); // E

    if (bVal === 'CANTIDAD') continue;

    if (isNumber(cVal) && eVal) {
      if (!current) current = newStage(name);
      current.ingredients.push({ qty: cVal, unit: dVal || '', name: eVal });
      continue;
    }

    if (typeof cVal === 'string' && cVal.trim() && !dVal && !eVal) {
      var text = cVal.trim();
      if (looksNumbered(text) || text.indexOf('\n') !== -1) {
        if (!current) current = newStage(name);
        current.steps = current.steps.concat(splitSteps(text));
      } else {
        if (current && !current.ingredients.length && !current.steps.length && current.description === null) {
          current.description = text;
        } else {
          if (current) stages.push(current);
          current = newStage(text);
        }
      }
      continue;
    }
  }
  if (current) stages.push(current);

  stages = stages.filter(function (s) {
    return s.ingredients.length || s.steps.length || s.description;
  });
  if (!stages.length) return null;

  // Merge the ingredient-table stage with the immediately following
  // steps-only stage: the template splits the first cooking phase across
  // two labeled blocks ("RECETA BASE" + "PROCEDIMIENTO...").
  if (
    stages.length >= 2 &&
    stages[0].ingredients.length && !stages[0].steps.length &&
    !stages[1].ingredients.length && stages[1].steps.length
  ) {
    var merged = {
      title: stages[0].title,
      procedureTitle: stages[1].title,
      description: stages[0].description,
      ingredients: stages[0].ingredients,
      steps: stages[1].steps,
    };
    stages = [merged].concat(stages.slice(2));
  } else {
    stages[0].procedureTitle = null;
    for (var k = 1; k < stages.length; k++) {
      if (stages[k].procedureTitle === undefined) stages[k].procedureTitle = null;
    }
  }

  return {
    id: slugify(name),
    name: name,
    yieldAmount: yieldAmount,
    yieldUnit: yieldUnit,
    stages: stages,
  };
}

function newStage(title) {
  return { title: title, description: null, ingredients: [], steps: [] };
}

/**
 * Ejecutar manualmente desde el editor de Apps Script para comprobar
 * que el parseo funciona antes de desplegar (ver "Registro de ejecución").
 */
function testParse() {
  var recipes = buildRecipes();
  Logger.log('Recetas encontradas: ' + recipes.length);
  recipes.forEach(function (r) {
    Logger.log(r.name + ' -> ' + r.stages.length + ' etapas');
  });
}
