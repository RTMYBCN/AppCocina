/**
 * RTMY — API de recetas para la app de preparaciones del taller.
 *
 * Lee en vivo el Excel "MANUAL COCINA" (versión reorganizada, fichas B/S/C)
 * y devuelve JSON con las recetas.
 *
 * Fichas que se publican: las pestañas cuyo nombre empieza por un código
 * de ficha — B01 (preparaciones base), S01 (sub-preparaciones) o
 * C01 (platos de carta). El resto de pestañas (ÍNDICE, PRECIOS,
 * INCIDENCIAS, PLANTILLA, ARROCES, BEBIDAS, facturas…) se ignoran.
 *
 * Nunca se expone información interna: solo se leen el nombre, la
 * descripción, el rendimiento, los ingredientes (columna C = nombre,
 * E = cantidad, F = unidad) y los bloques de pasos. Las columnas de
 * producto comercial, merma, €/unidad, coste y aviso (G–K), las filas de
 * coste/PVP y las "NOTAS DE LA REVISIÓN" no se leen ni se devuelven.
 *
 * Despliegue: Extensiones > Apps Script > pegar este código > Desplegar
 * > Nueva implementación > Aplicación web > Ejecutar como: Yo
 * > Quién tiene acceso: Cualquier usuario > Desplegar.
 * Copiar la URL que termina en /exec y pegarla en CONFIG.DATA_URL de app.js.
 */

var SPREADSHEET_ID = '1OhUs4RJdkgbZuFM0DcmgqDPI4QxQP7dP-6jTtLjQtbY';
var CACHE_KEY = 'recipes-v3';

// Código de la pestaña (B01, S02, C09…) -> sección del listado.
var SECTIONS = {
  B: { name: 'Preparaciones base', order: 1 },
  S: { name: 'Sub-preparaciones', order: 2 },
  C: { name: 'Platos de carta', order: 3 },
};
var TAB_RE = /^([BSC])(\d{2})(?:\s|$)/i;

function doGet(e) {
  // ?fresh=… (botón "Actualizar" de la app): ignora la caché de 5 min, relee
  // el Sheet y deja la caché al día para los demás.
  var fresh = !!(e && e.parameter && e.parameter.fresh);
  var cache = CacheService.getScriptCache();
  var cached = fresh ? null : cache.get(CACHE_KEY);
  var recipes = cached ? JSON.parse(cached) : buildRecipes();
  if (!cached) {
    try {
      cache.put(CACHE_KEY, JSON.stringify(recipes), 300);
    } catch (err) {
      // JSON demasiado grande para la caché (límite 100 KB): se sirve sin caché.
    }
  }
  var payload = {
    fetchedAt: new Date().toISOString(),
    sourceUpdatedAt: DriveApp.getFileById(SPREADSHEET_ID).getLastUpdated().toISOString(),
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
  var usedIds = {};

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName().trim();
    var tab = TAB_RE.exec(sheetName);
    if (!tab) continue;

    var values = sheet.getDataRange().getValues();
    var recipe = parseSheet(sheetName, tab[1].toUpperCase(), tab[1].toUpperCase() + tab[2], values);
    if (recipe) {
      var baseId = recipe.id || 'receta';
      var id = baseId;
      var suffix = 2;
      while (usedIds[id]) id = baseId + '-' + suffix++;
      recipe.id = id;
      usedIds[id] = true;
      recipe.sourceSheet = sheetName;
      recipes.push(recipe);
    }
  }

  // Orden estable: sección y, dentro de ella, número de ficha.
  recipes.sort(function (a, b) {
    if (a.sectionOrder !== b.sectionOrder) return a.sectionOrder - b.sectionOrder;
    return a.code < b.code ? -1 : a.code > b.code ? 1 : 0;
  });
  return recipes;
}

function valueAt(values, row, col) {
  if (row < 1 || col < 1 || row > values.length || col > (values[row - 1] || []).length) return null;
  var v = values[row - 1][col - 1];
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

// "ración" -> "ración"/"raciones" según la cantidad; el resto se deja tal cual.
function normalizeYieldUnit(unit, amount) {
  var u = cleanText(unit);
  if (!u) return null;
  if (/^raci[oó]n(es)?$/i.test(u)) return amount === 1 ? 'ración' : 'raciones';
  return u;
}

// Fila (1-based) de la primera celda de la columna `col` cuyo texto es `label`.
function findLabelRow(values, col, label, fromRow, toRow) {
  var target = label.toLowerCase();
  for (var r = fromRow; r <= toRow; r++) {
    var v = cleanText(valueAt(values, r, col));
    if (v && v.toLowerCase() === target) return r;
  }
  return null;
}

function parseSheet(sheetName, sectionKey, code, values) {
  var name = cleanText(valueAt(values, 2, 2)); // B2
  if (!name) return null;
  var section = SECTIONS[sectionKey];
  var lastRow = values.length;

  var descRow = findLabelRow(values, 2, 'Descripción', 1, Math.min(lastRow, 12));
  var description = descRow ? cleanText(valueAt(values, descRow, 3)) : null; // C
  var yieldRow = findLabelRow(values, 2, 'Rendimiento', 1, Math.min(lastRow, 12));
  var yieldAmountRaw = yieldRow ? valueAt(values, yieldRow, 5) : null; // E
  var yieldAmount = isNumber(yieldAmountRaw) ? yieldAmountRaw : null;
  var yieldUnit = normalizeYieldUnit(yieldRow ? valueAt(values, yieldRow, 6) : null, yieldAmount); // F

  // Tabla de ingredientes: empieza justo debajo de la cabecera "Ingrediente" (col. C).
  var headerRow = findLabelRow(values, 3, 'Ingrediente', 1, Math.min(lastRow, 15));
  var ingredients = [];
  var r = (headerRow || 7) + 1;
  for (; r <= lastRow; r++) {
    var gVal = cleanText(valueAt(values, r, 7));
    if (gVal && /^COSTE TOTAL/i.test(gVal)) break; // fin de la tabla (a partir de aquí, costes)
    var cVal = cleanText(valueAt(values, r, 3));
    var eVal = valueAt(values, r, 5);
    if (cVal && isNumber(eVal)) {
      var ing = { qty: eVal, unit: cleanText(valueAt(values, r, 6)) || '', name: cVal };
      var phase = cleanText(valueAt(values, r, 2)); // B = fase (Masa, Relleno…)
      if (phase) ing.phase = phase;
      ingredients.push(ing);
    }
  }

  // Bloques de pasos: columna B, tras la tabla de costes. Un bloque es un
  // título corto seguido de una celda con los pasos numerados.
  var blocks = [];
  var current = null;
  var pending = null;
  for (r = r + 1; r <= lastRow; r++) {
    var bVal = cleanText(valueAt(values, r, 2));
    if (!bVal) continue;
    if (/^NOTAS DE LA REVISI/i.test(bVal)) break; // notas internas: no se publican
    if (/^PENDIENTE/i.test(bVal)) {
      pending = bVal;
      continue;
    }
    var isStepText = looksNumbered(bVal) || bVal.indexOf('\n') !== -1 || bVal.length > 90;
    if (isStepText) {
      if (!current) {
        current = { title: null, steps: [] };
        blocks.push(current);
      }
      current.steps = current.steps.concat(splitSteps(bVal));
    } else {
      current = { title: bVal, steps: [] };
      blocks.push(current);
    }
  }
  blocks = blocks.filter(function (b) { return b.steps.length; });

  var stages = [];
  if (blocks.length) {
    blocks.forEach(function (b, idx) {
      stages.push({
        title: b.title || (idx === 0 ? 'Preparación' : 'Paso ' + (idx + 1)),
        procedureTitle: null,
        description: idx === 0 ? description : null,
        // La tabla de ingredientes completa va en la primera etapa (es la que se escala).
        ingredients: idx === 0 ? ingredients : [],
        steps: b.steps,
      });
    });
  } else if (ingredients.length || description) {
    // Ficha sin pasos todavía: se publica igualmente, con el aviso de la hoja.
    stages.push({
      title: 'Preparación',
      procedureTitle: null,
      description: description,
      ingredients: ingredients,
      steps: [pending || 'Pasos pendientes de redactar.'],
    });
  }
  if (!stages.length) return null;

  return {
    id: slugify(name),
    name: name,
    code: code,
    section: section.name,
    sectionOrder: section.order,
    yieldAmount: yieldAmount,
    yieldUnit: yieldUnit,
    stages: stages,
  };
}

/**
 * Ejecutar manualmente desde el editor de Apps Script para comprobar
 * que el parseo funciona antes de desplegar (ver "Registro de ejecución").
 */
function testParse() {
  var recipes = buildRecipes();
  Logger.log('Recetas encontradas: ' + recipes.length);
  recipes.forEach(function (r) {
    Logger.log(r.code + ' ' + r.name + ' -> ' + r.stages.length + ' etapas, ' +
      r.stages[0].ingredients.length + ' ingredientes');
  });
}
