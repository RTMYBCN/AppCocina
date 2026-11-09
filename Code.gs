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
  var CACHE_KEY = 'recipes-v4';
  var CACHE_CHUNK = 30000; // caracteres por trozo (límite de CacheService: 100 KB por valor)
  
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
    var cached = fresh ? null : cacheGetRecipes(cache);
    var recipes = cached || buildRecipes();
    if (!cached) {
      try {
        cachePutRecipes(cache, recipes);
      } catch (err) {
        // No se pudo guardar en la caché: se sirve igualmente, sin caché.
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
  
  // El JSON con los datos de planificación supera los 100 KB de un valor de caché:
  // se guarda troceado (recipes-v4:0, :1…) más una clave con el número de trozos.
  function cacheGetRecipes(cache) {
    var n = parseInt(cache.get(CACHE_KEY + ':n'), 10);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(CACHE_KEY + ':' + i);
    var parts = cache.getAll(keys);
    var json = '';
    for (var j = 0; j < n; j++) {
      if (parts[keys[j]] === null || parts[keys[j]] === undefined) return null; // trozo caducado
      json += parts[keys[j]];
    }
    try {
      return JSON.parse(json);
    } catch (err) {
      return null;
    }
  }
  
  function cachePutRecipes(cache, recipes) {
    var json = JSON.stringify(recipes);
    var n = Math.ceil(json.length / CACHE_CHUNK);
    var items = {};
    for (var i = 0; i < n; i++) items[CACHE_KEY + ':' + i] = json.substr(i * CACHE_CHUNK, CACHE_CHUNK);
    items[CACHE_KEY + ':n'] = String(n);
    cache.putAll(items, 300);
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
  
    var recipe = {
      id: slugify(name),
      name: name,
      code: code,
      section: section.name,
      sectionOrder: section.order,
      yieldAmount: yieldAmount,
      yieldUnit: yieldUnit,
      stages: stages,
    };
    // Datos de planificación (columnas M:Z de la ficha). Opcionales: una ficha sin
    // bloque, o con errores en él, se publica igual y la app sigue funcionando.
    var plan = parsePlan(values, stages);
    if (plan) recipe.plan = plan;
    return recipe;
  }
  
  /* ---------------- Planificación (bloque M:Z de cada ficha) ----------------
   * M1 = "PLANIFICACIÓN"; M2:N5 = cabecera (fase, vida útil, conservación,
   * fichas previas); fila de encabezados con "Bloque" en M; una fila por paso
   * de producción debajo. Se valida contra el texto real de los pasos: si una
   * fila no encaja se descarta y queda anotada en plan.warnings.
   * Claves de cada paso (cortas para que el JSON quepa en la caché):
   *   b bloque · s paso · d día · eq equipo · t tipo · pm min persona ·
   *   mm min máquina/espera · sc escala · rp repetición · ub uds por tanda ·
   *   aq otra cantidad · am min a esa cantidad · n notas · pend pendiente
   */
  var PLAN_COL = 13; // columna M
  var PLAN_HEADERS = ['Bloque', 'Paso', 'Inicio del paso', 'Día', 'Equipo', 'Tipo', 'Min persona',
    'Min máquina/espera', 'Escala', 'Repetición', 'Uds por tanda', 'Otra cantidad', 'Min a esa cantidad', 'Notas'];
  var PLAN_EQUIPOS = ['Thermomix grande', 'Thermomix pequeña', 'Thermomix (cualquiera)', 'Olla 25 L',
    'Olla pequeña', 'Olla grande (sin presión)', 'Vitro', 'Fuego de gas', 'Vitro/fuego (cualquiera)',
    'Inducción IsEasy', 'Horno', 'Airfryer', 'Freidora', 'Panificadora', 'Envasadora', 'Microondas',
    'Soplete', 'Nevera', 'Congelador', 'Manual'];
  var PLAN_TIPOS = ['Activo', 'Máquina', 'Cocción atendida', 'Espera', 'Activo + espera'];
  var PLAN_FASES = ['Producción', 'Servicio', 'Mixta'];
  var PLAN_CONSERVACION = ['Nevera', 'Congelador', 'Ambiente', 'Vacío en nevera', 'Mixta'];
  
  function planKey(s) {
    return String(s === null || s === undefined ? '' : s).trim().toLowerCase();
  }
  
  function planNumber(v) {
    if (isNumber(v)) return v;
    if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v.replace(',', '.')))) {
      return Number(v.replace(',', '.'));
    }
    return null;
  }
  
  function planPick(v, list) {
    var k = planKey(v);
    for (var i = 0; i < list.length; i++) if (planKey(list[i]) === k) return list[i];
    return null;
  }
  
  function parsePlan(values, stages) {
    var warnings = [];
    var markerRow = null;
    for (var r = 1; r <= Math.min(values.length, 12); r++) {
      if (planKey(valueAt(values, r, PLAN_COL)) === 'planificación') { markerRow = r; break; }
    }
    if (!markerRow) return null;
  
    var plan = { phase: null, shelfLifeDays: null, storage: null, requires: [], steps: [] };
  
    // Cabecera: etiquetas en M, valores en N (las 6 filas siguientes al marcador).
    for (r = markerRow + 1; r <= markerRow + 6; r++) {
      var label = planKey(valueAt(values, r, PLAN_COL));
      var val = valueAt(values, r, PLAN_COL + 1);
      if (label === 'fase de la ficha') {
        plan.phase = planPick(val, PLAN_FASES);
        if (val !== null && !plan.phase) warnings.push('Fase de la ficha no válida: ' + val);
      } else if (label === 'vida útil (días)') {
        plan.shelfLifeDays = planNumber(val);
      } else if (label === 'conservación') {
        plan.storage = planPick(val, PLAN_CONSERVACION);
        if (val !== null && !plan.storage) warnings.push('Conservación no válida: ' + val);
      } else if (label === 'requiere fichas previas') {
        var codes = String(val === null ? '' : val).split(/[,;\s]+/);
        for (var c = 0; c < codes.length; c++) {
          var code = codes[c].trim().toUpperCase();
          if (!code) continue;
          if (/^[BSC]\d{2}$/.test(code)) plan.requires.push(code);
          else warnings.push('Ficha previa no válida: ' + codes[c]);
        }
      }
    }
  
    // Fila de encabezados de la tabla (col. M = "Bloque").
    var headerRow = null;
    for (r = markerRow + 1; r <= Math.min(values.length, markerRow + 15); r++) {
      if (planKey(valueAt(values, r, PLAN_COL)) === 'bloque') { headerRow = r; break; }
    }
    if (!headerRow) {
      warnings.push('No se encuentra la tabla de pasos (encabezado "Bloque" en la columna M).');
      plan.warnings = warnings;
      return plan;
    }
    for (var h = 0; h < PLAN_HEADERS.length; h++) {
      if (planKey(valueAt(values, headerRow, PLAN_COL + h)) !== planKey(PLAN_HEADERS[h])) {
        warnings.push('Encabezado inesperado en la columna ' + (PLAN_COL + h) + ': se esperaba "' + PLAN_HEADERS[h] + '".');
        plan.warnings = warnings;
        return plan;
      }
    }
  
    for (r = headerRow + 1; r <= values.length; r++) {
      var bRaw = valueAt(values, r, PLAN_COL);
      var sRaw = valueAt(values, r, PLAN_COL + 1);
      if (bRaw === null && sRaw === null) continue;
      var where = 'fila ' + r;
      var b = planNumber(bRaw);
      var s = planNumber(sRaw);
      if (b === null || s === null || b < 1 || s < 1 || b % 1 || s % 1) {
        warnings.push(where + ': Bloque y Paso deben ser enteros.');
        continue;
      }
      var stage = stages[b - 1];
      var stepText = stage ? stage.steps[s - 1] : null;
      if (!stepText) {
        warnings.push(where + ': el bloque ' + b + ' paso ' + s + ' no existe en la ficha.');
        continue;
      }
      var ini = cleanText(valueAt(values, r, PLAN_COL + 2));
      if (ini && stepText.toLowerCase().indexOf(ini.toLowerCase()) !== 0) {
        warnings.push(where + ': el texto del paso ' + b + '.' + s + ' ya no empieza por "' + ini + '" (¿cambió el texto?). Fila ignorada.');
        continue;
      }
      var eq = planPick(valueAt(values, r, PLAN_COL + 4), PLAN_EQUIPOS);
      var tipo = planPick(valueAt(values, r, PLAN_COL + 5), PLAN_TIPOS);
      if (!eq) { warnings.push(where + ': Equipo no válido (' + valueAt(values, r, PLAN_COL + 4) + '). Fila ignorada.'); continue; }
      if (!tipo) { warnings.push(where + ': Tipo no válido (' + valueAt(values, r, PLAN_COL + 5) + '). Fila ignorada.'); continue; }
  
      var pm = planNumber(valueAt(values, r, PLAN_COL + 6));
      var mm = planNumber(valueAt(values, r, PLAN_COL + 7));
      var needsPm = tipo === 'Activo' || tipo === 'Cocción atendida' || tipo === 'Activo + espera';
      var needsMm = tipo === 'Máquina' || tipo === 'Espera' || tipo === 'Activo + espera';
      if ((needsPm && pm === null) || (needsMm && mm === null)) {
        warnings.push(where + ' (' + b + '.' + s + '): falta tiempo para el tipo "' + tipo + '".');
      }
      var rp = valueAt(values, r, PLAN_COL + 9);
      var rpOk = rp === null ? null : planPick(rp, ['POR TANDA', 'ENTRE TANDAS']);
      if (rp !== null && !rpOk) warnings.push(where + ': Repetición no válida (' + rp + ').');
      var scRaw = planKey(valueAt(values, r, PLAN_COL + 8));
      var notes = cleanText(valueAt(values, r, PLAN_COL + 13));
      var day = planNumber(valueAt(values, r, PLAN_COL + 3));
  
      var step = {
        b: b, s: s,
        d: day === null ? 1 : day,
        eq: eq, t: tipo,
        pm: pm, mm: mm,
        sc: scRaw === 'sí' || scRaw === 'si' ? true : scRaw === 'no' ? false : null,
      };
      if (rpOk) step.rp = rpOk;
      var ub = planNumber(valueAt(values, r, PLAN_COL + 10));
      var aq = planNumber(valueAt(values, r, PLAN_COL + 11));
      var am = planNumber(valueAt(values, r, PLAN_COL + 12));
      if (ub !== null) step.ub = ub;
      if (aq !== null && am !== null) { step.aq = aq; step.am = am; }
      if (notes) step.n = notes;
      if (notes && /PENDIENTE/i.test(notes)) step.pend = true;
      plan.steps.push(step);
    }
    if (warnings.length) plan.warnings = warnings;
    return plan;
  }
  
  /**
   * Ejecutar manualmente desde el editor de Apps Script para comprobar
   * que el parseo funciona antes de desplegar (ver "Registro de ejecución").
   */
  function testParse() {
    var recipes = buildRecipes();
    Logger.log('Recetas encontradas: ' + recipes.length);
    recipes.forEach(function (r) {
      var planInfo = r.plan
        ? ' | plan: ' + r.plan.steps.length + ' pasos' + (r.plan.warnings ? ', ' + r.plan.warnings.length + ' AVISOS' : '')
        : '';
      Logger.log(r.code + ' ' + r.name + ' -> ' + r.stages.length + ' etapas, ' +
        r.stages[0].ingredients.length + ' ingredientes' + planInfo);
      if (r.plan && r.plan.warnings) r.plan.warnings.forEach(function (w) { Logger.log('    ! ' + w); });
    });
  }
