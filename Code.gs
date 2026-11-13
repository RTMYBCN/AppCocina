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
    // ?jornada=CODE: estado de la jornada compartida (ver JORNADA_COMPARTIDA.md).
    if (e && e.parameter && e.parameter.jornada) return jornadaGet(e);
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
  
  // Fases declaradas en la celda de al lado del título de la etapa (columna C):
  // "Setas, Bechamel, Masa". Dicen qué ingredientes de la tabla se alistan en
  // esa etapa, para que el cocinero que solo hace la masa no saque de la nevera
  // los huevos del empanado.
  function splitPhases(text) {
    if (!text) return [];
    return text
      .split(/[,;]/)
      .map(function (x) { return x.trim(); })
      .filter(function (x) { return x.length > 0; });
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
          current = { title: null, phases: [], steps: [] };
          blocks.push(current);
        }
        current.steps = current.steps.concat(splitSteps(bVal));
      } else {
        current = { title: bVal, phases: splitPhases(cleanText(valueAt(values, r, 3))), steps: [] };
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
          // Qué fases de la tabla se alistan en esta etapa (vacío = ninguna propia).
          phases: b.phases || [],
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
        phases: [],
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
  
  /* ===================================================================
     JORNADA COMPARTIDA
  
     Varios dispositivos de la cocina sobre el mismo plan. El estado vive en
     las propiedades del script (PropertiesService), nunca en el Excel de
     recetas. El contrato está escrito en JORNADA_COMPARTIDA.md.
  
     Lecturas por GET (?jornada=CODE[&rev=N][&plan=1]), escrituras por POST
     con cuerpo JSON {action, code, …}. Toda escritura toma el candado del
     script: dos tablets que guardan a la vez no se pisan.
     =================================================================== */
  
  var J_PLAN_CHUNK = 8000;   // caracteres por trozo (límite de una propiedad: 9 KB)
  var J_TTL_MS = 36 * 60 * 60 * 1000; // una jornada sin tocar se borra
  var J_LOCK_MS = 15000;
  var J_MAX_SLOTS = 8;
  var J_MAX_TIMERS = 24;
  var J_MAX_DONE = 600;
  
  function jProps() {
    return PropertiesService.getScriptProperties();
  }
  
  function jKey(code, part) {
    return 'j:' + code + ':' + part;
  }
  
  function jOut(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  }
  
  function jError(message, extra) {
    var out = { ok: false, error: message, now: Date.now() };
    if (extra) for (var k in extra) out[k] = extra[k];
    return jOut(out);
  }
  
  // Un código de jornada son cuatro dígitos. Se acepta con espacios o guiones.
  function jNormCode(value) {
    var digits = String(value === null || value === undefined ? '' : value).replace(/\D/g, '');
    return /^\d{4}$/.test(digits) ? digits : null;
  }
  
  function jParse(raw, fallback) {
    if (!raw) return fallback;
    try {
      return JSON.parse(raw);
    } catch (err) {
      return fallback;
    }
  }
  
  // Un objeto agrupado por receta; cualquier otra cosa (o un formato viejo) se
  // descarta en vez de romper la jornada.
  function jMap(value) {
    return value && typeof value === 'object' && !(value instanceof Array) ? value : {};
  }
  
  function jEmptyLive() {
    return { rev: 0, at: 0, done: {}, owner: {}, timers: [], people: [] };
  }
  
  function jReadMeta(props, code) {
    return jParse(props.getProperty(jKey(code, 'meta')), null);
  }
  
  function jReadLive(props, code) {
    var live = jParse(props.getProperty(jKey(code, 'live')), null);
    if (!live) return jEmptyLive();
    live.done = jMap(live.done);
    live.owner = jMap(live.owner);
    live.timers = live.timers || [];
    live.people = live.people || [];
    return live;
  }
  
  // El plan no cabe en una propiedad: se trocea igual que la caché de recetas.
  function jReadPlan(props, code) {
    var n = parseInt(props.getProperty(jKey(code, 'plan:n')), 10);
    if (!n) return null;
    var json = '';
    for (var i = 0; i < n; i++) {
      var part = props.getProperty(jKey(code, 'plan:' + i));
      if (part === null || part === undefined) return null;
      json += part;
    }
    return jParse(json, null);
  }
  
  function jWritePlan(props, code, plan) {
    var old = parseInt(props.getProperty(jKey(code, 'plan:n')), 10) || 0;
    var json = JSON.stringify(plan);
    var n = Math.ceil(json.length / J_PLAN_CHUNK);
    for (var i = 0; i < n; i++) props.setProperty(jKey(code, 'plan:' + i), json.substr(i * J_PLAN_CHUNK, J_PLAN_CHUNK));
    for (var j = n; j < old; j++) props.deleteProperty(jKey(code, 'plan:' + j));
    props.setProperty(jKey(code, 'plan:n'), String(n));
  }
  
  function jDelete(props, code) {
    var n = parseInt(props.getProperty(jKey(code, 'plan:n')), 10) || 0;
    for (var i = 0; i < n; i++) props.deleteProperty(jKey(code, 'plan:' + i));
    props.deleteProperty(jKey(code, 'plan:n'));
    props.deleteProperty(jKey(code, 'meta'));
    props.deleteProperty(jKey(code, 'live'));
    var index = jParse(props.getProperty('j:index'), {}) || {};
    delete index[code];
    props.setProperty('j:index', JSON.stringify(index));
  }
  
  function jTouch(props, code, at) {
    var index = jParse(props.getProperty('j:index'), {}) || {};
    index[code] = at;
    props.setProperty('j:index', JSON.stringify(index));
  }
  
  // Las jornadas viejas se barren en la siguiente escritura: no hace falta
  // activador y así el almacén no crece sin fin.
  function jPrune(props, now) {
    var index = jParse(props.getProperty('j:index'), {}) || {};
    var codes = Object.keys(index);
    for (var i = 0; i < codes.length; i++) {
      if (now - (index[codes[i]] || 0) > J_TTL_MS) jDelete(props, codes[i]);
    }
  }
  
  function jFreeCode(props) {
    for (var attempt = 0; attempt < 40; attempt++) {
      var code = String(1000 + Math.floor(Math.random() * 9000));
      if (!props.getProperty(jKey(code, 'meta'))) return code;
    }
    return null;
  }
  
  /* `done` y `owner` se guardan AGRUPADOS POR RECETA:
       {"croquetas-de-jamon": ["9|12:3", "1|2:0"]}  ->  "bloque|paso:slot"
     Con la clave entera repetida en cada entrada (receta|bloque|paso:slot),
     una jornada de 245 pasos marcados pasaba de los 9 KB de una propiedad.
     Agrupada ocupa un tercio. */
  function jSplit(key) {
    var text = String(key || '');
    var i = text.indexOf('|');
    if (i <= 0 || i === text.length - 1) return null;
    return { r: text.slice(0, i), s: text.slice(i + 1) };
  }
  
  // Slot de quien marcó/cogió ese paso, o null si no está.
  function jGet(map, key) {
    var p = jSplit(key);
    if (!p || !map[p.r]) return null;
    var list = map[p.r];
    for (var i = 0; i < list.length; i++) {
      var parts = String(list[i]).split(':');
      if (parts[0] === p.s) return Number(parts[1]);
    }
    return null;
  }
  
  function jDrop(map, key) {
    var p = jSplit(key);
    if (!p || !map[p.r]) return false;
    var list = map[p.r];
    var dropped = false;
    for (var i = list.length - 1; i >= 0; i--) {
      if (String(list[i]).split(':')[0] === p.s) {
        list.splice(i, 1);
        dropped = true;
      }
    }
    if (!list.length) delete map[p.r];
    return dropped;
  }
  
  function jSet(map, key, slot) {
    var p = jSplit(key);
    if (!p) return false;
    jDrop(map, key);
    if (!map[p.r]) map[p.r] = [];
    map[p.r].push(p.s + ':' + slot);
    return true;
  }
  
  function jCount(map) {
    var n = 0;
    for (var k in map) n += map[k].length;
    return n;
  }
  
  function jSlot(value, people) {
    var slot = Math.round(Number(value));
    if (!isFinite(slot) || slot < 0) return 0;
    var max = Math.max(1, Math.min(J_MAX_SLOTS, people || J_MAX_SLOTS)) - 1;
    return Math.min(slot, max);
  }
  
  function jName(value) {
    return String(value === null || value === undefined ? '' : value).trim().slice(0, 24);
  }
  
  // Quién es este dispositivo. Si vuelve con el mismo nombre recupera su sitio;
  // si no, se le da el primer puesto libre del plan.
  function jAssignSlot(live, wanted, name, people) {
    var i;
    if (name) {
      for (i = 0; i < live.people.length; i++) if (jName(live.people[i].n) === name) return live.people[i].s;
    }
    if (wanted !== null && wanted !== undefined && wanted !== '') return jSlot(wanted, people);
    var taken = {};
    for (i = 0; i < live.people.length; i++) taken[live.people[i].s] = true;
    var max = Math.max(1, Math.min(J_MAX_SLOTS, people || J_MAX_SLOTS));
    for (i = 0; i < max; i++) if (!taken[i]) return i;
    return 0;
  }
  
  function jSeen(live, slot, name, now) {
    for (var i = 0; i < live.people.length; i++) {
      if (live.people[i].s === slot) {
        if (name) live.people[i].n = name;
        live.people[i].a = now;
        return;
      }
    }
    if (live.people.length < J_MAX_SLOTS) live.people.push({ s: slot, n: name || '', a: now });
  }
  
  var J_PERSON_TTL_MS = 6 * 60 * 60 * 1000; // una tablet que no da señales en 6 h deja su puesto libre

  // Si una tablet se apaga y no vuelve, su puesto se libera: si no, el
  // siguiente que entra se encuentra todos los sitios cogidos.
  function jPrunePeople(live, now) {
    for (var i = live.people.length - 1; i >= 0; i--) {
      if (now - (live.people[i].a || 0) > J_PERSON_TTL_MS) live.people.splice(i, 1);
    }
    return live;
  }

  function jSaveLive(props, code, live, now) {
    live.rev = (live.rev || 0) + 1;
    live.at = now;
    jPrunePeople(live, now);
    if (live.timers.length > J_MAX_TIMERS) live.timers = live.timers.slice(-J_MAX_TIMERS);
    props.setProperty(jKey(code, 'live'), JSON.stringify(live));
    jTouch(props, code, now);
    return live;
  }
  
  /* ---- Lectura (GET) ---- */
  
  function jornadaGet(e) {
    var p = (e && e.parameter) || {};
    var code = jNormCode(p.jornada);
    var now = Date.now();
    if (!code) return jError('Código de jornada no válido.');
    var props = jProps();
    var meta = jReadMeta(props, code);
    if (!meta) return jError('Esa jornada no existe.', { missing: true, code: code });
    var live = jReadLive(props, code);
    var out = { ok: true, now: now, code: code, rev: live.rev || 0, planRev: meta.planRev || 1 };
    // Sondeo: si el dispositivo ya tiene este rev, no se le manda nada más.
    var rev = parseInt(p.rev, 10);
    if (isFinite(rev) && rev === (live.rev || 0) && p.plan !== '1') {
      out.unchanged = true;
      return jOut(out);
    }
    out.live = live;
    if (p.plan === '1' || !isFinite(rev)) {
      out.meta = meta;
      out.plan = jReadPlan(props, code);
    }
    return jOut(out);
  }
  
  /* ---- Escritura (POST) ---- */
  
  function doPost(e) {
    var body = jParse(e && e.postData && e.postData.contents, null);
    if (!body || typeof body !== 'object') return jError('Cuerpo de la petición no válido.');
    var action = String(body.action || '');
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(J_LOCK_MS)) return jError('El servidor está ocupado, inténtalo otra vez.', { retry: true });
    try {
      if (action === 'create') return jornadaCreate(body);
      if (action === 'join') return jornadaJoin(body);
      if (action === 'patch') return jornadaPatch(body);
      if (action === 'close') return jornadaClose(body);
      return jError('Acción desconocida: ' + action);
    } catch (err) {
      return jError('Error del servidor: ' + (err && err.message ? err.message : err));
    } finally {
      lock.releaseLock();
    }
  }
  
  function jornadaCreate(body) {
    var props = jProps();
    var now = Date.now();
    jPrune(props, now);
    if (!body.plan || typeof body.plan !== 'object') return jError('Falta el plan.');
    var code = body.code === null || body.code === undefined || body.code === '' ? jFreeCode(props) : jNormCode(body.code);
    if (!code) return jError('Código de jornada no válido.');
    var existing = jReadMeta(props, code);
    if (existing && !body.replace) return jError('Ese código ya está en uso.', { taken: true, code: code });
    var people = Math.max(1, Math.min(J_MAX_SLOTS, Math.round(Number((body.inputs || {}).people) || 1)));
    var meta = {
      createdAt: existing ? existing.createdAt : now,
      at: now,
      planRev: existing ? (existing.planRev || 1) + 1 : 1,
      plannerVersion: Number(body.plannerVersion) || 0,
      dataStamp: body.dataStamp || null,
      inputs: body.inputs || null,
      people: people,
    };
    jWritePlan(props, code, body.plan);
    props.setProperty(jKey(code, 'meta'), JSON.stringify(meta));
    // Al recalcular el plan se conservan los nombres y lo ya hecho: las claves
    // de los pasos no cambian si no cambia la selección.
    var live = existing ? jPrunePeople(jReadLive(props, code), now) : jEmptyLive();
    if (!existing) live.done = {};
    var name = jName(body.name);
    var slot = jAssignSlot(live, body.slot, name, people);
    jSeen(live, slot, name, now);
    jSaveLive(props, code, live, now);
    return jOut({ ok: true, code: code, rev: live.rev, planRev: meta.planRev, slot: slot, now: now, replaced: !!existing });
  }
  
  function jornadaJoin(body) {
    var props = jProps();
    var now = Date.now();
    var code = jNormCode(body.code);
    if (!code) return jError('Código de jornada no válido.');
    var meta = jReadMeta(props, code);
    if (!meta) return jError('Esa jornada no existe.', { missing: true, code: code });
    var live = jPrunePeople(jReadLive(props, code), now);
    var name = jName(body.name);
    var slot = jAssignSlot(live, body.slot, name, meta.people);
    jSeen(live, slot, name, now);
    jSaveLive(props, code, live, now);
    return jOut({
      ok: true, code: code, rev: live.rev, planRev: meta.planRev || 1, slot: slot, now: now,
      meta: meta, plan: jReadPlan(props, code), live: live,
    });
  }
  
  function jornadaClose(body) {
    var code = jNormCode(body.code);
    if (!code) return jError('Código de jornada no válido.');
    jDelete(jProps(), code);
    return jOut({ ok: true, code: code, now: Date.now() });
  }
  
  // Varias operaciones en una sola escritura: un candado y un solo rev nuevo.
  function jornadaPatch(body) {
    var props = jProps();
    var now = Date.now();
    var code = jNormCode(body.code);
    if (!code) return jError('Código de jornada no válido.');
    var meta = jReadMeta(props, code);
    if (!meta) return jError('Esa jornada no existe.', { missing: true, code: code });
    var ops = body.ops;
    if (!ops || !ops.length) ops = [];
    var live = jPrunePeople(jReadLive(props, code), now);
    var slot = jSlot(body.slot, meta.people);
    var rejected = [];
    var changed = false;
  
    for (var i = 0; i < ops.length; i++) {
      var op = ops[i] || {};
      var key = String(op.key || '').slice(0, 40);
      var who = op.slot === null || op.slot === undefined ? slot : jSlot(op.slot, meta.people);
      switch (String(op.op)) {
        case 'done':
          if (!key) break;
          if (jGet(live.done, key) === null && jCount(live.done) >= J_MAX_DONE) {
            rejected.push({ op: 'done', key: key, error: 'demasiadas tareas marcadas' });
            break;
          }
          jSet(live.done, key, who);
          changed = true;
          break;
        case 'undone':
          if (key && jDrop(live.done, key)) changed = true;
          break;
        case 'claim': {
          if (!key) break;
          var owner = jGet(live.owner, key);
          if (owner !== null && owner !== who) {
            // Gana el primero: la tarea no se roba, se avisa.
            rejected.push({ op: 'claim', key: key, owner: owner });
            break;
          }
          if (owner === null) {
            jSet(live.owner, key, who);
            changed = true;
          }
          break;
        }
        case 'release': {
          if (!key) break;
          if (jGet(live.owner, key) === who && jDrop(live.owner, key)) changed = true;
          break;
        }
        case 'timer': {
          var seconds = Math.max(1, Math.min(24 * 3600, Math.round(Number(op.seconds) || 0)));
          var id = String(op.id || who + '-' + now).slice(0, 24);
          for (var t = live.timers.length - 1; t >= 0; t--) if (live.timers[t].i === id) live.timers.splice(t, 1);
          // El endsAt lo pone el servidor con SU reloj: dos tablets desfasadas
          // ven la misma cuenta atrás.
          live.timers.push({ i: id, k: key || null, l: String(op.label || '').slice(0, 40), e: now + seconds * 1000, b: who });
          changed = true;
          break;
        }
        case 'timerStop': {
          var stopId = String(op.id || '');
          for (var s = live.timers.length - 1; s >= 0; s--) {
            if (live.timers[s].i === stopId) {
              live.timers.splice(s, 1);
              changed = true;
            }
          }
          break;
        }
        case 'ping':
          break;
        default:
          rejected.push({ op: String(op.op), key: key, error: 'desconocida' });
      }
    }
  
    jSeen(live, slot, jName(body.name), now);
    jSaveLive(props, code, live, now);
    jPrune(props, now);
    return jOut({ ok: true, code: code, rev: live.rev, planRev: meta.planRev || 1, now: now, live: live, rejected: rejected, changed: changed });
  }
