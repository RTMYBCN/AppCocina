/* ============================================================
   Texto de los pasos: cantidades.

   Dos trabajos, en este orden:
   1) rescaleLiteralQuantities — la ficha a veces escribe la cantidad
      dentro del paso ("agregar los 40 gramos de aceite"). Si el cocinero
      escala la receta, ese número hay que reescribirlo o manda una
      cantidad que no es la del lote.
   2) injectIngredientQuantities — cuando el paso solo nombra el
      ingrediente ("saltear la cebolla"), se le añade la cantidad real.

   Fichero aparte de app.js para poder probarlo con node (no toca el DOM).
   ============================================================ */
// Envuelto en una función para no soltar nombres sueltos en la página:
// lo único que se publica es RTMYText.
(function () {
  function fmtQty(n) {
    if (n === null || n === undefined) return "";
    const rounded = n >= 10 ? Math.round(n) : Math.round(n * 10) / 10;
    return rounded.toString().replace(".", ",");
  }

  const UNIT_GENDER = {
    g: "m", gramo: "m", gramos: "m",
    kg: "m", kilo: "m", kilos: "m",
    ml: "m", mililitro: "m", mililitros: "m",
    cl: "m", l: "m", litro: "m", litros: "m",
    ud: "f", uds: "f", u: "f", "u.": "f", unidad: "f", unidades: "f",
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

  // `skipKeys` son los ingredientes que ya se han nombrado con su cantidad
  // en un paso anterior: repetir "los 1400 g de cebolla" en cada paso
  // ensucia el texto y a veces lo descuadra ("los 1400 g de cebolla debe
  // reducir"). Devuelve el texto y qué ingredientes ha nombrado.
  function injectTracked(text, ingredients, skipKeys) {
    if (!ingredients || !ingredients.length) return { text, named: [] };

    const byKey = new Map();
    ingredients.forEach((ing) => {
      const key = ingredientKeyTerm(ing.name);
      if (key.length < 3 || GENERIC_KEY_BLOCKLIST.has(key)) return;
      if (!byKey.has(key)) byKey.set(key, []);
      byKey.get(key).push(ing);
    });

    let result = text;
    const named = [];
    byKey.forEach((list, key) => {
      if (list.length > 1) return; // ambiguous shared first word — don't guess
      if (skipKeys && skipKeys.has(key)) return; // ya se dijo antes
      const ing = list[0];
      // El primer lookbehind evita partir una palabra: sin él, el "el" de
      // "del limón" cuenta como artículo y queda "dlos 4 g de limón".
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
        `(?<![A-Za-zÁÉÍÓÚÜÑáéíóúüñ])(?<!\\d+\\s*(?:g|gramos?|kg|kilos?|ml|mililitros?|ud|unidades?)\\s+del?\\s*)` +
          // Si la frase ya pone un número delante del ingrediente ("añadir 2
          // huevos"), es una parte, no el lote: añadirle el total del lote
          // detrás solo confunde.
          `(?<!\\d[\\d.,]*\\s{0,3})` +
          // Lo mismo cuando entre el número y el ingrediente hay una medida
          // de cocina que no es la de la ficha: "1 tarrina de jamón",
          // "3 L de aceite", "2 rebanadas de pan". Son porciones, no el lote.
          `(?<!\\d[\\d.,]*\\s*[A-Za-zÁÉÍÓÚÜÑáéíóúüñ.]{0,12}\\s+del?\\s*)` +
          `(el|la|los|las)?(\\s*)\\b(${escapeRegex(key)}(?:es|s)?)\\b` +
          // Y si la ficha ya puso la cantidad entre paréntesis detrás
          // ("jamón (75 g)"), no se le añade otra delante.
          `(?!\\s*\\(\\s*\\d+(?:[.,]\\d+)?\\s*(?:${UNIT_WORDS.join("|")})\\s*\\))`,
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
      named.push(key);
    });
    return { text: result, named };
  }

  function injectIngredientQuantities(text, ingredients, skipKeys) {
    return injectTracked(text, ingredients, skipKeys).text;
  }

  /* ---------------- Cantidades ya escritas en el paso ----------------
     La ficha escribe a veces la cantidad dentro del texto ("agregar los 40
     gramos de aceite"). Al escalar la receta ese número se queda viejo, así
     que se reescribe — pero solo cuando no hay duda:

     - el número va seguido de una unidad y de "de" + ese ingrediente, y
     - coincide con la cantidad de la ficha sin escalar.

     Así un "30 minutos" o un "105 °C" no se tocan nunca, ni tampoco un
     "raciones de 150 g" que no corresponde a ningún ingrediente. Si la
     ficha usa otra unidad de la misma familia ("1 kg de cebolla" con la
     ficha en gramos), se compara convertido y se responde en la unidad que
     usa el texto. */

  const UNIT_SCALE = {
    g: { family: "masa", k: 1 },
    gr: { family: "masa", k: 1 },
    gramo: { family: "masa", k: 1 },
    gramos: { family: "masa", k: 1 },
    kg: { family: "masa", k: 1000 },
    kilo: { family: "masa", k: 1000 },
    kilos: { family: "masa", k: 1000 },
    kilogramo: { family: "masa", k: 1000 },
    kilogramos: { family: "masa", k: 1000 },
    ml: { family: "volumen", k: 1 },
    mililitro: { family: "volumen", k: 1 },
    mililitros: { family: "volumen", k: 1 },
    cl: { family: "volumen", k: 10 },
    l: { family: "volumen", k: 1000 },
    litro: { family: "volumen", k: 1000 },
    litros: { family: "volumen", k: 1000 },
    ud: { family: "unidades", k: 1 },
    uds: { family: "unidades", k: 1 },
    unidad: { family: "unidades", k: 1 },
    unidades: { family: "unidades", k: 1 },
    u: { family: "unidades", k: 1 },
    "u.": { family: "unidades", k: 1 },
  };

  /* Las unidades que se cuentan no admiten decimales: medio huevo no se
     echa a la olla. Al escalar se redondea SIEMPRE hacia arriba — que
     sobre, no que falte. Decisión de Leo (2026-11-11). */
  function isCountableUnit(unit) {
    const u = (unit || "").trim().toLowerCase();
    return Boolean(UNIT_SCALE[u]) && UNIT_SCALE[u].family === "unidades";
  }

  function scaleQty(qty, unit, ratio) {
    const value = (Number(qty) || 0) * (ratio === undefined || ratio === null ? 1 : ratio);
    if (!isCountableUnit(unit)) return value;
    return Math.ceil(value - 1e-9);
  }

  const UNIT_WORDS = Object.keys(UNIT_SCALE)
    .sort((a, b) => b.length - a.length)
    .map(escapeRegex);

  function rescaleTracked(text, baseIngredients, ratio) {
    const touched = [];
    if (!text || !baseIngredients || !baseIngredients.length) return { text, touched };

    let out = text;
    baseIngredients.forEach((ing) => {
      const base = UNIT_SCALE[(ing.unit || "").toLowerCase()];
      if (!base || !(ing.qty > 0)) return;
      const key = ingredientKeyTerm(ing.name);
      if (key.length < 3 || GENERIC_KEY_BLOCKLIST.has(key)) return;
      const re = new RegExp(
        `(\\d+(?:[.,]\\d+)?)(\\s*)(${UNIT_WORDS.join("|")})\\b(\\s+de\\s+(?:la\\s+|el\\s+|los\\s+|las\\s+)?|\\s+del\\s+)(${escapeRegex(key)}(?:es|s)?)\\b`,
        "gi"
      );
      out = out.replace(re, (whole, num, sep, unitWord, middle, name) => {
        const written = UNIT_SCALE[unitWord.toLowerCase()];
        if (!written || written.family !== base.family) return whole;
        const valueInBase = parseFloat(num.replace(",", ".")) * written.k;
        // Solo se reescribe si el número es justo el de la ficha.
        if (!(Math.abs(valueInBase - ing.qty * base.k) < 0.01)) return whole;
        touched.push(key); // ya queda dicha: que no se repita en pasos siguientes
        if (!ratio || Math.abs(ratio - 1) < 1e-9) return whole;
        const scaled =
          written.family === "unidades"
            ? scaleQty(ing.qty, ing.unit, ratio)
            : (ing.qty * base.k * ratio) / written.k;
        return `${fmtQty(scaled)}${sep}${unitWord}${middle}${name}`;
      });

      // Tercer patrón: la cantidad de la ficha escrita con una medida de
      // cocina que no es una unidad ("las 11 hojas de gelatina" cuando la
      // ficha dice 11 ud). Si el número es justo el de la ficha, se escala
      // igual; si no lo fuera sería una porción y no se toca.
      const reMeasure = new RegExp(
        `(?<![\\d.,])(\\d+(?:[.,]\\d+)?)(\\s+)([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]{2,12})(\\s+del?\\s+)(${escapeRegex(key)}(?:es|s)?)\\b`,
        "gi"
      );
      out = out.replace(reMeasure, (whole, num, sep, word, middle, name) => {
        if (UNIT_SCALE[word.toLowerCase()]) return whole; // lo lleva el primer patrón
        if (!(Math.abs(parseFloat(num.replace(",", ".")) - ing.qty) < 0.01)) return whole;
        touched.push(key);
        return `${fmtQty(scaleQty(ing.qty, ing.unit, ratio || 1))}${sep}${word}${middle}${name}`;
      });

      // Segundo patrón: la ficha nombra la cantidad sin unidad, pegada al
      // ingrediente ("otra con 4 huevos batidos"), a veces con la unidad
      // repetida detrás entre paréntesis ("4 huevos (4 ud)"). Si no se
      // reescribe, al escalar la frase se contradice sola: "4 huevos (8 ud)".
      // Solo se toca cuando el número es exactamente el de la ficha.
      const reBare = new RegExp(
        `(?<![\\d.,])(\\d+(?:[.,]\\d+)?)(\\s+)(${escapeRegex(key)}(?:es|s)?)\\b` +
          `(\\s*\\(\\s*\\d+(?:[.,]\\d+)?\\s*(?:${UNIT_WORDS.join("|")})\\s*\\))?`,
        "gi"
      );
      out = out.replace(reBare, (whole, num, sep, name) => {
        if (!(Math.abs(parseFloat(num.replace(",", ".")) - ing.qty) < 0.01)) return whole;
        touched.push(key);
        // El paréntesis sobra: el número de delante ya dice la cantidad.
        return `${fmtQty(scaleQty(ing.qty, ing.unit, ratio || 1))}${sep}${name}`;
      });
    });
    return { text: out, touched };
  }

  function rescaleLiteralQuantities(text, baseIngredients, ratio) {
    return rescaleTracked(text, baseIngredients, ratio).text;
  }

  /* Único punto de entrada: primero se actualizan las cantidades escritas,
     después se añaden las que faltan. */
  function stepTextWithQuantities(text, baseIngredients, ratio, skipKeys) {
    return stepTexts([text], baseIngredients, ratio, skipKeys)[0];
  }

  /* Los pasos de una etapa se preparan juntos porque la cantidad de cada
     ingrediente se dice UNA vez, la primera que aparece. Repetirla en cada
     paso no aporta nada y descuadra frases ("los 1400 g de cebolla debe
     reducir"). */
  function stepTexts(steps, baseIngredients, ratio, skipKeys) {
    const said = new Set(skipKeys || []);
    const scaled = (baseIngredients || []).map((ing) => ({
      ...ing,
      qty: scaleQty(ing.qty, ing.unit, ratio || 1),
    }));
    return (steps || []).map((text) => {
      const step = rescaleTracked(text, baseIngredients, ratio);
      step.touched.forEach((k) => said.add(k));
      const out = injectTracked(step.text, scaled, said);
      out.named.forEach((k) => said.add(k));
      return out.text;
    });
  }

  /* ---------------- Fases de ingredientes ----------------
     Una receta de varias etapas lleva TODA la tabla de ingredientes en la
     primera etapa (es la que se escala). Sin más datos, el cocinero que solo
     va a hacer la masa de las croquetas alista también los huevos y el pan
     rallado del boleado. Cada ingrediente trae su fase (columna B de la ficha)
     y cada etapa declara qué fases se alistan en ella (celda de al lado del
     título). Aquí se cruzan las dos cosas. */

  function phaseKey(value) {
    return (value === undefined || value === null ? "" : String(value))
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  }

  // La declaración puede llegar como lista ya partida o como el texto crudo
  // de la celda ("Setas, Bechamel, Masa").
  function stagePhases(stage) {
    const raw = (stage || {}).phases;
    if (Array.isArray(raw)) return raw.map((x) => String(x).trim()).filter(Boolean);
    if (typeof raw === "string") return raw.split(/[,;]/).map((x) => x.trim()).filter(Boolean);
    return [];
  }

  function recipeUsesPhases(recipe) {
    const stages = (recipe || {}).stages || [];
    return stages.some((stage) => stagePhases(stage).length > 0);
  }

  /* Qué ingredientes hay que alistar para una etapa. Dos reglas de seguridad,
     porque en cocina esconder un ingrediente es peor que mostrar uno de más:
     - Si la ficha no declara ninguna fase, se comporta como siempre: todo en
       la primera etapa.
     - Un ingrediente cuya fase no reclama ninguna etapa (celda en blanco, una
       errata) cae en la primera etapa en vez de desaparecer. */
  function ingredientsForStage(recipe, stageIndex) {
    const stages = (recipe || {}).stages || [];
    const stage = stages[stageIndex];
    if (!stage) return [];
    const table = ((stages[0] || {}).ingredients || []).slice();
    const own = stage.ingredients && stage.ingredients.length ? stage.ingredients : null;
    if (!recipeUsesPhases(recipe)) return stageIndex === 0 ? table : own || [];

    const claimed = new Set();
    stages.forEach((s) => stagePhases(s).forEach((p) => claimed.add(phaseKey(p))));
    const wanted = new Set(stagePhases(stage).map(phaseKey));
    const orphan = stageIndex === 0;
    return table.filter((ing) => {
      const key = phaseKey(ing.phase);
      if (wanted.has(key)) return true;
      return orphan && !claimed.has(key);
    });
  }

  // Grupos de la tabla de ingredientes, en el orden en que aparecen en la
  // ficha. Devuelve un único grupo sin nombre cuando la ficha no usa fases.
  function groupByPhase(ingredients) {
    const groups = [];
    const byKey = new Map();
    (ingredients || []).forEach((ing, index) => {
      const key = phaseKey(ing.phase);
      if (!byKey.has(key)) {
        const group = { phase: (ing.phase || "").trim(), items: [] };
        byKey.set(key, group);
        groups.push(group);
      }
      byKey.get(key).items.push({ ingredient: ing, index });
    });
    return groups;
  }

  const RTMYText = {
    fmtQty,
    isCountableUnit,
    scaleQty,
    escapeRegex,
    ingredientKeyTerm,
    articleFor,
    injectIngredientQuantities,
    rescaleLiteralQuantities,
    stepTextWithQuantities,
    stepTexts,
    phaseKey,
    stagePhases,
    recipeUsesPhases,
    ingredientsForStage,
    groupByPhase,
  };

  if (typeof module !== "undefined" && module.exports) module.exports = RTMYText;
  if (typeof window !== "undefined") window.RTMYText = RTMYText;
})();
