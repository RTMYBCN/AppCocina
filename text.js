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
        `(?<![A-Za-zÁÉÍÓÚÜÑáéíóúüñ])(?<!\\d+\\s*(?:g|gramos?|kg|kilos?|ml|mililitros?|ud|unidades?)\\s+de\\s*)(el|la|los|las)?(\\s*)\\b(${escapeRegex(key)}(?:es|s)?)\\b`,
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
        `(\\d+(?:[.,]\\d+)?)(\\s*)(${UNIT_WORDS.join("|")})\\b(\\s+de\\s+(?:la\\s+|el\\s+|los\\s+|las\\s+)?)(${escapeRegex(key)}(?:es|s)?)\\b`,
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
  };

  if (typeof module !== "undefined" && module.exports) module.exports = RTMYText;
  if (typeof window !== "undefined") window.RTMYText = RTMYText;
})();
