# Recetas del Taller — Rice to Meet You

App web para escalar y guiar paso a paso las preparaciones oficiales de RTMY,
leyendo en vivo el Excel "MANUAL COCINA" del Drive.

Archivos:
- `index.html`, `styles.css`, `app.js` → la app (estática, sin dependencias).
- `Code.gs` → el backend (Google Apps Script) que lee el Excel y sirve JSON.
- `recipes.json` → copia de datos generada una vez, usada solo como demo/respaldo
  si `CONFIG.DATA_URL` en `app.js` está vacío. En producción no se edita a mano:
  la app pide los datos en vivo al Apps Script.

Recetas incluidas automáticamente: cualquier hoja del Excel cuya celda **E1**
contenga el texto `FICHA TÉCNICA · ESCANDALLO DE COCINA` (la plantilla oficial).
Si añades una receta nueva copiando esa plantilla, aparecerá sola en la app la
próxima vez que se recargue — no hace falta tocar el código.

Quedan fuera por ahora: **JAMON** (formato antiguo, distinto a la plantilla) y
**ARROCES** (paellas — el Excel solo tiene el costeo, no los pasos de
preparación escritos). Cuando quieras incluirlas, dímelo y las incorporamos.

---

## 1. Desplegar el backend (Google Apps Script) — 5 minutos, una sola vez

Esto crea una API gratuita y sin límites reales de uso que lee el Excel y
devuelve solo cantidades/ingredientes/pasos (nunca costes ni precios de
compra, que se quedan en el Excel).

1. Abre el Excel "MANUAL COCINA" en Google Sheets (súbelo a Sheets si aún es
   un .xlsx: Archivo > Guardar como Hoja de cálculo de Google, para poder
   engancharle un script).
2. Menú **Extensiones > Apps Script**.
3. Borra el contenido de `Código.gs` y pega todo el contenido de nuestro
   archivo `Code.gs`.
4. Arriba a la izquierda, pon nombre al proyecto, por ejemplo
   "RTMY Recetas API".
5. Pulsa el botón ▶️ **Ejecutar** sobre la función `testParse` la primera vez
   (te pedirá autorizar el script con tu cuenta — es tuyo, es seguro).
   Revisa **Ver > Registro de ejecución**: debería listar 21 recetas.
6. **Desplegar > Nueva implementación**.
   - Tipo: **Aplicación web**.
   - Ejecutar como: **Yo (tu cuenta)**.
   - Quién tiene acceso: **Cualquier usuario**.
7. Pulsa **Desplegar** y autoriza de nuevo si te lo pide.
8. Copia la URL que termina en `/exec`. Esa es la API en vivo de las recetas.

Cada vez que edites una receta en el Excel (ingredientes, pasos, cantidades),
la API la sirve actualizada al instante — no hay que volver a desplegar nada,
salvo que cambies el propio código de `Code.gs`.

---

## 2. Conectar la app a esa API

Abre `app.js` y en la primera sección pega la URL copiada:

```js
const CONFIG = {
  DATA_URL: "https://script.google.com/macros/s/AKfycb.../exec",
  ...
};
```

---

## 3. Publicar gratis (GitHub Pages, sin límites de uso)

```bash
cd "recetas-app"
git init
git add index.html styles.css app.js recipes.json Code.gs README.md
git commit -m "Recetas del taller RTMY"
# crea un repositorio vacío en github.com (puede ser privado o público)
git remote add origin https://github.com/TU-USUARIO/rtmy-recetas.git
git push -u origin main
```

En GitHub: **Settings > Pages > Branch: main > /(root) > Save**.
En 1-2 minutos la app queda publicada en algo como:
`https://TU-USUARIO.github.io/rtmy-recetas/`

Es 100% gratis y no tiene límite de visitas para un uso normal como este.

---

## Calidad y recuperación

- La app intenta primero la API en vivo y, si no responde, carga
  `recipes.json` indicando que se está usando la copia local.
- Las sesiones, el paso actual y los temporizadores se guardan en el
  dispositivo para sobrevivir a una recarga accidental.
- Para ejecutar las comprobaciones automáticas:

```bash
npm run check
```

La rama `archive/production-2026-07-27` conserva la versión de producción
anterior a esta revisión. La nueva versión debe probarse desde su propia rama
antes de fusionarla con `main`.

---

## 4. Insertarla en Squarespace

La forma más limpia es un **iframe** apuntando a la URL de GitHub Pages,
dentro de un bloque de código en la página que quieras (por ejemplo
`/recetas-del-taller`):

```html
<div style="max-width:520px;margin:0 auto;">
  <iframe
    src="https://TU-USUARIO.github.io/rtmy-recetas/"
    style="width:100%;height:1400px;border:0;"
    loading="lazy"
    title="Recetas del taller RTMY">
  </iframe>
</div>
```

Ajusta el `height` a tu gusto (la app crece según la pantalla del pasos).

---

## Notas de diseño

Paleta y tipografía tomadas del kit de marca RTMY (mismo usado en el
link-in-bio): crema `#fdf5ec`, melocotón `#ecc7a6`, granate `#7a2426`,
marrón `#5b3a2c` · Playfair Display (títulos) + Poppins (texto).

## Notas de alcance (decisiones tomadas contigo)

- Solo se escala el **ingrediente principal** (el primero de la tabla de la
  receta base). Las cantidades que aparecen sueltas dentro del texto de los
  pasos (p. ej. "raciones de 150 gramos") no se recalculan: son tamaños de
  porción fijos, no proporcionales al lote.
- Las etapas posteriores a la primera (macerado, rebozado, emplatado, etc.)
  se pueden empezar de forma independiente, tal como pediste para las Tiras
  de Pollo.
- La pantalla final muestra los pasos completos de la última etapa de la
  receta (normalmente "EMPLATADO" o "ENVASADO") a modo de repaso de
  racionado/conservación, tal como están escritos en el Excel — no se
  inventa ningún dato adicional.
