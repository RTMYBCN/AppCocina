# Jornada compartida — modelo de datos y protocolo

Varios dispositivos de la cocina trabajan sobre el **mismo plan**: uno lo
calcula, los demás se enganchan con un **código de jornada** de cuatro dígitos
y cada cocinero ve lo suyo. Lo que marca uno lo ven los demás.

El backend es el mismo Apps Script que publica las recetas. El estado vive en
`PropertiesService` (propiedades del script), no en el Excel: el Excel de
recetas no se toca.

## 1. Qué se comparte y qué no

| Compartido (va al servidor) | De cada dispositivo (se queda en `localStorage`) |
| --- | --- |
| El plan calculado (pasos, horas, persona, aparato) | La receta que alguien abre a mano fuera del plan |
| Las entradas del plan (recetas, personas, hora, turno) | La vista en la que está (`persona` / `linea`), el día que mira, el scroll |
| Qué tareas están hechas y quién las hizo | Quién es este dispositivo (su `slot`) |
| El dueño de cada tarea (quien la ha cogido) | El desfase de reloj contra el servidor |
| Los temporizadores de máquina | Los temporizadores de una preparación suelta |

La navegación **no** se sincroniza: si se sincronizara, a un cocinero le
cambiaría la pantalla mientras cocina.

## 2. Reglas de conflicto

- **Hecho**: gana la última escritura. Marcar y desmarcar son operaciones
  distintas (`done` / `undone`), nunca un interruptor, para que dos
  dispositivos no se alternen el valor.
- **Coger una tarea**: gana el primero. Si llega un `claim` sobre una tarea que
  ya tiene dueño, el servidor lo **rechaza** y devuelve quién la tiene; el
  dispositivo avisa en vez de robarla. El dueño puede soltarla (`release`).
- **Temporizadores**: el `id` lo pone el dispositivo (`slot` + contador), así
  que no colisionan. Pararlo lo puede hacer cualquiera.
- **El plan**: solo se reescribe entero, con `replace: true`. Recalcular sube
  el `planRev` y los demás dispositivos lo recargan y avisan al cocinero.

## 3. Relojes

Cada respuesta del servidor trae `now` (epoch ms del servidor). El dispositivo
guarda `offset = now - Date.now()` y **los temporizadores compartidos guardan
`endsAt` en hora de servidor**. Al pintar la cuenta atrás se usa
`endsAt - (Date.now() + offset)`. Así dos tablets con el reloj desfasado
muestran lo mismo.

## 4. Red caída

La interfaz es optimista: marcar "hecho" se pinta al instante en local y se
manda por detrás. Si falla la red, la operación queda en una **cola** en
`localStorage` y se reintenta en el siguiente sondeo. La app sigue funcionando
con lo último que tenga; nunca se pierde lo que el cocinero marcó.

## 5. Almacenamiento (propiedades del script)

Cada propiedad admite 9 KB, así que se separa lo que cambia mucho de lo que no:

| Clave | Contenido | Cada cuánto cambia |
| --- | --- | --- |
| `j:<code>:meta` | `{createdAt, at, rev, planRev, plannerVersion, dataStamp, inputs}` | al crear o recalcular |
| `j:<code>:plan:n` | número de trozos del plan | al crear o recalcular |
| `j:<code>:plan:<i>` | trozo `i` del plan (8000 caracteres) | al crear o recalcular |
| `j:<code>:live` | lo que cambia a menudo (ver abajo) | cada "hecho", cada tarea cogida, cada temporizador |
| `j:index` | `{code: at}` de las jornadas vivas | en cada escritura |

Se sondea **solo `live`**. Las jornadas sin tocar en 36 h se borran en la
siguiente escritura.

### El plan guardado

```json
{ "v": 1, "plannerVersion": 1, "start": 480, "end": 1043, "days": 1,
  "persons": 3, "busy": [120, 118, 95],
  "tasks": [["B05|3|12", 480, 495, 0, null, "Thermomix grande"]],
  "spans": [["TMX_G", "Thermomix grande", "B05", 480, 530]] }
```

`tasks` es `[clave, start, end, person, watcher, aparato]`. El **texto del paso
no se guarda**: cada dispositivo lo saca de sus recetas por la clave
`CÓDIGO|bloque|paso`. Si a un dispositivo le falta ese paso (datos más viejos),
lo pinta igual con la hora y el aparato y avisa de que debe actualizar.

`dataStamp` es el `sourceUpdatedAt` de las recetas con las que se calculó el
plan: si no coincide con el del dispositivo, se avisa (no se bloquea).

### `live`, lo que se sondea

```json
{ "rev": 42, "at": 1763000000000,
  "done": { "croquetas-de-jamon": ["3|12:2", "4|1:0"] },
  "owner": { "croquetas-de-jamon": ["3|12:2"] },
  "timers": [{ "i": "2-7", "k": "arroz|2|3", "l": "Arroz", "e": 1763000600000, "b": 2 }],
  "people": [{ "s": 0, "n": "Leo", "a": 1763000000000 }] }
```

`done` y `owner` van **agrupados por receta**: `"bloque|paso:slot"`. Con la
clave entera en cada entrada (`receta|bloque|paso:slot`) una jornada de 245
pasos marcados ocupaba 9,8 KB y no cabía en una propiedad; agrupada ocupa
3,3 KB (lo mide `test/jornada.test.js`). Por eso tampoco se guarda la hora de
cada "hecho": solo quién lo hizo.

## 6. Protocolo

Lecturas por `GET` (sin preflight, como la carga de recetas):

- `?jornada=CODE` → `{ok, now, code, rev, planRev, meta, plan, live}`
- `?jornada=CODE&rev=N` → si `N` es el `rev` actual, `{ok, now, rev, unchanged: true}`;
  si no, `{ok, now, rev, live}`. Es el sondeo: ~200 bytes cuando no hay cambios.
- `&plan=1` añade el plan (cuando cambia el `planRev`).

Escrituras por `POST` con cuerpo JSON y `Content-Type: text/plain` (petición
simple: tampoco hay preflight). Todas toman `LockService.getScriptLock()`.

| `action` | Cuerpo | Respuesta |
| --- | --- | --- |
| `create` | `{code?, name, inputs, plan, dataStamp, plannerVersion, replace?}` | `{ok, code, rev, planRev, slot, now}` |
| `join` | `{code, name, slot?}` | `{ok, code, rev, planRev, meta, plan, live, slot, now}` |
| `patch` | `{code, slot, ops: [...]}` | `{ok, rev, live, now, rejected: [...]}` |
| `close` | `{code}` | `{ok}` |

Operaciones de `patch`:

| `op` | Campos | Efecto |
| --- | --- | --- |
| `done` | `key`, `slot` | marca hecho (última escritura gana; se rechaza pasadas 600 tareas) |
| `undone` | `key` | quita el hecho |
| `claim` | `key`, `slot` | coge la tarea; si ya tiene dueño, **rechazada** |
| `release` | `key` | la suelta (solo su dueño) |
| `timer` | `id`, `key`, `label`, `seconds` | arranca un temporizador compartido; el servidor calcula `endsAt` con **su** reloj |
| `timerStop` | `id` | lo para |
| `ping` | `slot`, `name` | dice que este dispositivo sigue ahí |

Una tablet que no da señales en 6 h deja su puesto libre, para que quien entre
después no se encuentre todos los sitios cogidos.

Un `patch` con varias operaciones se aplica entero bajo un mismo candado y
sube el `rev` una sola vez.

## 7. Seguridad

La aplicación web se despliega con acceso "cualquier usuario": **el código de
jornada es la única barrera**. No se guarda nada sensible: claves de paso,
horas y números de slot. Las jornadas caducan a las 36 h.

## 8. Qué se ve en la app

- **Pantalla de elegir recetas**: caja para entrar con un código. Una tablet que
  llega a media mañana solo teclea el código: no elige nada.
- **Pantalla del plan**: panel verde con el código, tu nombre, qué persona eres
  y el estado (`Al día` / `N cambios por enviar` / `Sin conexión`). Tu franja
  sale marcada y con "· tú".
- **En cada paso**: quién lo marcó ("hecho por Marta"), quién lo ha cogido, y
  el botón **La cojo yo** en los pasos que el plan le dio a otra persona.
- **Temporizadores**: los tuyos se manejan como siempre; los de los demás se
  ven con su cuenta atrás (en hora de servidor) y no se tocan.
- Si cambias la selección estando en una jornada, aparece **recalcular el plan
  para todos**, que sube un `planRev` nuevo y lo recargan las demás tablets.

## 9. Sin código de jornada

La app funciona exactamente como hasta ahora, en modo de un solo dispositivo:
el plan se calcula en local y las tareas hechas se guardan en `localStorage`.
Compartir es un paso que hay que dar a mano.
