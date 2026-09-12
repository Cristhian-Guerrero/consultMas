# PROYECTO: consultMas V4.14.1

**Cliente:** A.S. Contadores & Asesores SAS — Pasto, Colombia
**Repo:** https://github.com/Cristhian-Guerrero/consultMas
**Rama activa:** `develop` (Git Flow — ver abajo)
**Última versión pusheada:** V4.14.1 (tag `v4.14.1`, release publicado)

> Este archivo es la fuente de verdad del proyecto — arquitectura, decisiones
> de diseño, gotchas y estado real. Léelo primero al retomar sesión. La
> memoria de Claude (fuera de este repo) guarda un resumen y apunta aquí.

## Estado actual

App de escritorio (Python + Tkinter) que consulta NITs/cédulas masivamente
contra el portal de la DIAN y genera reportes Excel. Desde V4.9 tiene un
acelerador interno vía TECNOPOS (sipos.com.co) para el modo Express.

## Estructura de archivos

```
main.py          ← punto de entrada (logging.basicConfig aquí)
config.py        ← URLs, selectores, timeouts, mensajes rotativos
core/
  browser.py     ← pool de navegadores Chromium (máx 4, limpieza cada 10 requests)
  scraper.py     ← scraping DIAN (Express + RUT Detallado) + TECNOPOS
  excel.py       ← formato y estilos del Excel generado
ui/
  app.py         ← GUI Tkinter — ConsultaRUTApp (_procesar es el loop principal)
version_info.txt ← metadatos del .exe para Windows (mantener en sync con VERSION)
```

## Modos de consulta

- **Express** (`tipo="basica"`): único modo seleccionable en la UI hoy.
- **RUT Detallado** (`tipo="rut_detallado"`): trae `Estado del Registro`
  (activo/inactivo), dato que ni DIAN Express ni TECNOPOS traen nunca.
  **⚠️ El código backend sigue completo y funcional, pero no hay ningún
  radio button ni control en `ui/app.py` para seleccionarlo** (se perdió en
  algún punto del refactor modular, no está documentado cuándo ni por qué).
  Confirmado revisando el archivo completo: solo existe un
  `ttk.Radiobutton` (Express). Sin resolver — pendiente confirmar con el
  cliente si todavía lo necesita.

## Arquitectura V4.14.0 — DIAN primero, TECNOPOS como fallback (cambio de dirección)

Hasta V4.13.1 el coordinador `consultar_nit()` consultaba **TECNOPOS
primero** como acelerador (para empresas confirmadas), y caía a DIAN si
TECNOPOS no servía. Desde V4.14.0 se **invirtió el orden**: DIAN es
siempre la primera consulta (fuente oficial), y TECNOPOS pasa a ser
exclusivamente (a) un complemento de email para personas cuando DIAN ya
tuvo éxito, y (b) el fallback de última instancia cuando DIAN no tiene
datos reales (No Inscrito o error). Esto sacrifica velocidad (Express
vuelve a depender siempre del navegador/DIAN, ~2.5-4.5s) a cambio de que
la fuente oficial sea la primera palabra en todos los casos.

**El fallback de V4.13/V4.13.1 ("TECNOPOS ayuda cuando DIAN no
encuentra") se preservó intacto** dentro de esta nueva arquitectura — sigue
siendo la razón de ser del proyecto según el cliente. Lo nuevo en V4.14.0:

- **Validación de DV antes de aceptar TECNOPOS** (con `calcular_dv()` — el
  mismo algoritmo oficial de 15 posiciones que ya usaba
  `consultar_nit_basica()`, no uno nuevo): si el DV que trae TECNOPOS no
  coincide con el DV calculado del NIT, se descarta y se conserva el
  resultado de DIAN. Filtra respuestas de TECNOPOS con datos inconsistentes
  antes de que lleguen al Excel. **Se mantiene en V4.14.1.**
- El campo interno `_email_fuente` (no visible en Excel) marca cuándo el
  email de una persona vino de TECNOPOS en vez de DIAN — mismo mecanismo de
  V4.12.

**V4.14.1 revirtió la única pieza de auditoría visible que había agregado
V4.14.0:** la columna `Fuente` (DIAN/TECNOPOS) del Excel y el campo interno
`_fuente` que la alimentaba se quitaron por completo — vuelta al criterio
de V4.9-V4.13.1 de "nunca exponer al usuario de dónde vino el dato". De
paso, V4.14.1 también quitó el texto `"Dato de TECNOPOS sin confirmar en
DIAN — verificar manualmente"` que `_fallback_tecnopos_sin_confirmar()`
venía escribiendo en Observaciones desde V4.13 — con esto, **ya no queda
ninguna señal visible en el Excel de que una fila vino sin confirmar de
TECNOPOS** (ni columna, ni Observaciones); cuando ese fallback aplica, la
fila cae al texto genérico "Consulta Express exitosa" en Observaciones,
igual que un dato 100% de DIAN. Fue una decisión explícita del cliente
("retornar al estado anterior"), documentada aquí porque revierte una
protección de trazabilidad que sí existía desde V4.13.

Detectar "DIAN sin datos reales" (para decidir si vale la pena intentar
TECNOPOS) revisa que `primerApellido`/`razonSocial` no sean el placeholder
`"-"` que usa `consultar_nit_basica()` para "No Inscrito" — un `status`
`'success'` de DIAN no basta por sí solo, porque "No Inscrito" también es
`'success'` (con campos en `"-"`), no `'error'`.

## TECNOPOS (sipos.com.co) — fallback y complemento de email desde V4.14

Fuente NO oficial de un tercero (TECNOPOS, software POS), endpoint
`GET https://sipos.com.co/api_rut.php?nit={nit}`, HTTP puro vía `requests`
(sin navegador), header `Referer: https://sipos.com.co/consultarut` como
único requisito. ~0.4-0.7s por consulta — rápido, pero desde V4.14.0 solo
se usa cuando DIAN ya respondió (persona, para el email) o cuando DIAN no
tiene datos reales (fallback).

**Uso: SOLO para empresas confirmadas.** `consultar_tecnopos()` en
`core/scraper.py` llama a `es_empresa(razon_social)`; si no confirma
empresa, retorna `None` de inmediato. Nunca se expone al usuario de dónde
vino el dato salvo por la columna `Fuente` explícita (ver arriba).

### Email de personas naturales (V4.12) — merge TECNOPOS + DIAN

Desde V4.12, cuando `consultar_tecnopos()` detecta que un NIT NO es
empresa, ya no descarta todo — devuelve `{'email': ..., '_es_persona':
True}` (el email no tiene ambigüedad de orden, a diferencia del nombre). El
coordinador `consultar_nit()` ve ese flag, consulta DIAN normalmente para
obtener el nombre confiable, y `_merge_tecnopos_email_con_dian()` inyecta
el email dentro de `resultado_dian['data']['email']` antes de devolver el
resultado final. `ui/app.py` no necesitó ningún cambio — ya leía
`data.get("email")` de forma genérica. El nombre de personas naturales
sigue viniendo 100% de DIAN, nunca de TECNOPOS.

Limitación conocida y aceptada: el merge solo ocurre en el intento 1 de
`consultar_nit()` (mismo gate que el resto de TECNOPOS). Si DIAN necesita
reintento (`attempt=2` o `3`), el email ya obtenido en el intento 1 no se
vuelve a fusionar — caso raro (solo cuando DIAN da timeout), y solo afecta
el email (dato complementario), nunca el nombre.

### TECNOPOS como apoyo cuando DIAN no encuentra el NIT (V4.13)

Razón de ser del proyecto según el cliente: "esta página es de ayuda a la
DIAN — si no lo encuentra en la DIAN, esta página sí". Desde V4.13, cuando
DIAN responde con **error definitivo** (no un timeout ambiguo — ver abajo)
para un NIT que TECNOPOS sí tenía, se usa TECNOPOS como último recurso vía
`_fallback_tecnopos_sin_confirmar()`: la razón social se muestra **completa,
sin partir** en apellido/nombre (el orden sigue sin ser confiable — ver
V4.10 abajo), con `Observaciones = "Dato de TECNOPOS sin confirmar en DIAN
— verificar manualmente"` para que quede explícito que no pasó por DIAN.

**Solo aplica cuando `resultado_dian['status'] == 'error'`, NO cuando es
`'retry'`** (timeout ambiguo — podría resolverse solo con reintentar). Esto
es deliberado: no reemplazar un caso que todavía podría salir bien por DIAN
con un dato sin confirmar.

**Fix V4.13.1 — bug real encontrado con NITs `123456789`/`987654321`:**
`consultar_nit()` originalmente solo consultaba TECNOPOS en `attempt == 1`.
Pero DIAN da `'retry'` (no `'error'`) en el intento 1 casi siempre — el
`'error'` definitivo recién aparece en el intento 2 o 3, momento en el que
el coordinador ya no volvía a mirar TECNOPOS, y el dato que ya se tenía se
perdía para siempre. Confirmado en vivo: `_procesar()` terminaba mostrando
"No Inscrito" para NITs que TECNOPOS sí tenía. Fix: TECNOPOS se consulta en
**cada intento**, no solo el primero — es rápido (~0.5-0.7s) y los
reintentos ya son el camino lento de por sí, así que repetirla ahí no tiene
costo real. Validado end-to-end simulando la cascada real de `_procesar()`
(intento 1 retry → intento 2 con fallback aplicado correctamente).

### Reintentos en TECNOPOS (V4.12.1)

`_fetch_tecnopos()` reintenta hasta 3 veces, pero **solo ante fallas reales
de red** (`Timeout`, `ConnectionError`) — nunca ante un `success:false`
limpio (un NIT no encontrado no cambia al reintentar, verificado con
pruebas reales). Backoff lineal 0.5s/1.0s entre intentos.

### Por qué el nombre NO se usa de TECNOPOS para personas naturales (hallazgo crítico, V4.10)

TECNOPOS **no tiene un orden de palabras consistente** en `razon_social`
para personas naturales. Confirmado con 3 NITs reales:
- `79750160` → "HERRERA CARRION WILSON ALBERTO" → apellidos primero
- `52156616` → "NIÑO ESPEJO LINA JIBE" → apellidos primero
- `87069568` → "ALBERTO EDMUNDO SANCHEZ MARTINEZ" → **nombres primero**

Ninguna regla posicional fija sirve para los tres casos a la vez, y no hay
señal en el JSON para distinguir cuál es cuál. Por eso se descartó
completamente intentar separar nombres desde TECNOPOS — `_separar_nombre()`
sigue en el código pero sin usar en este flujo, a propósito. DIAN sigue
siendo la única fuente para personas naturales porque tiene apellidos y
nombres en campos HTML separados de verdad (no un string a adivinar).

### Limitación conocida — detección de empresa vía sufijo

`es_empresa()` (en `core/scraper.py`) detecta empresa por palabra-sufijo:
`SAS, SA, LTDA, LIMITADA, CIA, EU, EIRL, ESE, IPS, FUNDACION, COOPERATIVA,
ASOCIACION, ONG, SOCIEDAD, ESP`. También une siglas sueltas de una letra
("S A" → "SA", confirmado real con NIT 830114921 "COLOMBIA MOVIL S A ESP",
TECNOPOS a veces no manda los puntos).

Una razón social **sin sufijo societario reconocible** (ej. "DROGUERIA
DAGUA", NIT 800140016 — nombre comercial, no forma jurídica) no se detecta
como empresa. **Ya no es riesgo de datos incorrectos** (desde el fix de
"TECNOPOS solo para empresas": si `es_empresa()` no la reconoce, cae a
DIAN igual, que sí sabe internamente si es persona o empresa) — es
solamente una oportunidad de velocidad perdida para esos NITs puntuales.
No tiene fix simple (necesitaría un diccionario de sustantivos de tipo de
negocio, con cobertura nunca 100% y riesgo de falsos positivos).

## Validación hecha (Fase 1, offline, sin tocar producción)

Contra 101 razones sociales reales (5 del histórico de consultMas + 96 del
registro público de Cámara de Comercio, `datos.gov.co` dataset `xpg6-d7rc`):
tasa de clasificación errónea 0.00% tras los fixes de `LIMITADA`/`SOCIEDAD`/
`ESP` + unión de siglas sueltas. Reproducible con
`core.scraper.es_empresa` + `core.scraper._separar_nombre` sin red.

## Quirk importante — DIAN portal (preexistente, no tocado en V4.9/V4.10)

Los IDs del HTML del portal DIAN Express tienen los nombres INVERTIDOS
respecto a las etiquetas visuales: HTML id `primerApellido` → contiene el
NOMBRE, HTML id `primerNombre` → contiene el APELLIDO. Corregido en
`ui/app.py:_procesar()` al mapear a columnas de Excel. Confirmado con
`inspeccion_dian.py` para NIT 79750160 (V4.8) y re-validado en V4.10 contra
el histórico real (ver arriba).

## Columnas del Excel (modo Express, 12 columnas desde V4.14.1 — igual que V4.11-V4.13.1)

`NIT, DV, Primer Apellido, Segundo Apellido, Primer Nombre, Otros Nombres,
Razón Social, Email, Fecha Consulta, Estado Consulta, Tipo de Consulta,
Observaciones`

`Email` casi siempre viene de TECNOPOS (nunca de DIAN, que no expone ese
campo) — "-" cuando no hay dato. `Dirección/Ciudad/Actividad` se agregaron
en V4.10 y se quitaron en V4.11: en la práctica casi siempre salían "-"
(cuando TECNOPOS las trae, viene sin `dv`, y esa respuesta se descarta
entera — no se publica un NIT sin DV confirmado — así que caía a DIAN, que
tampoco las tiene), así que no aportaban valor real y se simplificó el
Excel.

**V4.14.0 había agregado una columna `Fuente` (DIAN/TECNOPOS); V4.14.1 la
quitó** a pedido explícito del cliente — ver arquitectura arriba. El Excel
no tiene ninguna señal (ni columna, ni texto en Observaciones) de qué
motor resolvió cada fila desde V4.14.1.

RUT Detallado no cambió: mismas columnas de siempre, sin Email (ese modo no
pasa por TECNOPOS en absoluto — `consultar_nit()` delega directo a
`consultar_nit_rut_detallado()`).

## Estrategia de branches (Git Flow, desde V4.14.1)

- **`main`** → producción. Solo recibe merges de `develop` cuando hay una
  release lista. Cada release lleva tag `vX.Y.Z` (ver abajo).
- **`develop`** → desarrollo. Base de trabajo normal; recibe merges de
  `feature/*` y `fix/*` vía PR.
- **`feature/*`, `fix/*`** → ramas de corta duración, siempre creadas desde
  `develop`, mergeadas de vuelta a `develop` vía PR.
- `Master` (huérfana, un solo "Initial commit" con un README de 1 línea,
  sin relación de ancestro con `main`) y `refactor/modular-structure` (ya
  fusionada por completo a `develop`/`main`) se **eliminaron** en V4.14.1 —
  ver commit de branching. El default branch del repo en GitHub pasó de
  `Master` a `main`.

## Versioning con tags y Releases (desde V4.14.1)

Tags semánticos `vX.Y.Z` sobre `main`, creados **después** de mergear
`develop → main`:

```bash
git checkout main && git merge --ff-only develop && git push origin main
git tag -a v4.15.0 -m "Release v4.15.0: ..."
git push origin v4.15.0
```

Esto dispara `build-windows.yml`, que reconoce el tag y sube el artifact
como `ConsultaDIAN-v4.15.0` (con el .exe adentro ya renombrado igual, no
`ConsultaDIAN.exe` genérico). Un release de GitHub con el .exe adjunto se
crea a mano tras confirmar el build verde:

```bash
gh run download <run_id> -n ConsultaDIAN-v4.15.0   # o vía gh api si "path traversal" (ver abajo)
gh release create v4.15.0 ConsultaDIAN-v4.15.0.exe --title "v4.15.0" --notes "..."
```

**Gotcha confirmado en V4.14.1:** `gh run download -n <nombre>` puede fallar
con `would result in path traversal` para ciertos nombres de artifact (no
identificada la causa exacta). Workaround que funcionó: `gh api
repos/.../actions/artifacts/<id>/zip > artifact.zip` + `unzip`.

Push normal (sin tag) a `main` o `develop` también compila y sube un
artifact temporal: `ConsultaDIAN-Latest-main` / `ConsultaDIAN-Latest-develop`
— pensado para verificar que compila, no para entregar al cliente. Un
`pull_request` compila igual (CI) pero no publica ningún artifact.

`feature/*`/`fix/*` NO disparan build automático a propósito: el nombre de
artifact usa `github.ref_name` tal cual, que para esas ramas incluye la
`/` del prefijo — un artifact no puede llamarse `ConsultaDIAN-Latest-feature/x`.
Si se necesita en el futuro, hay que sanear el nombre (reemplazar `/` por
`-`) antes de habilitar el trigger.

## Flujo de trabajo (dentro de develop/feature)

1. `git checkout develop && git pull` (o `git checkout -b feature/x develop`
   para algo más grande)
2. `source venv/bin/activate`
3. Cambios en los módulos correspondientes
4. Probar con `python main.py` (NO `app.py`, no existe en este árbol)
5. Compilar-check + regresión offline antes de commit:
   ```bash
   python -m py_compile core/scraper.py ui/app.py core/excel.py
   ```
6. Al terminar un cambio: **actualizar VERSION en `ui/app.py`,
   `version_info.txt`, y este archivo** — mantenerlos sincronizados en el
   mismo commit.
7. `git add <archivos> && git commit -m "..." && git push origin develop`
   (o la rama `feature/*`, luego PR a `develop`)
8. GitHub Actions compila el .exe automáticamente (`gh run watch <id>
   --repo Cristhian-Guerrero/consultMas --exit-status` para seguirlo)
9. Cuando el trabajo en `develop` esté listo para salir a producción:
   mergear a `main` y taguear (ver "Versioning con tags" arriba).

## Historial de versiones

- **V4.7** — Toggle duplicados, fix Excel vacío, corrección mapeo nombres (rama `main`)
- **V4.8** — Refactor arquitectura modular, caché por NIT, fix numpy en CI
- **V4.9** — TECNOPOS como acelerador interno para Express (con bug: intentaba separar nombres de personas también)
- **V4.10** — Fix crítico: TECNOPOS solo para empresas (orden de palabras de personas naturales no es confiable); fix `es_empresa()` (LIMITADA, SOCIEDAD, ESP, siglas sueltas); columnas Email/Dirección/Ciudad/Actividad en Excel Express
- **V4.11** — Simplificación: se quitan Dirección/Ciudad/Actividad del Excel (casi siempre "-", no aportaban); se mantiene Email
- **V4.12** — Recupera Email de TECNOPOS para personas naturales, fusionado con el nombre real de DIAN (`_merge_tecnopos_email_con_dian`) — antes se descartaba todo, incluido el email, solo por la ambigüedad del nombre
- **V4.12.1** — Reintentos automáticos en TECNOPOS (máx 3, solo ante fallas de red, backoff 0.5s/1.0s)
- **V4.13** — TECNOPOS como último recurso cuando DIAN da error definitivo (no timeout) — razón social completa sin partir, marcada "sin confirmar" en Observaciones
- **V4.13.1** — Fix: TECNOPOS se consulta en cada reintento, no solo el 1ro — el 'error' definitivo de DIAN casi siempre llega en el intento 2/3, y el dato de TECNOPOS se perdía si no se volvía a mirar ahí
- **V4.14.0** — Inversión de arquitectura: DIAN primero siempre (fuente oficial), TECNOPOS pasa de acelerador a fallback/complemento de email. Se preserva íntegro el fallback "TECNOPOS ayuda cuando DIAN no encuentra" de V4.13/V4.13.1. Agrega: validación de DV (reutilizando `calcular_dv()` ya existente, algoritmo oficial de 15 posiciones) antes de aceptar cualquier dato de TECNOPOS; columna `Fuente` (DIAN/TECNOPOS) en el Excel Express como auditoría visible; se elimina `_merge_tecnopos_email_con_dian()` (código muerto tras el reordenamiento — su lógica quedó inline en el coordinador)
- **V4.14.1** — Revierte la columna `Fuente` y el campo interno `_fuente` (12 columnas de nuevo, como V4.11-V4.13.1); también quita el texto "Dato de TECNOPOS sin confirmar en DIAN — verificar manualmente" que `_fallback_tecnopos_sin_confirmar()` escribía en Observaciones desde V4.13 — decisión explícita del cliente de no exponer de dónde vino el dato. Todo lo demás de V4.14.0 se mantiene intacto (DIAN primero, fallback TECNOPOS con validación de DV, email fallback). Además: se adopta Git Flow (`main`/`develop`, se elimina `Master` y `refactor/modular-structure`), se crea el primer tag `v4.14.1` con release en GitHub, y `build-windows.yml` pasa a generar artifacts con nombre profesional (`ConsultaDIAN-v4.14.1` en tags, `ConsultaDIAN-Latest-<rama>` en pushes normales, nada en pull requests)

## Pendiente / conocido sin resolver

- Selector de RUT Detallado ausente en la UI (ver arriba) — confirmar con Betto.
- "DROGUERIA DAGUA"-type: razones sociales sin sufijo reconocible, pérdida de velocidad (no de correctitud).
- Testing en Windows real de V4.10 (Email/Dirección/Ciudad/Actividad + fix empresas) — pendiente confirmación de Betto.
- **V4.14.0/V4.14.1 solo se validaron con tests unitarios (mocks, sin red real a DIAN/TECNOPOS)** — falta E2E real: correr Express contra NITs reales (persona DIAN, empresa DIAN, No Inscrito con y sin TECNOPOS) y confirmar visualmente el Excel de 12 columnas en Windows. Se hace vía push a CI (`refactor/modular-structure` → GitHub Actions compila el .exe), no local.
