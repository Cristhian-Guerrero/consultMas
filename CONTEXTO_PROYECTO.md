# PROYECTO: consultMas V4.13

**Cliente:** A.S. Contadores & Asesores SAS — Pasto, Colombia
**Repo:** https://github.com/Cristhian-Guerrero/consultMas
**Rama activa:** `refactor/modular-structure`
**Última versión pusheada:** V4.13

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

## TECNOPOS (sipos.com.co) — acelerador interno para Express, desde V4.9

Fuente NO oficial de un tercero (TECNOPOS, software POS), endpoint
`GET https://sipos.com.co/api_rut.php?nit={nit}`, HTTP puro vía `requests`
(sin navegador), header `Referer: https://sipos.com.co/consultarut` como
único requisito. ~0.4-0.7s por consulta vs ~2.5-4.5s de DIAN Express.

**Uso: SOLO para empresas confirmadas.** `consultar_tecnopos()` en
`core/scraper.py` llama a `es_empresa(razon_social)`; si no confirma
empresa, retorna `None` de inmediato y el coordinador `consultar_nit()`
cae a `consultar_nit_basica()` (DIAN real). Nunca se expone al usuario de
dónde vino el dato — mismo Excel, mismos campos.

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
`'retry'`** (timeout en el primer intento, ambiguo — podría resolverse
solo con reintentar). Esto es deliberado: no reemplazar un caso que
todavía podría salir bien por DIAN con un dato sin confirmar. Efecto
secundario conocido y aceptado: si DIAN falla por red (no por "no
encontrado" real) en los 3 intentos, el fallback de TECNOPOS no se activa
porque el coordinador solo consulta TECNOPOS en `attempt == 1` — caso raro,
documentado, no resuelto (requeriría rediseñar el mecanismo de reintentos).

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

## Columnas del Excel (modo Express, desde V4.11)

`NIT, DV, Primer Apellido, Segundo Apellido, Primer Nombre, Otros Nombres,
Razón Social, Email, Fecha Consulta, Estado Consulta, Tipo de Consulta,
Observaciones`

`Email` solo viene de TECNOPOS (nunca de DIAN) — "-" cuando el NIT se
resolvió por DIAN. `Dirección/Ciudad/Actividad` se agregaron en V4.10 y se
quitaron en V4.11: en la práctica casi siempre salían "-" (cuando TECNOPOS
las trae, viene sin `dv`, y esa respuesta se descarta entera — no se
publica un NIT sin DV confirmado — así que caía a DIAN, que tampoco las
tiene), así que no aportaban valor real y se simplificó el Excel.

RUT Detallado no cambió: mismas columnas de siempre, sin Email.

## Flujo de trabajo

1. `source venv/bin/activate`
2. Cambios en los módulos correspondientes
3. Probar con `python main.py` (NO `app.py`, no existe en esta rama)
4. Compilar-check + regresión offline antes de commit:
   ```bash
   python -m py_compile core/scraper.py ui/app.py core/excel.py
   ```
5. Al terminar un cambio: **actualizar VERSION en `ui/app.py`,
   `version_info.txt`, y este archivo** — mantenerlos sincronizados en el
   mismo commit.
6. `git add <archivos> && git commit -m "..." && git push origin refactor/modular-structure`
7. GitHub Actions compila el .exe automáticamente (`gh run watch <id>
   --repo Cristhian-Guerrero/consultMas --exit-status` para seguirlo)
8. Descargar .exe desde Actions → Artifacts

## Historial de versiones

- **V4.7** — Toggle duplicados, fix Excel vacío, corrección mapeo nombres (rama `main`)
- **V4.8** — Refactor arquitectura modular, caché por NIT, fix numpy en CI
- **V4.9** — TECNOPOS como acelerador interno para Express (con bug: intentaba separar nombres de personas también)
- **V4.10** — Fix crítico: TECNOPOS solo para empresas (orden de palabras de personas naturales no es confiable); fix `es_empresa()` (LIMITADA, SOCIEDAD, ESP, siglas sueltas); columnas Email/Dirección/Ciudad/Actividad en Excel Express
- **V4.11** — Simplificación: se quitan Dirección/Ciudad/Actividad del Excel (casi siempre "-", no aportaban); se mantiene Email
- **V4.12** — Recupera Email de TECNOPOS para personas naturales, fusionado con el nombre real de DIAN (`_merge_tecnopos_email_con_dian`) — antes se descartaba todo, incluido el email, solo por la ambigüedad del nombre
- **V4.12.1** — Reintentos automáticos en TECNOPOS (máx 3, solo ante fallas de red, backoff 0.5s/1.0s)
- **V4.13** — TECNOPOS como último recurso cuando DIAN da error definitivo (no timeout) — razón social completa sin partir, marcada "sin confirmar" en Observaciones

## Pendiente / conocido sin resolver

- Selector de RUT Detallado ausente en la UI (ver arriba) — confirmar con Betto.
- "DROGUERIA DAGUA"-type: razones sociales sin sufijo reconocible, pérdida de velocidad (no de correctitud).
- Testing en Windows real de V4.10 (Email/Dirección/Ciudad/Actividad + fix empresas) — pendiente confirmación de Betto.
