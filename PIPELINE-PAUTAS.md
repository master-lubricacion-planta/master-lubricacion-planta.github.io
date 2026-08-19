# Pipeline de pautas: Máximo/SharePoint → App (para ejecutar con Sonnet)

Objetivo: incorporar las pautas de inspección de todas las OT del plan semanal a
https://master-lubricacion-planta.github.io/pautas.html

## Descubrimientos clave (ya probados — NO re-investigar)

**⚠️ MÉTODO CORREGIDO (semana 33, Domingo): la búsqueda por TAG en SharePoint NO es confiable.**
Se probó primero el método "buscar por TAG en SharePoint" (puntos 2-3 antiguos, abajo) y en la
verificación posterior contra Máximo se encontró que **32 de 40 pautas asignadas así estaban
mal** (traía el documento de la disciplina/frecuencia equivocada) y que **5 OT que parecían LUB
por su descripción o nombre de archivo en realidad son Tipo=INSP** (inspección mecánica, no
lubricación) — Máximo es la única fuente de verdad fiable. Usar SIEMPRE el método de la sección
"Verificación por Máximo" de abajo como método PRINCIPAL, no como paso opcional.

1. **Las pautas viven adjuntas a cada OT en Máximo** (Attachments), y también existen como copia
   en la biblioteca SharePoint `https://teckresources.sharepoint.com/sites/QB2OP/Gerencia Operaciones
   Integradas/3. CF-Documentos Controlados Operacionales - Técnicos/` con nombres
   `QB2-XXXX-IOC4-PLN-<n>.xlsx` (1.759 archivos) — SharePoint solo sirve para DESCARGAR el archivo
   una vez que Máximo dijo cuál es el correcto, no para decidir cuál es.

2. **Verificación por Máximo (MÉTODO PRINCIPAL — hacer para cada OT, no es opcional)**:
   App "Seguimiento de órdenes de trabajo (TCK)" en `teck.maximo.com`.
   - Filtrar exacto por lote de OT: en el campo "Orden de trabajo" de Vista de lista, escribir
     `=OT1,=OT2,=OT3,...` (con el signo `=` antes de CADA número, separados por coma) y Enter.
     Sin el `=` el filtro hace "contains" y trae basura (registros ACT... no relacionados).
   - Para cada OT: abrir el registro (click en el número) → leer **"Tipo de Orden de trabajo"**
     (panel derecho). **Solo Tipo=LUB es lubricación** — excluir INSP, MECH, PDM genérico, etc.
     aunque la descripción de la OT o el nombre del archivo adjunto contengan la palabra "LUB".
   - Click en el link "Attachments" (ícono de clip, arriba del panel derecho) → popup
     "View Attachments" → columna "Document" trae el nombre real `QB2-...-PLN-<n>.xlsx`.
     `get_page_text` lee el popup sin necesidad de screenshot (más rápido).
   - Los doclinks NO están expuestos por OSLC/API (probado: 500/400/vacío, y `fetch()` con cookies
     de sesión es bloqueado por política del navegador) — hay que leer la UI.

3. **Descargar el PLN correcto desde SharePoint SIN bloqueo de Chrome**: navegar la pestaña a
   `https://teckresources.sharepoint.com/sites/QB2OP/_layouts/15/download.aspx?SourceUrl=<ruta URL-encoded>`
   (una navegación por archivo → cae en `C:\Users\alexa\Downloads`). Para varios archivos a la vez,
   abrir una pestaña nueva por archivo y navegar todas en paralelo (bypasea el límite de Chrome de
   "varias descargas automáticas"; navegar secuencial en la misma pestaña también funciona pero es
   más lento).

4. **Delegar el trabajo mecánico de OT-por-OT a un agente en background** (Agent tool): dado el
   volumen (decenas de OT por día), conviene lanzar un agente por día que recorra la lista con el
   método del punto 2 y reporte una tabla {ot, tipo, estado, archivo(s)}. Indicarle que abra su
   PROPIA pestaña nueva en el mismo navegador (la sesión ya está autenticada y se comparte entre
   pestañas) si se van a correr agentes en paralelo, para no pisarse unos a otros navegando la
   misma pestaña.

## Archivos del pipeline (scratchpad de la sesión original, copiar si se pierde)

- `plan_w33.json` — plan semanal 33 completo (7 días, 259 OT con tag/desc/esp y pauta asociada o null).
- `pautas.json` — pautas ya extraídas (13). El extractor lee cualquier pauta xlsx:
  título (r2,c29), equipo (r6,c15), área (r6,c1), ubicación técnica (r8,c1),
  actividades desde r21: n° (c1), grupo (fila sin número), descripción (c10), límites (c26).
  Termina en la fila "Repuestos".
- `build.py` — build completo: inyecta datos en `app_template.html` y `pautas_template.html`,
  genera `index.html` y `pautas.html` en `lubricacion-pwa`, sube versión del service worker.
- Imágenes de equipo: mayor imagen embebida del xlsx → `pautas-img/<PLN>.jpg` (máx 560px, JPEG q72).

## Procedimiento por lote (repetir por día del plan) — MÉTODO ACTUAL (vía Máximo)

1. Tomar los números de OT sin pauta del día (`plan_w33.json`, pauta=null).
2. Delegar a un Agent (subagente, uno por día, SIEMPRE SECUENCIAL — nunca en paralelo, ver
   advertencia de sesión compartida arriba) que recorra Máximo con el filtro `=OT1,=OT2,...` y
   reporte {ot, tipo_orden_trabajo, estado, archivo(s)_adjunto(s)} de cada una.
3. Con el resultado: descartar toda OT con Tipo ≠ LUB (aunque la descripción o el nombre del
   archivo digan "lubricación"). De las que sí son LUB, listar los PLN nuevos (no presentes aún
   en `pautas.json`).
4. Descargar los PLN nuevos por `download.aspx` (varias pestañas en paralelo está bien para esto,
   es SharePoint, no Máximo — no hay riesgo de sesión compartida).
5. Correr `extract_pauta.py` sobre cada uno → agrega a `pautas.json` + extrae imagen.
6. Mapear OT→PLN en `plan_w33.json` con el diccionario verificado (ya no hace falta comparar
   título vs descripción — el archivo adjunto en Máximo YA es la asociación correcta).
7. `python build.py` → copiar `pautas.json`/`plan_w33.json` a `_src/` → `git add -A && git commit
   && git push` en `lubricacion-pwa`.
8. Verificar en https://master-lubricacion-planta.github.io/pautas.html (esperar ~1 min el deploy).

Casos especiales encontrados: algún PLN vive en otra biblioteca/disciplina (`QB2-1400-IOC4-PLN-…`,
`QB2-0300-MTN2-PLN-…` en vez de `QB2-0300-IOC4-PLN-…`) — si el download.aspx da error, buscar el
Path exacto por SharePoint search antes de reintentar. Algunas OT LUB genuinas no tienen ningún
xlsx adjunto (solo foto, o adjunto en otra biblioteca no ubicada) — quedan sin pauta digital,
está bien, la app las muestra como "sin pauta digital" sin romperse.

## PDF con formato idéntico al de Máximo (feature "Descargar PDF")

`build_pauta_html.py` (scratchpad + `_src/`) genera por cada pauta un
`pautas-html/<PLN>.html` (réplica visual del xlsx vía xlsx2html) + un `cellmap` en
`pautas.json` (coordenadas de celda donde inyectar OT, fechas, técnicos 1-2, y
valor/estado por actividad). En runtime (`pautas.html`) se rellena el HTML, se rasteriza
con html2canvas y se pagina con jsPDF en hojas CARTA apaisadas, cortando siempre en
bordes de fila. NO usar html2pdf.js (falla silenciosamente devolviendo canvas de altura 0).

Correcciones de fidelidad que aplica el post-procesado (lecciones aprendidas, no
re-descubrir): (1) emular `wrap_text` de Excel — celdas sin ajustar texto van con
nowrap y se derraman como en Excel, las ajustadas usan pre-wrap respetando saltos
manuales; (2) anchos de columna con la FÓRMULA REAL de Excel `px = chars*7 + 5`
(MDW Calibri 11) — la de xlsx2html (~chars*9.6) deforma todas las distancias — y
OJO: `ws.column_dimensions` define RANGOS (dim.min..dim.max), hay que expandirlos o
se pierden la mayoría de los anchos; `table-layout:fixed` + ancho de tabla = suma
exacta; (3) tamaños de letra en PUNTOS (`font-size: Npt`) — xlsx2html escribe el
número de pt como px dejando el texto ~25% más chico; (4) rellenos de tema con tinte
resueltos a mano (gris #D9D9D9 = tema 0 con tinte −0.15); (5) Calibri global;
(6) columnas OCULTAS: xlsx2html las omite del colgroup pero emite sus <td> → hay que
ocultarlos y recortar cols sobrantes, si no aparece una "columna fantasma";
(7) las imágenes ancladas dentro de celdas fusionadas (sección FIGURAS) se descartan
por xlsx2html → se inyectan a mano en la celda maestra aplicando el recorte srcRect
de Excel (la imagen cruda puede ser una captura de pantalla completa); (8) la fila
B/M/NA del encabezado puede estar 1 o 2 filas bajo la fila "N°"; (9) el texto
enriquecido dentro de celdas (negrita/cursiva de "Comentarios Condicionales") se
recupera con `load_workbook(rich_text=True)` y se reinyecta como <b>/<i>; (10) xlsx2html
además renderiza SU PROPIA copia de la imagen (posicionada en absoluto, sin recorte) dentro
del mismo td → `inject_figures` la elimina o la figura sale duplicada y desborda la tabla
(pasaba en PLN-875/878/881/883/1063/970). El logo del encabezado (200×53, fila<5) sí es
legítimo y también viene en absoluto — no tocarlo.

Paginación del PDF (runtime, `descargarPdf`): el corte de página toma el ÚLTIMO borde de
fila que quepa (avance mínimo 40px DOM); si una fila sola es más alta que la página
(figuras gigantes) se corta duro igual que la impresora de Máximo. ANTES de medir
`rowTops`/`scrollHeight` hay que esperar la carga de las imágenes y `document.fonts.ready` +
`setTimeout` — y NO usar `img.decode()` ni `requestAnimationFrame`, que en pestañas en
segundo plano quedan pendientes PARA SIEMPRE (colgaba el botón en "Preparando…").

La imagen del equipo del FORMULARIO (`pautas-img/<PLN>.jpg`, la genera `extract_pauta.py`)
también aplica el recorte srcRect — sin él salía la captura de pantalla completa del
planificador. Se sirve con `?v=<versión SW>` como cache-busting (el SW usa
`ignoreSearch:true` al hacer match offline, así que no rompe).

Ojo desarrollo local: usar `_src/serve_dev.py` (gzip) como servidor — el antivirus retiene y TRUNCA en ~510KB las respuestas grandes sin comprimir de
cada archivo nuevo/modificado servido por HTTP en localhost (después va instantáneo). No
es bug de la app; en producción (HTTPS GitHub Pages) no ocurre. Si el deploy de Pages
falla con 429 "Too Many Requests" es transitorio de GitHub: `gh run rerun <id>`.

Regla de producto: las FOTOS adjuntas por el técnico NO van en ningún PDF ni informe
impreso — quedan solo guardadas en la app (IndexedDB del teléfono).

QA: `qa_cellmaps.py` valida que las 104 pautas tengan cellmap completo (OT, técnicos,
valor y estado por actividad, HTML existente). `qa_htmls.py` valida estáticamente los
HTML generados (layout fijo, pt, Calibri, colgroup = ancho de tabla, celdas del cellmap
presentes, cantidad de figuras). `build.py` valida además que toda pauta referenciada por
el plan exista y tenga HTML — falla fuerte en vez de publicar roto. Correr los tres tras
regenerar.

## Estado actual (semana 33 — 2026-08-15)

Las 7 jornadas de la semana 33 fueron verificadas contra Máximo (259 OT totales):

| Día        | Con pauta LUB | Total OT |
|------------|---------------|----------|
| Sábado     | 23            | 23       |
| Domingo    | 40            | 45       |
| Jueves     | 33            | 48       |
| Viernes    | 55            | 60       |
| Lunes      | 40            | 51       |
| Martes     | 22            | 31       |
| Miércoles  | 0             | 1        |
| **Total**  | **213**       | **259**  |

Pautas digitalizadas en biblioteca: 103 (`pautas.json`). Las OT sin pauta son: (a) genuinamente
Tipo≠LUB (mecánicas/inspección, descartadas correctamente), o (b) Tipo=LUB pero sin xlsx
localizable (adjunto solo foto, o en biblioteca SharePoint distinta a la habitual — casos aislados,
no bloquean el resto).

Cada semana nueva: pedir el Excel del programa al usuario, extraer con el patrón de
`plan_w33.json` (hoja CON_LUB_YxxWxx, días JUEVES→MIÉRCOLES), y repetir el procedimiento por
lote de arriba — pero como la biblioteca de pautas ya tiene 103 documentos, muchas OT de semanas
futuras (mismo TAG, misma pauta recurrente) deberían resolverse sin tener que descargar nada
nuevo, solo verificando en Máximo qué PLN está adjunto.
