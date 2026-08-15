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

## Procedimiento por lote (repetir por día del plan)

1. Tomar TAGs de las OT sin pauta del día (`plan_w33.json`, pauta=null).
2. Búsqueda SharePoint por cada TAG → lista de PLN candidatos (dedupe global).
3. Descargar los PLN nuevos (download.aspx, uno por uno) → mover de Downloads a `C:\Users\alexa\Desktop\Pautas\`.
4. Correr el extractor sobre los nuevos → agregar a `pautas.json` + extraer imagen.
5. Asociar OT→PLN (título vs descripción de OT; regla del punto 3 de arriba). Actualizar el mapeo en el script de extracción del plan.
6. `python build.py` → `git add -A && git commit && git push` en `lubricacion-pwa`.
7. Verificar en https://master-lubricacion-planta.github.io/pautas.html (esperar ~1 min el deploy).

## Estado actual

- Semana 33 cargada completa (259 OT). Sábado 15.08: 23/23 con pauta digital. Resto de días: pendiente.
- Pautas digitalizadas: 13 (ver pautas.json). PLN-734 (Chute Bypass Pebbles) ya descargada en Downloads, pendiente de incorporar.
- Cada semana nueva: pedir el Excel del programa al usuario, extraer con el patrón de `plan_w33.json` (hoja CON_LUB_YxxWxx, días JUEVES→MIÉRCOLES), y cruzar TAGs con la biblioteca de pautas ya acumulada.
