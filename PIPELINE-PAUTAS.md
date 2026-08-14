# Pipeline de pautas: Máximo/SharePoint → App (para ejecutar con Sonnet)

Objetivo: incorporar las pautas de inspección de todas las OT del plan semanal a
https://master-lubricacion-planta.github.io/pautas.html

## Descubrimientos clave (ya probados — NO re-investigar)

1. **Las pautas NO viven en Máximo**: los attachments de las OT son enlaces a la biblioteca SharePoint
   `https://teckresources.sharepoint.com/sites/QB2OP/Gerencia Operaciones Integradas/3. CF-Documentos Controlados Operacionales - Técnicos/`
   con nombres `QB2-XXXX-IOC4-PLN-<n>.xlsx` (1.759 archivos). No hace falta entrar OT por OT a Máximo.

2. **Asociar TAG → pauta con el buscador de SharePoint** (indexa el contenido; el TAG va dentro del archivo).
   En una pestaña de `teckresources.sharepoint.com` (sesión del usuario), por cada TAG del plan:
   ```js
   const q='"<TAG>" path:"https://teckresources.sharepoint.com/sites/QB2OP/Gerencia Operaciones Integradas/3. CF-Documentos Controlados Operacionales - Técnicos" filetype:xlsx';
   fetch("/sites/QB2OP/_api/search/query?querytext='"+encodeURIComponent(q.replace(/'/g,"''"))+"'&selectproperties='Path'&rowlimit=20",{headers:{Accept:'application/json;odata=nometadata'}})
   ```
   Devuelve varias pautas candidatas por TAG (distintas disciplinas/frecuencias).

3. **Elegir la pauta correcta entre candidatas**: descargarlas y comparar el título de la pauta
   (celda fila 2, col 29: ej. "PI M-Agitador TK almac floculante") con la descripción de la OT
   (PI/PM/PDM + S/M + equipo). La especialidad del plan es LUB. En duda, preferir coincidencia
   de prefijo (PI↔PI, PDM↔PDM) y frecuencia (S=semanal, M=mensual).

4. **Descargar SIN bloqueo de Chrome**: navegar la pestaña a
   `https://teckresources.sharepoint.com/sites/QB2OP/_layouts/15/download.aspx?SourceUrl=<ruta URL-encoded>`
   (una navegación por archivo → cae en `C:\Users\alexa\Downloads`). NO usar clicks de descarga masiva (Chrome los bloquea).

5. **Verificación por Máximo (opcional, para casos dudosos)**: en pestaña de `teck.maximo.com`:
   - OT → datos: `fetch('/maximo/oslc/os/mxwodetail?oslc.where=wonum="<OT>"&oslc.select=wonum,jpnum,location,assetnum&lean=1',{headers:{Accept:'application/json'}})` (funciona).
   - Los doclinks NO están expuestos por OSLC (probado: 500/400/vacío). Para confirmar el PLN de una OT hay que abrir la OT en la UI (app "Seguimiento de órdenes de trabajo (TCK)" → buscar OT → Attachments) y leer el nombre del archivo.

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
