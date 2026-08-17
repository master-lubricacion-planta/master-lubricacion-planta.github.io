# -*- coding: utf-8 -*-
"""Barrido estatico de los 104 HTML de pautas: estructura, correcciones de
fidelidad aplicadas, y coherencia colgroup vs cellmap."""
import json, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')
SCRATCH = r'C:\Users\alexa\AppData\Local\Temp\claude\C--Users-alexa-Desktop-Claude-code-Ses-1\ee5231c6-7fdd-4492-a68b-a3a70ea480d3\scratchpad'
PWA = r'C:\Users\alexa\Desktop\Claude code Ses 1\lubricacion-pwa'
pautas = json.load(open(os.path.join(SCRATCH, 'pautas.json'), encoding='utf-8'))
problemas = 0
resumen = {'con_figura': 0, 'sin_figura': [], 'anchos': {}}
for p in pautas:
    pid = p['id']
    f = os.path.join(PWA, 'pautas-html', pid + '.html')
    issues = []
    if not os.path.exists(f):
        print(pid, '-> SIN HTML'); problemas += 1; continue
    html = open(f, encoding='utf-8').read()
    if 'table-layout: fixed' not in html: issues.append('sin table-layout fixed')
    if 'font-size: 10.0pt' not in html and 'font-size: 11.0pt' not in html: issues.append('fuentes no en pt')
    if 'Calibri' not in html: issues.append('sin Calibri')
    # ancho de tabla declarado == suma del colgroup
    cols = [int(w) for w in re.findall(r'<col\s+style="width: (\d+)px">', html)]
    m = re.search(r'<table[^>]*width:(\d+)px', html)
    if not cols or not m: issues.append('colgroup/width no encontrados')
    elif abs(sum(cols) + 4 - int(m.group(1))) > 1: issues.append(f'ancho tabla {m.group(1)} != suma cols {sum(cols)+4}')
    # celdas del cellmap presentes en el HTML
    cm = p.get('cellmap') or {}
    sheet = p.get('sheet', '')
    faltan = []
    refs = [cm.get('ot')] + (cm.get('tecnicos') or [])[:2]
    for a in (cm.get('activities') or []):
        refs.append(a.get('valor')); refs.append(a.get('estado') or a.get('b'))
    for r in refs:
        if r and f'id="{sheet}!{r}"' not in html:
            faltan.append(r)
    if faltan: issues.append('celdas cellmap ausentes en HTML: ' + ','.join(faltan[:5]))
    # figura
    n_imgs = html.count('<img')
    if n_imgs >= 2: resumen['con_figura'] += 1
    else: resumen['sin_figura'].append(pid)
    resumen['anchos'][sum(cols) if cols else 0] = resumen['anchos'].get(sum(cols) if cols else 0, 0) + 1
    if issues:
        problemas += 1
        print(pid, '->', '; '.join(issues))
print()
print('problemas:', problemas, 'de', len(pautas))
print('con figura de equipo:', resumen['con_figura'], '| sin figura:', len(resumen['sin_figura']), resumen['sin_figura'][:8])
print('anchos de tabla distintos:', dict(sorted(resumen['anchos'].items())))
