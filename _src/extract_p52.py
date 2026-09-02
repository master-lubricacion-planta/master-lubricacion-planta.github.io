# -*- coding: utf-8 -*-
"""Extrae del 'Programa de lubricación 52 semanas 2026.xlsx' (hoja Master LUB.):
- p52_items.json: las 340 'Inspección general de lubricación' de frecuencia Semanal
  [{n, area, tag, eq}]
- p52_prev.json: historial ya digitado en la planilla, por semana:
  {"<w>": {"<n>": {"e": "R|C|RE", "d": 0-6|null, "ot": "", "obs": ""}}}
Reglas de lectura por bloque de semana (7 dias + Observaciones + OT):
  celda de dia con R/C/RE -> estado + dia; numero largo -> OT; otro texto -> obs.
"""
import openpyxl, json, re, sys
sys.stdout.reconfigure(encoding='utf-8')

SRC = r'C:\Users\alexa\Downloads\Programa de lubricación 52 semanas 2026.xlsx'
wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb['Master LUB.']

items = []
filas = []
for r in range(8, 1174):
    desc = str(ws.cell(row=r, column=5).value or '').strip()
    freq = str(ws.cell(row=r, column=9).value or '').strip()
    if desc == 'Inspección general de lubricación' and freq == 'Semanal':
        n = ws.cell(row=r, column=1).value
        items.append({
            'n': int(n),
            'area': str(ws.cell(row=r, column=2).value or '').strip(),
            'tag': str(ws.cell(row=r, column=3).value or '').strip(),
            'eq': str(ws.cell(row=r, column=4).value or '').strip(),
        })
        filas.append((int(n), r))

def blo(k):
    return 10 + (k - 1) * 9

EST = {'R': 'R', 'C': 'C', 'RE': 'RE'}
prev = {}
for n, r in filas:
    for k in range(1, 53):
        b = blo(k)
        e = None; d = None; ot = ''; obs = []
        for i in range(7):
            v = ws.cell(row=r, column=b + i).value
            if v is None: continue
            s = str(v).strip()
            if not s: continue
            up = s.upper()
            if up in EST and e is None:
                e = EST[up]; d = i
            elif re.fullmatch(r'\d{5,9}', s):
                ot = ot or s
            else:
                obs.append(s)
        for col, kind in ((b + 7, 'obs'), (b + 8, 'ot')):
            v = ws.cell(row=r, column=col).value
            if v is None: continue
            s = str(v).strip()
            if not s: continue
            if re.fullmatch(r'\d{5,9}', s):
                ot = ot or s
            elif s.upper() in EST and e is None:
                e = EST[s.upper()]
            else:
                obs.append(s)
        if e or ot or obs:
            ent = {}
            if e: ent['e'] = e
            if d is not None: ent['d'] = d
            if ot: ent['ot'] = ot
            if obs: ent['obs'] = ' · '.join(dict.fromkeys(obs))[:180]
            prev.setdefault(str(k), {})[str(n)] = ent

json.dump(items, open('p52_items.json', 'w', encoding='utf-8'), ensure_ascii=False)
json.dump(prev, open('p52_prev.json', 'w', encoding='utf-8'), ensure_ascii=False)
import os
tot = sum(len(v) for v in prev.values())
print('items:', len(items), '| semanas con datos:', len(prev), '| entradas historicas:', tot)
print('tamanos: items', os.path.getsize('p52_items.json')//1024, 'KB | prev', os.path.getsize('p52_prev.json')//1024, 'KB')
con_r = sum(1 for w in prev.values() for x in w.values() if x.get('e')=='R')
print('con R:', con_r)
