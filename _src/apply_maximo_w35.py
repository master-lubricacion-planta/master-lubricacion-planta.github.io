# -*- coding: utf-8 -*-
"""Aplica los resultados de los agentes de Maximo al plan_w35.json.
Entrada: maximo_w35.txt con lineas "OT | TIPO | ARCHIVOS" (los ; separan archivos).
- TIPO != LUB  -> pauta queda null (no es lubricacion)
- TIPO == LUB con adjunto QB2-...-PLN-n.xlsx -> pauta = ese id
Reporta los PLN que NO estan en la biblioteca (hay que descargarlos)."""
import json, re, sys
sys.stdout.reconfigure(encoding='utf-8')

plan = json.load(open('plan_w35.json', encoding='utf-8'))
pautas = json.load(open('pautas.json', encoding='utf-8'))
lib = {p['id'] for p in pautas}

res = {}
for ln in open('maximo_w35.txt', encoding='utf-8'):
    partes = [x.strip() for x in ln.split('|')]
    if len(partes) < 3 or not partes[0].isdigit():
        continue
    ot, tipo, archivos = partes[0], partes[1].upper(), partes[2]
    plns = re.findall(r'QB2-[A-Z0-9]+-[A-Z0-9]+-PLN-\d+', archivos.upper().replace(' ', ''))
    res[ot] = {'tipo': tipo, 'plns': list(dict.fromkeys(plns))}

aplicadas = 0; no_lub = []; sin_adj = []; nuevos = set(); multi = []
for d in plan['days']:
    for o in d['ots']:
        r = res.get(o['ot'])
        if not r or o['pauta']:
            continue
        if r['tipo'] != 'LUB':
            no_lub.append((o['ot'], r['tipo'])); continue
        if not r['plns']:
            sin_adj.append(o['ot']); continue
        if len(r['plns']) > 1:
            multi.append((o['ot'], r['plns']))
        pln = r['plns'][0]
        o['pauta'] = pln; aplicadas += 1
        if pln not in lib:
            nuevos.add(pln)

json.dump(plan, open('plan_w35.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
tot = sum(len(d['ots']) for d in plan['days'])
conp = sum(1 for d in plan['days'] for o in d['ots'] if o['pauta'])
print(f'resultados leidos: {len(res)} | aplicadas: {aplicadas} | plan: {conp}/{tot} con pauta')
print('no-LUB (quedan sin pauta):', no_lub)
print('LUB sin adjunto xlsx:', sin_adj)
if multi: print('OJO multiples PLN (se tomo el primero):', multi)
print('PLN NUEVOS a descargar de SharePoint:', sorted(nuevos))
pendientes = [(d['nombre'], o['ot']) for d in plan['days'] for o in d['ots'] if o['esp']=='LUB' and not o['pauta'] and o['ot'] not in [x[0] for x in no_lub] and o['ot'] not in sin_adj]
print('aun sin resultado de Maximo:', pendientes)
