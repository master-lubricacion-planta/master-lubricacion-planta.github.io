# -*- coding: utf-8 -*-
"""Aplica el resultado de Maximo al plan de la semana (generaliza apply_maximo_w35.py).

Uso: python3 apply_maximo.py 37
Entrada: maximo_w<nn>.txt con lineas "OT | TIPO | ARCHIVOS" (los ; separan archivos).
Regla (decision de Nicolas, semana 35): el programa CON_LUB ya es de la cuadrilla de
lubricacion, asi que TODA OT con pauta xlsx adjunta se incluye aunque su Tipo sea
INSP/PM/PDM. Solo quedan fuera las que no tienen xlsx (pdf escaneado / sin adjuntos).
Reporta los PLN que NO estan en la biblioteca (hay que descargarlos)."""
import json, re, sys, os
base = os.path.dirname(os.path.abspath(__file__))
W = sys.argv[1]
plan = json.load(open(os.path.join(base, f'plan_w{W}.json'), encoding='utf-8'))
pautas = json.load(open(os.path.join(base, 'pautas.json'), encoding='utf-8'))
lib = {p['id'] for p in pautas}

res = {}
for ln in open(os.path.join(base, f'maximo_w{W}.txt'), encoding='utf-8'):
    partes = [x.strip() for x in ln.split('|')]
    if len(partes) < 3 or not partes[0].isdigit():
        continue
    ot, tipo, archivos = partes[0], partes[1].upper(), partes[2]
    plns = re.findall(r'QB2-[A-Z0-9]+-[A-Z0-9]+-PLN-\d+', archivos.upper().replace(' ', ''))
    res[ot] = {'tipo': tipo, 'plns': list(dict.fromkeys(plns))}

aplicadas = 0; sin_adj = []; nuevos = set(); multi = []
for d in plan['days']:
    for o in d['ots']:
        r = res.get(o['ot'])
        if not r or o['pauta']:
            continue
        if not r['plns']:
            sin_adj.append((o['ot'], r['tipo'])); continue
        if len(r['plns']) > 1:
            multi.append((o['ot'], r['plns']))
        pln = r['plns'][0]
        o['pauta'] = pln; aplicadas += 1
        if pln not in lib:
            nuevos.add(pln)

json.dump(plan, open(os.path.join(base, f'plan_w{W}.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
tot = sum(len(d['ots']) for d in plan['days'])
conp = sum(1 for d in plan['days'] for o in d['ots'] if o['pauta'])
print(f'semana {W}: resultados leidos {len(res)} | aplicadas {aplicadas} | plan {conp}/{tot} con pauta')
print('sin xlsx adjunto (quedan sin pauta digital):', sin_adj)
if multi: print('OJO multiples PLN (se tomo el primero):', multi)
print('PLN NUEVOS a descargar:', ' '.join(sorted(nuevos)) or 'ninguno')
pend = [(d['nombre'], o['ot']) for d in plan['days'] for o in d['ots'] if o['esp'] == 'LUB' and not o['pauta'] and o['ot'] not in res]
print('aun sin resultado de Maximo:', pend)
