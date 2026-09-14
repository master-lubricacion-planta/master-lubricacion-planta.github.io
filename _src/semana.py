# -*- coding: utf-8 -*-
"""Carga semanal del programa CON_LUB en UN comando (fase 2, sep-2026).

Uso:
  python3 semana.py 38            # paso 1: parsea el Excel y lista OT para Maximo
  python3 semana.py 38 --aplicar  # paso 2: con maximo_w38.txt listo, aplica, extrae
                                  #         pautas nuevas, genera HTML, QA, build,
                                  #         publica datos y hace push
  python3 semana.py 38 --sin-push # igual que --aplicar pero sin git push

Entradas esperadas:
  ../descargas/PROGRAMA_CON_LUB_Y26W<nn>.xlsx   (o el nombre original en ~/Downloads:
                                                 "PROGRAMA CON_LUB_Y26W<nn>.xlsx" / "Copia de ...")
  maximo_w<nn>.txt                              (OT | TIPO | ARCHIVOS, sale de la API de Maximo)
  ../descargas/pautas-xlsx/<PLN>.xlsx           (PLN nuevos bajados de SharePoint)
"""
import os, sys, glob, shutil, subprocess, json, re

base = os.path.dirname(os.path.abspath(__file__))
pwa = os.path.dirname(base)
if len(sys.argv) < 2 or not sys.argv[1].isdigit():
    sys.exit(__doc__)
W = sys.argv[1]
aplicar = '--aplicar' in sys.argv or '--sin-push' in sys.argv
push = '--sin-push' not in sys.argv

def run(*cmd, **kw):
    print('\n$', ' '.join(cmd))
    r = subprocess.run(cmd, cwd=base, **kw)
    if r.returncode != 0:
        sys.exit(f'FALLO: {" ".join(cmd)}')

# --- Excel del programa: buscarlo en descargas/ o en ~/Downloads
dest = os.path.join(pwa, 'descargas', f'PROGRAMA_CON_LUB_Y26W{W}.xlsx')
if not os.path.exists(dest):
    cands = glob.glob(os.path.expanduser(f'~/Downloads/*CON_LUB_Y26W{W}*.xlsx'))
    if not cands:
        sys.exit(f'No encuentro el Excel de la semana {W} (ni en descargas/ ni en ~/Downloads)')
    cands.sort(key=os.path.getmtime)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    shutil.copy(cands[-1], dest)
    print('Excel copiado desde', cands[-1])

if not aplicar:
    run('python3', 'parse_plan.py', W)
    print(f'\nSIGUIENTE: consultar Maximo con las OT pendientes -> guardar maximo_w{W}.txt -> python3 semana.py {W} --aplicar')
    sys.exit(0)

# --- paso 2
if not os.path.exists(os.path.join(base, f'maximo_w{W}.txt')):
    sys.exit(f'Falta maximo_w{W}.txt (resultado de Maximo)')
if not os.path.exists(os.path.join(base, f'plan_w{W}.json')):
    run('python3', 'parse_plan.py', W)
run('python3', 'apply_maximo.py', W)

# PLN nuevos: los referenciados por el plan que no esten en pautas.json
plan = json.load(open(os.path.join(base, f'plan_w{W}.json'), encoding='utf-8'))
lib = {p['id'] for p in json.load(open(os.path.join(base, 'pautas.json'), encoding='utf-8'))}
nuevos = sorted({o['pauta'] for d in plan['days'] for o in d['ots'] if o.get('pauta') and o['pauta'] not in lib})
xdir = os.path.join(pwa, 'descargas', 'pautas-xlsx')
os.makedirs(xdir, exist_ok=True)
tras = os.path.expanduser('~/Documents/Aplicacion lubricación traspaso/Traspaso-Lubricacion-Mac/pautas-xlsx')
faltan = []
for p in nuevos:
    x = os.path.join(xdir, p + '.xlsx')
    if not os.path.exists(x):
        for c in (os.path.join(tras, p + '.xlsx'), os.path.expanduser(f'~/Downloads/{p}.xlsx')):
            if os.path.exists(c):
                shutil.copy(c, x); print('xlsx desde', c); break
    if not os.path.exists(x):
        faltan.append(p)
if faltan:
    sys.exit('Faltan los xlsx de: ' + ' '.join(faltan) + f'\n(bajarlos de SharePoint a {xdir} y volver a correr)')
for p in nuevos:
    run('python3', 'extract_pauta.py', os.path.join(xdir, p + '.xlsx'), p)
if nuevos:
    run('python3', 'build_pauta_html.py', *nuevos)
run('python3', 'qa_cellmaps.py')
run('python3', 'qa_htmls.py')
run('python3', 'build.py')
run('python3', '-W', 'ignore', 'publicar_datos.py')
if push:
    run('git', 'add', '-A', cwd=pwa)
    r = subprocess.run(['git', 'commit', '-q', '-m', f'Semana {W} del programa CON_LUB cargada ({len(nuevos)} pautas nuevas)\n\nCo-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>'], cwd=pwa)
    run('git', 'push', 'origin', 'main', cwd=pwa)
print(f'\nSEMANA {W} LISTA | pautas nuevas: {len(nuevos)}')
