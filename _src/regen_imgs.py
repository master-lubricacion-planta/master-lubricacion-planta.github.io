# -*- coding: utf-8 -*-
"""Regenera pautas-img/<PLN>.jpg para las 104 pautas aplicando el recorte srcRect
(la version anterior guardaba la imagen embebida cruda, que en varias pautas es
una captura de pantalla completa)."""
import json, os, sys
sys.stdout.reconfigure(encoding='utf-8')
import extract_pauta as E

SCRATCH = E.SCRATCH
PWA = E.PWA
DL = r'C:\Users\alexa\Downloads'
ALT = r'C:\Users\alexa\Desktop\Pautas'
pautas = json.load(open(os.path.join(SCRATCH, 'pautas.json'), encoding='utf-8'))
ok = 0; sin_img = []; err = []
for p in pautas:
    pid = p['id']
    xlsx = os.path.join(DL, pid + '.xlsx')
    if not os.path.exists(xlsx):
        xlsx = os.path.join(ALT, pid + '.xlsx')
    if not os.path.exists(xlsx):
        err.append((pid, 'sin xlsx')); continue
    try:
        rec, has_img = E.extract(xlsx, pid)
        if has_img: ok += 1
        else:
            sin_img.append(pid)
            viejo = os.path.join(PWA, 'pautas-img', pid + '.jpg')
            if os.path.exists(viejo):
                os.remove(viejo)  # imagen vieja sin figura real (era el logo u otra cosa)
    except Exception as e:
        err.append((pid, str(e)[:80]))
print('con imagen (recortada):', ok, '| sin figura:', len(sin_img), sin_img)
print('errores:', err)
