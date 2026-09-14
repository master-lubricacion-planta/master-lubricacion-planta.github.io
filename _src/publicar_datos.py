# -*- coding: utf-8 -*-
"""Sube a Firestore (coleccion 'datos') el paquete _src/publicar/datos.json que
genera build.py. Cada conjunto queda como documento meta `datos/<nombre>`
{partes, v, bytes, actualizado} + documentos `datos/<nombre>.p<i>` {s: trozo}
de menos de 900 KB (limite de Firestore: 1 MiB por documento).

Requiere la llave de cuenta de servicio del proyecto Firebase en
_src/secretos/serviceAccount.json (carpeta ignorada por git):
  Consola Firebase -> Configuracion del proyecto -> Cuentas de servicio
  -> Generar nueva clave privada.

Uso: python3 build.py && python3 publicar_datos.py
Solo sube los conjuntos cuyo contenido cambio (compara el hash `v`).
"""
import json, os, sys, hashlib

base = os.path.dirname(os.path.abspath(__file__))
KEY = os.path.join(base, 'secretos', 'serviceAccount.json')
PAQ = os.path.join(base, 'publicar', 'datos.json')
MAX = 900_000  # bytes por parte

if not os.path.exists(KEY):
    sys.exit('Falta la llave: ' + KEY + '\n(Consola Firebase -> Configuracion del proyecto -> Cuentas de servicio -> Generar nueva clave privada)')
if not os.path.exists(PAQ):
    sys.exit('Falta ' + PAQ + ' (correr build.py primero)')

import firebase_admin
from firebase_admin import credentials, firestore

firebase_admin.initialize_app(credentials.Certificate(KEY))
db = firebase_admin.firestore.client()
paq = json.load(open(PAQ, encoding='utf-8'))

def partir(texto):
    partes, cur, n = [], [], 0
    for ch in texto:
        l = len(ch.encode('utf-8'))
        if n + l > MAX:
            partes.append(''.join(cur)); cur, n = [], 0
        cur.append(ch); n += l
    if cur: partes.append(''.join(cur))
    return partes

subidos = 0
for nombre, obj in paq['datos'].items():
    s = json.dumps(obj, ensure_ascii=False, separators=(',', ':'))
    b = s.encode('utf-8')
    h = hashlib.sha1(b).hexdigest()[:12]
    ref = db.collection('datos').document(nombre)
    meta = ref.get()
    if meta.exists and meta.to_dict().get('v') == h:
        print('  =', nombre, '(sin cambios)'); continue
    partes = partir(s)
    batch = db.batch()
    for i, p in enumerate(partes):
        batch.set(db.collection('datos').document(f'{nombre}.p{i}'), {'s': p, 'i': i, 'v': h})
    # partes sobrantes de una version anterior mas larga
    viejas = meta.to_dict().get('partes', 0) if meta.exists else 0
    for i in range(len(partes), viejas):
        batch.delete(db.collection('datos').document(f'{nombre}.p{i}'))
    batch.set(ref, {'partes': len(partes), 'v': h, 'bytes': len(b), 'build': paq.get('v', ''),
                    'actualizado': firestore.SERVER_TIMESTAMP})
    batch.commit()
    subidos += 1
    print('  ^', nombre, f'{len(b)//1024} KB en {len(partes)} parte(s)')

print('publicar_datos OK |', subidos, 'conjunto(s) actualizado(s) de', len(paq['datos']))

# --- Replicas HTML de las pautas (pautas-html/<PLN>.html, 150 KB - 1.8 MB c/u):
# van tambien a la coleccion 'datos' como `rep.<PLN>` (texto JSON, en partes) para
# reusar la misma regla de lectura y el mismo cargador del candado. La app solo
# las pide al generar un PDF. La carpeta pautas-html/ ya no se publica en el sitio.
rep_dir = os.path.join(os.path.dirname(base), 'pautas-html')
subidas = 0; total = 0
for fn in sorted(os.listdir(rep_dir)):
    if not fn.endswith('.html'): continue
    total += 1
    nombre = 'rep.' + fn[:-5]
    s = json.dumps(open(os.path.join(rep_dir, fn), encoding='utf-8').read(), ensure_ascii=False)
    b = s.encode('utf-8')
    h = hashlib.sha1(b).hexdigest()[:12]
    ref = db.collection('datos').document(nombre)
    meta = ref.get()
    if meta.exists and meta.to_dict().get('v') == h: continue
    partes = partir(s)
    batch = db.batch()
    for i, p in enumerate(partes):
        batch.set(db.collection('datos').document(f'{nombre}.p{i}'), {'s': p, 'i': i, 'v': h})
    viejas = meta.to_dict().get('partes', 0) if meta.exists else 0
    for i in range(len(partes), viejas):
        batch.delete(db.collection('datos').document(f'{nombre}.p{i}'))
    batch.set(ref, {'partes': len(partes), 'v': h, 'bytes': len(b), 'actualizado': firestore.SERVER_TIMESTAMP})
    batch.commit()
    subidas += 1
    print('  ^', nombre, f'{len(b)//1024} KB en {len(partes)} parte(s)')
print('replicas OK |', subidas, 'actualizada(s) de', total)
