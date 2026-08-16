# -*- coding: utf-8 -*-
"""QA de los cellmaps: detecta pautas donde la inyeccion de datos del PDF
quedaria incompleta (sin OT, sin tecnicos, sin estado/valor, o con menos
actividades mapeadas que las del formulario)."""
import json, sys, os
sys.stdout.reconfigure(encoding='utf-8')
SCRATCH = r'C:\Users\alexa\AppData\Local\Temp\claude\C--Users-alexa-Desktop-Claude-code-Ses-1\ee5231c6-7fdd-4492-a68b-a3a70ea480d3\scratchpad'
PWA = r'C:\Users\alexa\Desktop\Claude code Ses 1\lubricacion-pwa'
pautas = json.load(open(os.path.join(SCRATCH, 'pautas.json'), encoding='utf-8'))
problemas = 0
for p in pautas:
    cm = p.get('cellmap')
    issues = []
    if not cm:
        issues.append('SIN CELLMAP')
    else:
        if not cm.get('ot'): issues.append('sin celda OT')
        if not cm.get('tecnicos'): issues.append('sin celdas tecnicos')
        acts_form = len(p.get('acts', []))
        acts_cm = len(cm.get('activities', []))
        if acts_cm != acts_form:
            issues.append(f'actividades: form={acts_form} cellmap={acts_cm}')
        else:
            sin_estado = [a['n'] for a in cm['activities'] if not (a.get('estado') or a.get('b'))]
            if sin_estado: issues.append(f'sin celda estado en acts {sin_estado[:5]}')
            sin_valor = [a['n'] for a in cm['activities'] if not a.get('valor')]
            if sin_valor: issues.append(f'sin celda valor en acts {sin_valor[:5]}')
        # el HTML debe existir
        if not os.path.exists(os.path.join(PWA, 'pautas-html', p['id'] + '.html')):
            issues.append('SIN HTML')
    if issues:
        problemas += 1
        print(p['id'], '|', p['titulo'][:35], '->', '; '.join(issues))
print()
print(f'{problemas} pautas con problemas de {len(pautas)}')
