# -*- coding: utf-8 -*-
import json, base64, io, os, re, unicodedata
from urllib.parse import quote
from PIL import Image
base=r'C:\Users\alexa\AppData\Local\Temp\claude\C--Users-alexa-Desktop-Claude-code-Ses-1\ee5231c6-7fdd-4492-a68b-a3a70ea480d3\scratchpad'
dl=r'C:\Users\alexa\Downloads'
pwa=r'C:\Users\alexa\Desktop\Claude code Ses 1\lubricacion-pwa'
BASEURL='https://master-lubricacion-planta.github.io/fichas/'
fold=lambda s: re.sub(r'\s+',' ',unicodedata.normalize('NFD',s).encode('ascii','ignore').decode().lower()).strip()
RULES=[
 (r'gadus s3 t ?100( 2)?$','gadus-s3-t100-2.pdf','gadus-s3-t100-2.pdf'),
 (r'gadus s2 v ?220 2$','gadus-s2-v220-2.pdf','gadus-s2-v220-2.pdf'),
 (r'gadus s2 v ?220 1$','enex-gadus-familia.pdf#page=21','gadus-s2-v220-1.pdf'),
 (r'gadus s2 v ?220 0$','enex-gadus-familia.pdf#page=19','gadus-s2-v220-0.pdf'),
 (r'gadus s3 v ?220 ?c?2(,.*)?$','gadus-s3-v220c-2.pdf','gadus-s3-v220c-2.pdf'),
 (r'gadus s3 v ?220 ?c?1$','gadus-s3-v220c-1.pdf','gadus-s3-v220c-1.pdf'),
 (r'gadus s5 v ?100 2$','gadus-s5-v100-2.pdf','gadus-s5-v100-2.pdf'),
 (r'gadus s5 v ?220( 2)?$','gadus-s5-v220-2.pdf','gadus-s5-v220-2.pdf'),
 (r'gadus s2 v ?100 3$','gadus-s2-v100-3.pdf','gadus-s2-v100-3.pdf'),
 (r'gadus s3 v ?460 ?d2.*$','gadus-s3-v460d-2.pdf','gadus-s3-v460d-2.pdf'),
 (r'gadus s2 og 85.*$','gadus-s2-og-85.pdf','gadus-s2-og-85.pdf'),
 (r'gadus s3 +high spe+ed coupling( grease)?$','gadus-s3-hsc.pdf','gadus-s3-hsc.pdf'),
 (r'gadus s2 ac 2$','gadus-s2-v220ac-2.pdf','gadus-s2-v220ac-2.pdf'),
 (r'omala (s2 g(xv|x)? ?220|iso vg 220)$','omala-s2-gx-220.pdf','omala-s2-gx-220.pdf'),
 (r'omala s2 gx? ?150$','omala-s2-gx-150.pdf','omala-s2-gx-150.pdf'),
 (r'omala s2 gx? ?320$','enex-omala-familia.pdf#page=13','omala-s2-gx-320.pdf'),
 (r'omala s2 gx? ?680$','enex-omala-familia.pdf#page=19','omala-s2-gx-680.pdf'),
 (r'omala s2 gx? ?68$','enex-omala-familia.pdf#page=1','omala-s2-gx-68.pdf'),
 (r'omala (s4 gxv|gxv s4) ?220$','enex-omala-familia.pdf#page=28','omala-s4-gxv-220.pdf'),
 (r'omala s4 gxv ?680$','omala-s4-gxv-680.pdf','omala-s4-gxv-680.pdf'),
 (r'tellus s2 vx? ?32$','tellus-s2-vx-32.pdf','tellus-s2-vx-32.pdf'),
 (r'tellus s2 vx? ?46$','tellus-s2-vx-46.pdf','tellus-s2-vx-46.pdf'),
 (r'tellus s2 vx? ?68$','tellus-s2-vx-68.pdf','tellus-s2-vx-68.pdf'),
 (r'tellus s2 mx ?32$','tellus-s2-mx-32.pdf','tellus-s2-mx-32.pdf'),
 (r'tellus s2 mx ?46$','tellus-s2-mx-46.pdf','tellus-s2-mx-46.pdf'),
 (r'tellus s2 mx ?68$','tellus-s2-mx-68.pdf','tellus-s2-mx-68.pdf'),
 (r'corena s4 r ?46$','corena-s4-r-46.pdf','corena-s4-r-46.pdf'),
 (r'corena s4 r ?68$','corena-s4-r-68.pdf','corena-s4-r-68.pdf'),
 (r'morlina s2 ba? ?220$','morlina-s2-b-220.pdf','morlina-s2-b-220.pdf'),
 (r'mouvex gliceine.*$',None,'glicerina-99-5.pdf'),
]
data=json.load(open(base+r'\lubdata.json',encoding='utf-8'))
names=set()
for r in data:
    for f in ('lub_actual','lub_recomendado'):
        if r.get(f): names.add(r[f].strip())
lubdocs={}; hdsdocs={}
for n in sorted(names):
    k=fold(n)
    for pat,ft,hds in RULES:
        if re.match(r'^(shell )?'+pat,k):
            if ft: lubdocs[k]=BASEURL+ft
            if hds: hdsdocs[k]=BASEURL+'hds/'+hds
            break
photos={
 'chancado': os.path.join(dl,'Chancado Imagen.jpg'),
 'molienda': os.path.join(dl,'Molienda imagen.jpg'),
 'pebble':   os.path.join(dl,'Pebble imagen.png'),
 'flotacion':os.path.join(dl,'Flotación imagen.png'),
 'tmf':      os.path.join(dl,'TMF imagen.jpg'),
 'flotcol':  os.path.join(dl,'manual_flotacion_colectiva.jpg'),
 'espesador':os.path.join(base,'manual_espesador.jpg'),
 'agua':     os.path.join(base,'manual_agua.jpg'),
 'concentrado':os.path.join(base,'manual_concentrado.jpg'),
}
uris={}
for k,f in photos.items():
    im=Image.open(f).convert('RGB')
    if im.width>420: im=im.resize((420,int(im.height*420/im.width)),Image.LANCZOS)
    buf=io.BytesIO(); im.save(buf,'JPEG',quality=70,optimize=True)
    uris[k]='data:image/jpeg;base64,'+base64.b64encode(buf.getvalue()).decode()
tecklogo='data:image/png;base64,'+base64.b64encode(open(base+r'\teck-logo.png','rb').read()).decode()
eqdocs={}
for line in open(base+r'\eqdocs.tsv',encoding='utf-8'):
    line=line.rstrip('\n')
    if not line.strip(): continue
    path,tags=line.split('\t')
    fname=path.split('/')[-1]
    name=re.sub(r'^\d+\.\s*','',fname)[:-4]
    url='https://teckresources.sharepoint.com/sites/manualinteractivo/'+quote(path)
    for t in tags.split(','): eqdocs[t.strip()]={'u':url,'n':name}
js=json.dumps(data,ensure_ascii=False,separators=(',',':')).replace('</','<\\/')
tpl=open(base+r'\app_template.html',encoding='utf-8').read()
app=(tpl.replace('__DATA__',js)
        .replace('__PHOTOS__',json.dumps(uris))
        .replace('__EQDOCS__',json.dumps(eqdocs,ensure_ascii=False))
        .replace('__LUBDOCS__',json.dumps(lubdocs,ensure_ascii=False))
        .replace('__HDSDOCS__',json.dumps(hdsdocs,ensure_ascii=False))
        .replace('__TECKLOGO__',tecklogo)
        .replace('__NREC__',str(len(data))))
open(base+r'\maestro-lubricacion.html','w',encoding='utf-8').write(app)
head='<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n<meta name="theme-color" content="#0b1c47">\n<link rel="manifest" href="manifest.json">\n<link rel="apple-touch-icon" href="apple-touch-icon.png">\n<link rel="icon" type="image/png" href="icon-192.png">\n<meta name="apple-mobile-web-app-capable" content="yes">\n<meta name="apple-mobile-web-app-status-bar-style" content="default">\n</head>\n<body>\n'
tail='\n<script>\nif(\'serviceWorker\' in navigator){\n  window.addEventListener(\'load\',()=>navigator.serviceWorker.register(\'sw.js\'));\n}\n</script>\n</body>\n</html>\n'
open(os.path.join(pwa,'index.html'),'w',encoding='utf-8').write(head+app+tail)

pautas=json.load(open(base+r'\pautas.json',encoding='utf-8'))
planes={}
for _w in ('33','35'):
    _f=base+'\\plan_w'+_w+'.json'
    if os.path.exists(_f): planes[_w]=json.load(open(_f,encoding='utf-8'))

# --- Validacion anti-regresion: toda pauta referenciada por el plan DEBE existir en
# pautas.json y tener su HTML generado. (Una referencia rota crasheaba la vista del
# dia completa: paso con PLN-815.) El build FALLA fuerte en vez de publicar roto.
ids_pautas={p['id'] for p in pautas}
rotos=[]
for w,plan in planes.items():
    for d in plan['days']:
        for o in d['ots']:
            if o.get('pauta') and o['pauta'] not in ids_pautas:
                rotos.append(f"semana {w} {d['nombre']} OT {o['ot']} -> {o['pauta']}")
if rotos:
    raise SystemExit('PLAN ROTO: OTs que apuntan a pautas inexistentes:\n  '+'\n  '.join(rotos))
sin_html=[pid for pid in ids_pautas if not os.path.exists(os.path.join(pwa,'pautas-html',pid+'.html'))]
usadas={o['pauta'] for plan in planes.values() for d in plan['days'] for o in d['ots'] if o.get('pauta')}
sin_html_usadas=sorted(set(sin_html)&usadas)
if sin_html_usadas:
    raise SystemExit('PAUTAS USADAS SIN HTML (correr build_pauta_html.py): '+', '.join(sin_html_usadas))
sin_cellmap_usadas=sorted(p['id'] for p in pautas if p['id'] in usadas and not p.get('cellmap'))
if sin_cellmap_usadas:
    print('AVISO: pautas usadas sin cellmap (el boton PDF avisara al tecnico):', ', '.join(sin_cellmap_usadas))
imgplns=sorted(f[:-4] for f in os.listdir(os.path.join(pwa,'pautas-img')) if f.endswith('.jpg'))
overrides={}
ovf=base+r'\overrides.json'
if os.path.exists(ovf): overrides=json.load(open(ovf,encoding='utf-8'))
# version del build (misma que tomara sw.js tras el bump de abajo): sirve de
# cache-busting para las imagenes de equipo (?v=N) sin romper el match offline
# del SW (usa ignoreSearch:true)
_sw=open(os.path.join(pwa,'sw.js'),encoding='utf-8').read()
BUILDV=str(int(re.search(r'lubricacion-v(\d+)',_sw).group(1))+1)
ptpl=open(base+r'\pautas_template.html',encoding='utf-8').read()
pout=(ptpl.replace('__IMGV__',BUILDV)
          .replace('__PAUTAS__',json.dumps(pautas,ensure_ascii=False).replace('</','<\\/'))
          .replace('__PLANES__',json.dumps(planes,ensure_ascii=False).replace('</','<\\/'))
          .replace('__IMGPLNS__',json.dumps(imgplns))
          .replace('__OVERRIDES__',json.dumps(overrides,ensure_ascii=False).replace('</','<\\/'))
          .replace('__TECKLOGO__',tecklogo))
open(os.path.join(pwa,'pautas.html'),'w',encoding='utf-8').write(pout)

ntpl=open(base+r'\panorama_template.html',encoding='utf-8').read()
nout=(ntpl.replace('__PAUTAS__',json.dumps(pautas,ensure_ascii=False).replace('</','<\\/'))
          .replace('__PLANES__',json.dumps(planes,ensure_ascii=False).replace('</','<\\/'))
          .replace('__TECKLOGO__',tecklogo))
open(os.path.join(pwa,'panorama.html'),'w',encoding='utf-8').write(nout)

rtpl=open(base+r'\ruta_template.html',encoding='utf-8').read()
open(os.path.join(pwa,'ruta.html'),'w',encoding='utf-8').write(rtpl.replace('__TECKLOGO__',tecklogo))

# Plan 52 semanas (seguimiento interno de inspecciones de lubricación)
p52i=open(base+r'\p52_items.json',encoding='utf-8').read()
p52p=open(base+r'\p52_prev.json',encoding='utf-8').read()
ptpl52=open(base+r'\plan52_template.html',encoding='utf-8').read()
open(os.path.join(pwa,'plan52.html'),'w',encoding='utf-8').write(
    ptpl52.replace('__P52ITEMS__',p52i.replace('</','<\\/'))
          .replace('__P52PREV__',p52p.replace('</','<\\/'))
          .replace('__TECKLOGO__',tecklogo))

# Tendencias (dashboard de análisis de aceite, contenido generado en cowork)
ttpl=open(base+r'\tendencias_template.html',encoding='utf-8').read()
tbody=open(base+r'\tendencias_body.html',encoding='utf-8').read()
open(os.path.join(pwa,'tendencias.html'),'w',encoding='utf-8').write(
    ttpl.replace('__TENDENCIAS_BODY__',tbody).replace('__TECKLOGO__',tecklogo))

sw=open(os.path.join(pwa,'sw.js'),encoding='utf-8').read()
m=re.search(r'lubricacion-v(\d+)',sw)
sw=sw.replace(m.group(0),'lubricacion-v'+str(int(m.group(1))+1))
if './pautas.html' not in sw:
    sw=sw.replace("'./manifest.json'","'./pautas.html', './manifest.json'")
if './panorama.html' not in sw:
    sw=sw.replace("'./manifest.json'","'./panorama.html', './manifest.json'")
if './ruta.html' not in sw:
    sw=sw.replace("'./manifest.json'","'./ruta.html', './manifest.json'")
if './tendencias.html' not in sw:
    sw=sw.replace("'./manifest.json'","'./tendencias.html', './manifest.json'")
if './plan52.html' not in sw:
    sw=sw.replace("'./manifest.json'","'./plan52.html', './manifest.json'")
open(os.path.join(pwa,'sw.js'),'w',encoding='utf-8').write(sw)
print('build OK | pautas.html', os.path.getsize(os.path.join(pwa,'pautas.html'))//1024,'KB')
