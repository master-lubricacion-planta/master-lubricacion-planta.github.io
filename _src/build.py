# -*- coding: utf-8 -*-
import json, base64, io, os, re, unicodedata
from urllib.parse import quote
from PIL import Image
base=os.path.dirname(os.path.abspath(__file__))  # _src del repo
img=os.path.join(base,'imagenes-areas')          # fotos de areas del Maestro
dl=os.path.join(os.path.dirname(base),'descargas')  # xlsx nuevos (Maximo/planillas)
pwa=os.path.dirname(base)                        # raiz del repo
BASEURL='https://master-lubricacion-planta.github.io/fichas/'
# candado de acceso + modo visualizador: se inyecta al inicio del <body> de toda
# página publicada (gate.html tiene las credenciales y el flag LUB_VIEWER)
_gate=open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'gate.html'),encoding='utf-8').read()
_FB='https://www.gstatic.com/firebasejs/10.14.1/'
_SDK_APP='<script src="'+_FB+'firebase-app-compat.js"></script>'
_SDK_AUTH='<script src="'+_FB+'firebase-auth-compat.js"></script>'
_SDK_FS='<script src="'+_FB+'firebase-firestore-compat.js"></script>'
def gate(html):
    # el candado necesita firebase-app + auth + firestore antes del <body>
    if _SDK_APP in html:
        if _SDK_AUTH not in html: html=html.replace(_SDK_APP,_SDK_APP+'\n'+_SDK_AUTH,1)
    else:
        html=html.replace('</head>',_SDK_APP+'\n'+_SDK_AUTH+'\n'+_SDK_FS+'\n</head>',1)
    return re.sub(r'<body[^>]*>', lambda m: m.group(0)+'\n'+_gate, html, count=1)
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
data=json.load(open(base+'/lubdata.json',encoding='utf-8'))
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
 'chancado': os.path.join(img,'Chancado Imagen.jpg'),
 'molienda': os.path.join(img,'Molienda imagen.jpg'),
 'pebble':   os.path.join(img,'Pebble imagen.png'),
 'flotacion':os.path.join(img,'Flotación imagen.png'),
 'tmf':      os.path.join(img,'TMF imagen.jpg'),
 'flotcol':  os.path.join(img,'manual_flotacion_colectiva.jpg'),
 'espesador':os.path.join(img,'manual_espesador.jpg'),
 'agua':     os.path.join(img,'manual_agua.jpg'),
 'concentrado':os.path.join(img,'manual_concentrado.jpg'),
}
uris={}
for k,f in photos.items():
    im=Image.open(f).convert('RGB')
    if im.width>420: im=im.resize((420,int(im.height*420/im.width)),Image.LANCZOS)
    buf=io.BytesIO(); im.save(buf,'JPEG',quality=70,optimize=True)
    uris[k]='data:image/jpeg;base64,'+base64.b64encode(buf.getvalue()).decode()
tecklogo='data:image/png;base64,'+base64.b64encode(open(base+'/teck-logo.png','rb').read()).decode()
eqdocs={}
for line in open(base+'/eqdocs.tsv',encoding='utf-8'):
    line=line.rstrip('\n')
    if not line.strip(): continue
    path,tags=line.split('\t')
    fname=path.split('/')[-1]
    name=re.sub(r'^\d+\.\s*','',fname)[:-4]
    url='https://teckresources.sharepoint.com/sites/manualinteractivo/'+quote(path)
    for t in tags.split(','): eqdocs[t.strip()]={'u':url,'n':name}
# ---------------------------------------------------------------------------
# DATOS: desde sep-2026 los datos NO van embebidos en las paginas (el repo y el
# sitio son publicos). build.py los junta en _src/publicar/datos.json y
# publicar_datos.py los sube a Firestore (coleccion 'datos', en partes <1MB).
# El candado (gate.html) los carga tras el login y quedan en cache offline.
# En las paginas, cada __PLACEHOLDER__ pasa a ser LUB.datos.<nombre> y el
# <script> que lo usa se envuelve en lubCuandoListo([...], function(){...}).
# ---------------------------------------------------------------------------
import datetime
DATOS={}
def nube(html, mapa):
    for ph,nombre in mapa.items():
        html=html.replace(ph,'LUB.datos.'+nombre)
    def wrap(m):
        cuerpo=m.group(2)
        usados=sorted(set(re.findall(r'LUB\.datos\.([A-Za-z0-9_]+)',cuerpo)))
        if not usados: return m.group(0)
        return m.group(1)+'\nlubCuandoListo('+json.dumps(usados)+', function(){\n'+cuerpo+'\n});\n'+m.group(3)
    return re.sub(r'(<script>)(.*?)(</script>)', wrap, html, flags=re.S)

DATOS['lubdata']=data; DATOS['eqdocs']=eqdocs; DATOS['lubdocs']=lubdocs; DATOS['hdsdocs']=hdsdocs
tpl=open(base+'/app_template.html',encoding='utf-8').read()
app=nube(tpl,{'__DATA__':'lubdata','__EQDOCS__':'eqdocs','__LUBDOCS__':'lubdocs','__HDSDOCS__':'hdsdocs'})
app=(app.replace('__PHOTOS__',json.dumps(uris))
        .replace('__TECKLOGO__',tecklogo)
        .replace('__NREC__',str(len(data))))
head='<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n<meta name="theme-color" content="#0b1c47">\n<link rel="manifest" href="manifest.json">\n<link rel="apple-touch-icon" href="apple-touch-icon.png">\n<link rel="icon" type="image/png" href="icon-192.png">\n<meta name="apple-mobile-web-app-capable" content="yes">\n<meta name="apple-mobile-web-app-status-bar-style" content="default">\n</head>\n<body>\n'
tail='\n<script>\nif(\'serviceWorker\' in navigator){\n  window.addEventListener(\'load\',()=>navigator.serviceWorker.register(\'sw.js\'));\n}\n</script>\n</body>\n</html>\n'
open(os.path.join(pwa,'index.html'),'w',encoding='utf-8').write(gate(head+app+tail))

pautas=json.load(open(base+'/pautas.json',encoding='utf-8'))
planes={}
import glob as _glob
for _f in sorted(_glob.glob(base+'/plan_w*.json')):
    _w=re.search(r'plan_w(\d+)\.json$',_f).group(1)
    planes[_w]=json.load(open(_f,encoding='utf-8'))

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
ovf=base+'/overrides.json'
if os.path.exists(ovf): overrides=json.load(open(ovf,encoding='utf-8'))
# version del build (misma que tomara sw.js tras el bump de abajo): sirve de
# cache-busting para las imagenes de equipo (?v=N) sin romper el match offline
# del SW (usa ignoreSearch:true)
_sw=open(os.path.join(pwa,'sw.js'),encoding='utf-8').read()
BUILDV=str(int(re.search(r'lubricacion-v(\d+)',_sw).group(1))+1)
DATOS['pautas']=pautas; DATOS['planes']=planes; DATOS['overrides']=overrides
MAPA_PAUTAS={'__PAUTAS__':'pautas','__PLANES__':'planes','__OVERRIDES__':'overrides'}
ptpl=open(base+'/pautas_template.html',encoding='utf-8').read()
pout=(nube(ptpl,MAPA_PAUTAS).replace('__IMGV__',BUILDV)
          .replace('__IMGPLNS__',json.dumps(imgplns))
          .replace('__TECKLOGO__',tecklogo))
open(os.path.join(pwa,'pautas.html'),'w',encoding='utf-8').write(gate(pout))

ntpl=open(base+'/panorama_template.html',encoding='utf-8').read()
open(os.path.join(pwa,'panorama.html'),'w',encoding='utf-8').write(gate(
    nube(ntpl,MAPA_PAUTAS).replace('__TECKLOGO__',tecklogo)))

DATOS['rprev']=json.load(open(base+'/rprev.json',encoding='utf-8'))
rtpl=open(base+'/ruta_template.html',encoding='utf-8').read()
open(os.path.join(pwa,'ruta.html'),'w',encoding='utf-8').write(gate(
    nube(rtpl,{'__RPREV__':'rprev'}).replace('__TECKLOGO__',tecklogo)))

# Plan 52 semanas (seguimiento interno de inspecciones de lubricación)
DATOS['p52_items']=json.load(open(base+'/p52_items.json',encoding='utf-8'))
DATOS['p52_prev']=json.load(open(base+'/p52_prev.json',encoding='utf-8'))
ptpl52=open(base+'/plan52_template.html',encoding='utf-8').read()
open(os.path.join(pwa,'plan52.html'),'w',encoding='utf-8').write(gate(
    nube(ptpl52,{'__P52ITEMS__':'p52_items','__P52PREV__':'p52_prev'}).replace('__TECKLOGO__',tecklogo)))

# Toma de muestras (programa mensual)
DATOS['mu_items']=json.load(open(base+'/mu_items.json',encoding='utf-8'))
DATOS['mu_prev']=json.load(open(base+'/mu_prev.json',encoding='utf-8'))
mtpl=open(base+'/muestras_template.html',encoding='utf-8').read()
open(os.path.join(pwa,'muestras.html'),'w',encoding='utf-8').write(gate(
    nube(mtpl,{'__MUITEMS__':'mu_items','__MUPREV__':'mu_prev'}).replace('__TECKLOGO__',tecklogo)))

# Consumibles (stock y movimientos de lubricantes)
DATOS['cons_maestro']=json.load(open(base+'/cons_maestro.json',encoding='utf-8'))
DATOS['cons_stock']=json.load(open(base+'/cons_stock.json',encoding='utf-8'))
DATOS['cons_hist']=json.load(open(base+'/cons_hist.json',encoding='utf-8'))
ctpl=open(base+'/consumo_template.html',encoding='utf-8').read()
open(os.path.join(pwa,'consumo.html'),'w',encoding='utf-8').write(gate(
    nube(ctpl,{'__CMAESTRO__':'cons_maestro','__CSTOCK__':'cons_stock','__CHIST__':'cons_hist'}).replace('__TECKLOGO__',tecklogo)))

# Tendencias (dashboard ALS): el `const DATA=[...]` del body pasa a la nube
ttpl=open(base+'/tendencias_template.html',encoding='utf-8').read()
tbody=open(base+'/tendencias_body.html',encoding='utf-8').read()
_m=re.search(r'const DATA=(\[.*?\]);\n',tbody,re.S)
DATOS['tendencias']=json.loads(_m.group(1))
tbody=tbody[:_m.start(1)]+'LUB.datos.tendencias'+tbody[_m.end(1):]
open(os.path.join(pwa,'tendencias.html'),'w',encoding='utf-8').write(gate(
    nube(ttpl.replace('__TENDENCIAS_BODY__',tbody),{}).replace('__TECKLOGO__',tecklogo)))

# Usuarios (administración de cuentas; solo rol admin, el candado redirige al resto)
utpl=open(base+'/usuarios_template.html',encoding='utf-8').read()
open(os.path.join(pwa,'usuarios.html'),'w',encoding='utf-8').write(gate(
    utpl.replace('__TECKLOGO__',tecklogo)))

# paquete de datos para publicar_datos.py (fuera del repo: .gitignore)
os.makedirs(base+'/publicar',exist_ok=True)
json.dump({'v':BUILDV,'generado':datetime.datetime.now().isoformat(timespec='seconds'),'datos':DATOS},
          open(base+'/publicar/datos.json','w',encoding='utf-8'),ensure_ascii=False,separators=(',',':'))
print('datos.json:', os.path.getsize(base+'/publicar/datos.json')//1024,'KB en',len(DATOS),'conjuntos -> python3 publicar_datos.py para subirlos')

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
if './muestras.html' not in sw:
    sw=sw.replace("'./manifest.json'","'./muestras.html', './manifest.json'")
if './consumo.html' not in sw:
    sw=sw.replace("'./manifest.json'","'./consumo.html', './manifest.json'")
if './usuarios.html' not in sw:
    sw=sw.replace("'./manifest.json'","'./usuarios.html', './manifest.json'")
open(os.path.join(pwa,'sw.js'),'w',encoding='utf-8').write(sw)
print('build OK | pautas.html', os.path.getsize(os.path.join(pwa,'pautas.html'))//1024,'KB')
