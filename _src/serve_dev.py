# -*- coding: utf-8 -*-
"""Servidor de desarrollo con gzip para la PWA.

El antivirus del PC retiene y trunca (~510KB) las respuestas HTTP locales
grandes sin comprimir; con gzip pautas.html viaja ~100KB y pasa limpio.
Uso: python serve_dev.py [puerto]  (default 8792, sirve lubricacion-pwa/)
"""
import http.server, gzip, os, sys

RAIZ = r'C:\Users\alexa\Desktop\Claude code Ses 1\lubricacion-pwa'
PUERTO = int(sys.argv[1]) if len(sys.argv) > 1 else 8792
COMPRIMIBLES = ('.html', '.js', '.json', '.css', '.svg', '.md')

class H(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=RAIZ, **kw)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()

    def do_GET(self):
        ruta = self.translate_path(self.path)
        if (os.path.isfile(ruta) and ruta.lower().endswith(COMPRIMIBLES)
                and 'gzip' in self.headers.get('Accept-Encoding', '')):
            with open(ruta, 'rb') as f:
                cuerpo = gzip.compress(f.read(), 6)
            self.send_response(200)
            self.send_header('Content-Type', self.guess_type(ruta))
            self.send_header('Content-Encoding', 'gzip')
            self.send_header('Content-Length', str(len(cuerpo)))
            self.end_headers()
            self.wfile.write(cuerpo)
            return
        super().do_GET()

    def log_message(self, *a):
        pass

if __name__ == '__main__':
    print(f'sirviendo {RAIZ} en http://127.0.0.1:{PUERTO} (gzip)')
    http.server.ThreadingHTTPServer(('127.0.0.1', PUERTO), H).serve_forever()
