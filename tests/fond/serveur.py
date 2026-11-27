# Imite le serveur Apps Script : POST /exec -> 302 vers /echo?k=… ; GET /echo -> la réponse JSON.
import json, sys, urllib.parse
from http.server import BaseHTTPRequestHandler, HTTPServer
REP = {}; RANG = [0]
class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def do_POST(self):
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query); mode = q.get('mode', ['ok'])[0]
        body = self.rfile.read(int(self.headers.get('Content-Length', 0))).decode('utf-8')
        if mode == '500': self.send_response(500); self.end_headers(); self.wfile.write(b'boom'); return
        if mode == 'html': rep = '<!DOCTYPE html><html><body>Erreur</body></html>'
        else:
            try:
                b = json.loads(body); p = b.get('payload', {})
                if self.headers.get('Content-Type') != 'text/plain;charset=utf-8': rep = json.dumps({'ok': False, 'erreur': 'type', 'code': 'TYPE'})
                elif b.get('jeton') == 'MAUVAIS': rep = json.dumps({'ok': False, 'erreur': 'Téléphone non enregistré.', 'code': 'NON_ENREGISTRE'})
                elif p.get('occupe'): rep = json.dumps({'ok': False, 'erreur': 'Erreur du serveur', 'code': 'SERVEUR'})
                elif p.get('refuse'): rep = json.dumps({'ok': False, 'erreur': 'Photo trop ancienne (plus de 31 jours).', 'code': ''}, ensure_ascii=False)
                else:
                    RANG[0] += 1
                    rep = json.dumps({'ok': True, 'data': {'recu': True, 'cid': p.get('cid'), 'rang': RANG[0], 'jeton_ok': b.get('jeton') in ('J', 'JET"ON\\1'), 'note_ok': p.get('note', 'é "@@JETON@@" ne doit pas changer') == 'é "@@JETON@@" ne doit pas changer', 'taille': len(p.get('photo', '')) // 100, 'action': b.get('action')}}, separators=(',', ':'))
                if p.get('cid') == 'A': RANG[0] = 0 if False else RANG[0]
            except Exception as e: rep = json.dumps({'ok': False, 'erreur': 'json: ' + str(e), 'code': 'JSON'})
        k = str(len(REP)); REP[k] = rep
        self.send_response(302); self.send_header('Location', '/echo?k=' + k); self.end_headers()
    def do_GET(self):
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query); rep = REP.get(q.get('k', [''])[0])
        if rep is None: self.send_response(404); self.end_headers(); return
        d = rep.encode('utf-8'); self.send_response(200); self.send_header('Content-Type', 'application/json; charset=utf-8'); self.send_header('Content-Length', str(len(d))); self.end_headers(); self.wfile.write(d)
HTTPServer(('127.0.0.1', int(sys.argv[1])), H).serve_forever()
