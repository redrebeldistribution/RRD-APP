# Faux serveur du bureau pour les essais sur émulateur (aucune donnée réelle n'est touchée).
# Imite le serveur Apps Script : POST /exec -> 302 vers /echo?k=… ; GET /echo -> la réponse JSON.
# Station d'essai : code TST-0001, prix 75 le gallon.
#
#   python3 faux_bureau.py 8765 journal.jsonl
#
# Commandes pour les scénarios (sans passer par l'application) :
#   GET /_coupure/oui  -> le serveur ne répond plus (comme un téléphone sans signal)
#   GET /_coupure/non  -> le serveur répond de nouveau
#   GET /_recu         -> tout ce que le serveur a reçu (liste JSON)
#   GET /_vider        -> oublie ce qui a été reçu
import json, sys, threading, time, urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
JOURNAL = sys.argv[2] if len(sys.argv) > 2 else 'journal.jsonl'
PRIX = 75
JETON = 'JETON-ESSAI'
ETAT = {'coupure': False, 'recu': [], 'reponses': {}, 'n': 0}
VERROU = threading.Lock()


def date_haiti():
    t = time.time() - 5 * 3600           # Haïti = UTC-5 ; avant 4 h du matin c'est encore la veille (comme le serveur)
    if time.gmtime(t).tm_hour < 4:
        t -= 86400
    return time.strftime('%Y-%m-%d', time.gmtime(t))


def init_station():
    return {
        'version': 'essai', 'devise': 'dollars haïtiens', 'station': {'code': 'TST', 'nom': 'Station TEST'},
        'date': date_haiti(), 'jours_max': 7, 'argent': 1, 'champs': [], 'envoyes': [], 'periodes': ['', 'AM'],
        'noms_connus': ['Jean'], 'prerempli': None, 'plafonds': {}, 'vente_pilote': 1,
        'vente_prix': {'prix': PRIX, 'depuis': '2026-11-01'}, 'annules': [], 'jours': [],
    }


def repondre(b):
    action, p = b.get('action'), b.get('payload') or {}
    if action == 'ping':
        return {'ok': True, 'data': {'version': 'essai'}}
    if action == 'enregistrer':
        if str(b.get('code', '')).strip().upper() != 'TST-0001':
            return {'ok': False, 'erreur': "Code d'activation inconnu ou désactivé.", 'code': 'CODE'}
        return {'ok': True, 'data': {'jeton': JETON, 'station': {'code': 'TST', 'nom': 'Station TEST'}}}
    if b.get('jeton') != JETON:
        return {'ok': False, 'erreur': 'Téléphone non enregistré.', 'code': 'NON_ENREGISTRE'}
    if action == 'init':
        return {'ok': True, 'data': init_station()}
    if action == 'vente':
        if p.get('genre') == 'vente' and not (float(p.get('gallons') or 0) > 0):
            return {'ok': False, 'erreur': 'Nombre de gallons illisible.', 'code': ''}
        return {'ok': True, 'data': {'recu': True, 'cid': p.get('cid')}}
    if action == 'envoyer':
        return {'ok': True, 'data': {'recu': {'ref': p.get('ref'), 'soumis_le': time.strftime('%Y-%m-%d %H:%M:%S')}}}
    if action in ('photo', 'bordereau'):
        return {'ok': True, 'data': {'recu': True, 'cid': p.get('cid')}}
    if action == 'photos':
        return {'ok': True, 'data': {'date': date_haiti(), 'heure': 10, 'heures': list(range(7, 17)), 'heures_vues': list(range(5, 20)),
                                     'types': ['COMPTEUR', 'STATION', 'CITERNE'], 'faits': {}, 'suivie': True}}
    if action == 'messages':
        return {'ok': True, 'data': {'date': date_haiti(), 'messages': []}}
    if action == 'votes':
        return {'ok': True, 'data': {'mois': date_haiti()[:7], 'ouvert': False, 'code': 'TST', 'miens': [], 'groupes': [], 'deja': [], 'accords': []}}
    if action == 'pushToken':
        return {'ok': True, 'data': {'garde': True}}
    if action == 'chatEntrer':
        return {'ok': False, 'erreur': "Chat fermé pendant l'essai.", 'code': ''}
    return {'ok': True, 'data': {}}


class H(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    def log_message(self, *a):
        pass

    def envoyer(self, code, corps=b'', type_='application/json; charset=utf-8', entetes=None):
        self.send_response(code)
        for k, v in (entetes or {}).items():
            self.send_header(k, v)
        self.send_header('Content-Type', type_)
        self.send_header('Content-Length', str(len(corps)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(corps)

    def couper(self):            # pas de réponse du tout : le téléphone croit qu'il n'a pas de signal
        self.close_connection = True
        try:
            self.connection.shutdown(2)
        except Exception:
            pass

    def do_OPTIONS(self):
        self.envoyer(204, entetes={'Access-Control-Allow-Methods': 'GET, POST', 'Access-Control-Allow-Headers': 'Content-Type'})

    def do_POST(self):
        n = int(self.headers.get('Content-Length', 0))
        brut = self.rfile.read(n).decode('utf-8', 'replace')
        if ETAT['coupure']:
            return self.couper()
        try:
            b = json.loads(brut)
        except Exception:
            b = {}
        rep = repondre(b)
        trace = {'le': time.strftime('%H:%M:%S'), 'action': b.get('action'), 'payload': b.get('payload') or {}, 'ok': rep.get('ok')}
        if isinstance(trace['payload'], dict) and 'photo' in trace['payload']:
            trace['payload'] = dict(trace['payload'], photo='(%d caractères)' % len(str(trace['payload']['photo'])))
        with VERROU:
            ETAT['recu'].append(trace)
            with open(JOURNAL, 'a', encoding='utf-8') as f:
                f.write(json.dumps(trace, ensure_ascii=False) + '\n')
            ETAT['n'] += 1
            k = str(ETAT['n'])
            ETAT['reponses'][k] = json.dumps(rep, ensure_ascii=False).encode('utf-8')
        if urllib.parse.urlparse(self.path).path.endswith('/exec'):
            self.envoyer(302, entetes={'Location': '/echo?k=' + k})
        else:
            self.envoyer(200, ETAT['reponses'][k])

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        if u.path == '/_coupure/oui':
            ETAT['coupure'] = True; return self.envoyer(200, b'{"coupure":true}')
        if u.path == '/_coupure/non':
            ETAT['coupure'] = False; return self.envoyer(200, b'{"coupure":false}')
        if u.path == '/_recu':
            return self.envoyer(200, json.dumps(ETAT['recu'], ensure_ascii=False).encode('utf-8'))
        if u.path == '/_vider':
            ETAT['recu'] = []; return self.envoyer(200, b'{}')
        if ETAT['coupure']:
            return self.couper()
        if u.path == '/echo':
            rep = ETAT['reponses'].get(urllib.parse.parse_qs(u.query).get('k', [''])[0])
            return self.envoyer(200, rep) if rep else self.envoyer(404)
        if u.path == '/web.json':      # mise à jour automatique : rien de nouveau pendant l'essai
            return self.envoyer(200, b'{"web":0,"native_min":0}')
        self.envoyer(404)


if __name__ == '__main__':
    open(JOURNAL, 'w').close()
    print('Faux bureau sur le port', PORT, flush=True)
    ThreadingHTTPServer(('0.0.0.0', PORT), H).serve_forever()
