# Essais automatiques de l'application RRD, sur l'émulateur Android (GitHub Actions) ou dans Chrome (ordinateur).
#
#   Émulateur : python3 scenarios.py emulateur http://127.0.0.1:8765 dossier_captures
#   Chrome    : python3 scenarios.py chrome    http://127.0.0.1:8765 dossier_captures http://127.0.0.1:8080/index.html
#
# L'application est pilotée de l'intérieur (comme un doigt sur les boutons), le faux bureau dit ce qu'il a reçu.
# Chaque étape fait une capture d'écran. Le script s'arrête au premier essai raté et dit lequel.
import json, os, re, subprocess, sys, time, urllib.request

MODE, BUREAU, CAPTURES = sys.argv[1], sys.argv[2].rstrip('/'), sys.argv[3]
PAQUET = 'com.redrebel.rrdapp'
os.makedirs(CAPTURES, exist_ok=True)
RESULTATS = []


def bureau(chemin):
    with urllib.request.urlopen(BUREAU + chemin, timeout=10) as r:
        return json.loads(r.read().decode('utf-8'))


def recus(action):
    return [x for x in bureau('/_recu') if x['action'] == action]


# ---------------------------------------------------------------- pilotes
class Chrome:
    def __init__(self, adresse):
        from playwright.sync_api import sync_playwright
        self.adresse = adresse
        self.pw = sync_playwright().start()
        self.nav = self.pw.chromium.launch(executable_path='/usr/bin/google-chrome', args=['--no-sandbox'])
        self.ctx = self.nav.new_context(viewport={'width': 400, 'height': 820}, timezone_id='America/Port-au-Prince')
        self.page = self.ctx.new_page()
        self.page.on('pageerror', lambda e: ERREURS_JS.append(str(e)))
        self.page.goto(adresse)

    def ev(self, js):
        return self.page.evaluate('async () => { ' + js + ' }')

    def capture(self, nom):
        self.page.screenshot(path=os.path.join(CAPTURES, nom + '.png'))

    def redemarrer(self):
        self.page.goto(self.adresse)

    def fin(self):
        self.nav.close(); self.pw.stop()


class Emulateur:
    """Pilote la vue web de l'application installée, par le débogage à distance d'Android (adb + DevTools)."""

    def __init__(self):
        self.ws = None
        self.lancer()

    def adb(self, *a, **k):
        return subprocess.run(['adb'] + list(a), capture_output=True, text=True, **k).stdout

    def lancer(self):
        self.adb('shell', 'monkey', '-p', PAQUET, '-c', 'android.intent.category.LAUNCHER', '1')
        self.connecter()

    def connecter(self):
        import websocket
        fin = time.time() + 90
        while time.time() < fin:
            pid = self.adb('shell', 'pidof', PAQUET).strip()
            prises = re.findall(r'@(webview_devtools_remote_\d+)', self.adb('shell', 'cat', '/proc/net/unix'))
            prise = next((p for p in prises if pid and p.endswith('_' + pid.split()[0])), None)
            if prise:
                self.adb('forward', '--remove-all')
                self.adb('forward', 'tcp:9333', 'localabstract:' + prise)
                try:
                    pages = json.loads(urllib.request.urlopen('http://127.0.0.1:9333/json', timeout=5).read())
                    page = next((p for p in pages if p.get('type') == 'page' and 'webSocketDebuggerUrl' in p), None)
                    if page:
                        self.ws = websocket.create_connection(page['webSocketDebuggerUrl'], timeout=30, suppress_origin=True)
                        self.n = 0
                        self.envoi('Runtime.enable')
                        attendre(lambda: self.ev("return document.readyState === 'complete' && !!document.getElementById('app')"), 30, 'page chargée')
                        return
                except Exception:
                    pass
            time.sleep(2)
        raise Exception("Impossible de se connecter à la vue web de l'application (débogage à distance)")

    def envoi(self, methode, **params):
        self.n += 1
        self.ws.send(json.dumps({'id': self.n, 'method': methode, 'params': params}))
        while True:
            m = json.loads(self.ws.recv())
            if m.get('method') == 'Runtime.exceptionThrown':
                ERREURS_JS.append(m['params']['exceptionDetails'].get('exception', {}).get('description', str(m['params'])))
            if m.get('id') == self.n:
                return m

    def ev(self, js):
        m = self.envoi('Runtime.evaluate', expression='(async () => { ' + js + ' })()', awaitPromise=True, returnByValue=True)
        r = m.get('result', {})
        if 'exceptionDetails' in r:
            raise Exception('JS : ' + str(r['exceptionDetails'].get('exception', {}).get('description', r['exceptionDetails'])))
        return r.get('result', {}).get('value')

    def capture(self, nom):
        with open(os.path.join(CAPTURES, nom + '.png'), 'wb') as f:
            f.write(subprocess.run(['adb', 'exec-out', 'screencap', '-p'], capture_output=True).stdout)

    def redemarrer(self):
        try:
            self.ws.close()
        except Exception:
            pass
        self.adb('shell', 'am', 'force-stop', PAQUET)
        time.sleep(2)
        self.lancer()

    def fin(self):
        try:
            self.ws.close()
        except Exception:
            pass


ERREURS_JS = []


# ---------------------------------------------------------------- outils
def attendre(f, secondes, quoi):
    fin, derniere = time.time() + secondes, None
    while time.time() < fin:
        try:
            v = f()
            if v:
                return v
        except Exception as e:
            derniere = e
        time.sleep(1)
    raise AssertionError('Attendu en %d s : %s%s' % (secondes, quoi, (' (' + str(derniere) + ')') if derniere else ''))


def texte():
    return A.ev("return document.body.innerText")


def voit(motif, secondes=20):
    return attendre(lambda: re.search(motif, texte()), secondes, 'texte « %s » à l\'écran' % motif)


def toucher(motif):
    """Touche le premier bouton (ou lien) visible dont le texte correspond."""
    def essai():
        return A.ev("""const re = new RegExp(%s);
          const b = [...document.querySelectorAll('button, a')].find(x => x.offsetParent !== null && !x.disabled && re.test(x.innerText.trim()));
          if (!b) return false; b.scrollIntoView({block: 'center'}); b.click(); return true;""" % json.dumps(motif))
    attendre(essai, 20, 'bouton « %s »' % motif)
    time.sleep(0.8)


def ecrire(id_, valeur):
    attendre(lambda: A.ev("""const e = document.getElementById(%s); if (!e) return false;
      e.focus(); e.value = %s; e.dispatchEvent(new Event('input', {bubbles: true})); e.dispatchEvent(new Event('change', {bubbles: true})); return true;"""
                          % (json.dumps(id_), json.dumps(valeur))), 20, 'case « %s »' % id_)
    time.sleep(0.5)


def essai(nom):
    def deco(f):
        def lancer():
            t = time.time()
            try:
                f()
                A.capture(nom)
                RESULTATS.append((nom, True, '', time.time() - t))
                print('✅', nom, flush=True)
            except Exception as e:
                try:
                    A.capture(nom + '-ECHEC')
                except Exception:
                    pass
                RESULTATS.append((nom, False, str(e), time.time() - t))
                print('❌', nom, '—', e, flush=True)
                raise
        lancer.nom = nom
        return lancer
    return deco


# ---------------------------------------------------------------- scénarios
@essai('01-mauvais-code-refuse')
def mauvais_code():
    voit("Code d'activation", 60)
    ecrire('code', 'XXX-9999')
    toucher('^Enregistrer$')
    voit('inconnu ou désactivé')


@essai('02-activation-TST')
def activation():
    ecrire('code', 'TST-0001')
    toucher('^Enregistrer$')
    voit('Station TEST', 30)
    voit('Vente', 30)
    attendre(lambda: recus('init'), 30, 'le bureau reçoit « init »')


@essai('03-vente-cash-total-et-monnaie')
def vente_cash():
    toucher('Vente$')
    voit('Prix du bureau : 75')
    toucher('^20$')
    ecrire('vtRecu', '2000')
    t = voit(r'1\s?500')
    assert re.search(r'Monnaie à rendre : 500', texte()), "la monnaie à rendre (500) n'est pas affichée"
    toucher('^Enregistrer$')
    voit('Vente enregistrée')
    v = attendre(lambda: [x for x in recus('vente') if float(x['payload'].get('gallons') or 0) == 20], 45, 'le bureau reçoit la vente de 20 gal')
    p = v[-1]['payload']
    assert p['paiement'] == 'cash', 'paiement attendu cash, reçu %s' % p['paiement']
    assert float(p['prix_vu']) == 75, 'prix vu attendu 75, reçu %s' % p['prix_vu']
    assert p.get('recu_no'), 'pas de numéro de reçu'


@essai('04-credit-sans-nom-refuse')
def credit_sans_nom():
    toucher('Nouvelle vente')
    toucher('^5$')
    toucher('Crédit')
    toucher('^Enregistrer$')
    voit('nom du client')
    avant = len(recus('vente'))
    time.sleep(3)
    assert len(recus('vente')) == avant, 'une vente à crédit sans nom est partie au bureau'


@essai('05-credit-avec-nom')
def credit_nom():
    ecrire('vtClient', 'Client Essai')
    toucher('^Enregistrer$')
    voit('Vente enregistrée')
    attendre(lambda: [x for x in recus('vente') if x['payload'].get('client') == 'Client Essai' and x['payload'].get('paiement') == 'credit'], 45,
             'le bureau reçoit la vente à crédit de Client Essai')


@essai('06-sans-signal-la-vente-attend')
def hors_ligne():
    bureau('/_coupure/oui')
    toucher('Nouvelle vente')
    toucher('^10$')
    toucher('Cash')
    toucher('^Enregistrer$')
    voit('Vente enregistrée')
    toucher('Nouvelle vente')
    voit('en attente', 30)
    assert not [x for x in recus('vente') if float(x['payload'].get('gallons') or 0) == 10], 'la vente est partie alors que le bureau était coupé'


@essai('07-application-fermee-la-vente-reste')
def redemarrage():
    A.redemarrer()
    voit('Station TEST', 60)
    toucher('Vente$')
    voit('10 gal', 30)
    voit('en attente', 30)


@essai('08-signal-revenu-la-vente-part-seule')
def retour_signal():
    bureau('/_coupure/non')
    # l'application réessaie toute seule (au plus tard chaque minute)
    attendre(lambda: [x for x in recus('vente') if float(x['payload'].get('gallons') or 0) == 10], 150, 'la vente en attente arrive au bureau')
    def plus_en_attente():
        toucher('Accueil'); toucher('Vente$')
        return 'en attente' not in texte()
    attendre(plus_en_attente, 90, "« en attente » disparaît de l'écran des ventes")


@essai('09-signaler-une-erreur')
def signaler():
    A.ev("window.prompt = () => 'Essai : 10 gal et pas 20'; return true;")
    toucher('Signaler une erreur')
    voit('Erreur signalée au bureau')
    attendre(lambda: [x for x in recus('vente') if x['payload'].get('genre') == 'erreur' and 'Essai' in (x['payload'].get('motif') or '')], 45,
             "le bureau reçoit l'erreur signalée")


@essai('10-pas-de-bouton-supprimer')
def pas_supprimer():
    t = texte()
    assert not re.search(r'Supprimer|Effacer la vente', t), 'un bouton pour supprimer une vente est visible'


@essai('11-photos-ecran-ouvre')
def photos():
    toucher('Accueil')
    toucher('Photos de l')
    voit('COMPTEUR|Compteur|compteur', 30)


SCENARIOS = [mauvais_code, activation, vente_cash, credit_sans_nom, credit_nom, hors_ligne, redemarrage, retour_signal, signaler, pas_supprimer, photos]

if __name__ == '__main__':
    A = Chrome(sys.argv[4]) if MODE == 'chrome' else Emulateur()
    rate = False
    try:
        for s in SCENARIOS:
            try:
                s()
            except Exception:
                rate = True
                break
    finally:
        try:
            A.fin()
        except Exception:
            pass
        with open(os.path.join(CAPTURES, 'resume.md'), 'w', encoding='utf-8') as f:
            f.write('| Essai | Résultat | Durée |\n|---|---|---|\n')
            for nom, ok, err, d in RESULTATS:
                f.write('| %s | %s | %.0f s |\n' % (nom, '✅' if ok else '❌ ' + err.replace('|', '/')[:200], d))
            faits = {r[0] for r in RESULTATS}
            for s in SCENARIOS:
                if s.nom not in faits:
                    f.write('| %s | ⏭ pas lancé |  |\n' % s.nom)
            if ERREURS_JS:
                f.write('\n**Erreurs JavaScript vues :**\n\n' + '\n'.join('- ' + e.splitlines()[0][:200] for e in ERREURS_JS[:20]) + '\n')
        with open(os.path.join(CAPTURES, 'recu_par_le_bureau.json'), 'w', encoding='utf-8') as f:
            try:
                json.dump(bureau('/_recu'), f, ensure_ascii=False, indent=1)
            except Exception:
                pass
    sys.exit(1 if rate else 0)
