/* ======================= Ventes (pilote POS) =======================
 * Écran « Vente » : l'employé note chaque vente en 3 touches (Vente, gallons, Enregistrer).
 * Seulement gallons et mode de paiement ; le prix est celui du bureau et n'est jamais montré ni tapé sur le téléphone.
 * Pilote (propriétaire, 28/11/2026) : allumé seulement pour la station de formation (code TST, ex. TST-0001).
 * Brochette viendra après l'essai, sur accord séparé du propriétaire : il suffira alors d'ajouter sa station dans
 * VENTE_PILOTE du serveur (Mobile.gs) ; rien d'autre ne change dans l'application.
 * Chaque vente est gardée sur le téléphone tout de suite, puis envoyée quand il y a du signal. Une vente envoyée
 * ne peut pas être effacée par la station : « Signaler une erreur » prévient le bureau, qui corrige.
 * Tant que le serveur ne connaît pas encore les ventes (pas de vente_pilote dans init), elles restent en attente sur le téléphone. */
var VT = { envoi: false, gal: 0, mode: 'cash', db: null, liste: [] };
var VT_TAILLES = [5, 10, 20];
var VT_STATIONS_APP = ['TST'];   // garde-fou côté téléphone : même si le serveur se trompe, l'écran ne s'ouvre que pour ces stations

function vtInit() { return S.init || CFG.get('init', null) || {}; }
function venteOn() {
  if (CFG.get('role', '') || !CFG.get('jeton', '')) return false;
  return VT_STATIONS_APP.indexOf(String(stationCode() || '')) >= 0;
}
function vtServeurPret() { return !!vtInit().vente_pilote; }

function vtDb() {
  if (VT.db) return Promise.resolve(VT.db);
  return new Promise(function (res, rej) {
    var r = indexedDB.open('rrd_ventes', 1);
    r.onupgradeneeded = function () { r.result.createObjectStore('ventes', { keyPath: 'cid' }); };
    r.onsuccess = function () { VT.db = r.result; res(VT.db); };
    r.onerror = function () { rej(r.error); };
  });
}
function vtTx(mode, fn) {
  return vtDb().then(function (d) {
    return new Promise(function (res, rej) {
      var t = d.transaction('ventes', mode), s = t.objectStore('ventes'), out = fn(s);
      t.oncomplete = function () { res(out && out.result !== undefined ? out.result : out); }; t.onerror = function () { rej(t.error); }; t.onabort = function () { rej(t.error); };
    });
  });
}
function vtMiennes() { var c = stationCode(); return vtTx('readonly', function (s) { return s.getAll(); }).then(function (l) { return (l || []).filter(function (r) { return r.station === c; }); }); }
function vtGarder(r) {
  return vtTx('readwrite', function (s) { s.put(r); }).then(function () {
    if (NAT.Filesystem) NAT.Filesystem.writeFile({ path: 'rrd_ventes/' + r.cid + '.json', directory: 'DATA', data: JSON.stringify(r), encoding: 'utf8', recursive: true }).catch(function () { });
  });
}
function vtCid() { return refAleatoire('V' + String(stationCode() || ''), dateRapport()).replace(/[^A-Za-z0-9_-]/g, '-'); }
function vtCorps(r) {
  return ['vente', { payload: { cid: r.cid, genre: r.genre, date: r.date, pris_ms: r.pris_ms, gallons: r.gallons, paiement: r.paiement, client: r.client, parent: r.parent, motif: r.motif } }];
}
function vtDuJour(l, date) {
  var v = l.filter(function (r) { return r.genre === 'vente' && r.date === date; });
  var sig = {}; l.forEach(function (r) { if (r.genre === 'erreur') sig[r.parent] = r; });
  var gal = 0, credit = 0; v.forEach(function (r) { gal += Number(r.gallons) || 0; if (r.paiement === 'credit') credit += Number(r.gallons) || 0; });
  return { ventes: v, signalees: sig, gallons: Math.round(gal * 100) / 100, credit: Math.round(credit * 100) / 100,
    attente: l.filter(function (r) { return r.statut === 'attente'; }).length };
}

/* Bloc sur l'accueil : bouton Vente et total du jour. Rien pour les autres stations. */
function venteAccueilHtml(l) {
  if (!venteOn()) return '';
  var j = vtDuJour(l || [], dateRapport());
  return '<button class="bigbtn" onclick="ecranVente()">🧾 Vente</button>' +
    '<p class="mut center">Aujourd\'hui : ' + j.ventes.length + ' vente(s), ' + fmt(j.gallons) + ' gal' + (j.attente ? ' · ⏳ ' + j.attente + ' en attente d\'envoi' : '') + '</p>';
}
function vtAvecListe(x) {   // appelé par accueil() : charge les ventes du téléphone ; ne fait rien pour les autres stations
  if (!venteOn()) { VT.liste = []; return Promise.resolve(x); }
  return vtMiennes().then(function (l) { VT.liste = l; return x; }, function () { VT.liste = []; return x; });
}

function ecranVente(msg) {
  if (!venteOn()) return accueil();
  if (msg) { VT.msg = msg; VT.msgT = Date.now(); } else if (Date.now() - (VT.msgT || 0) < 10000) msg = VT.msg;   // la confirmation reste 10 s, même si l'écran se rafraîchit
  vtMiennes().then(function (l) { return l; }, function () { return []; }).then(function (l) {
    var j = vtDuJour(l, dateRapport());
    var h = '<h1>🧾 Vente</h1><div class="sub">Journée du ' + dfr(dateRapport()) + ' · ' + j.ventes.length + ' vente(s), ' + fmt(j.gallons) + ' gal</div>';
    if (msg) h += '<div class="okb">' + esc(msg) + '</div>';
    if (j.attente && !vtServeurPret()) h += '<div class="note">Essai : les ventes restent sur ce téléphone tant que le bureau n\'a pas ouvert la réception des ventes.</div>';
    else if (j.attente) h += '<div class="wait">⏳ ' + j.attente + ' en attente d\'envoi. Elles partiront toutes seules dès qu\'il y aura du signal.</div>';
    h += '<label>Gallons</label><div class="ph3">' + VT_TAILLES.map(function (g) {
      return '<button class="' + (VT.gal === g ? 'fait' : 'sec') + '" style="font-size:24px;font-weight:600" onclick="vtTaille(' + g + ')">' + g + '</button>';
    }).join('') + '<button class="' + (VT.gal && VT_TAILLES.indexOf(VT.gal) < 0 ? 'fait' : 'sec') + '" onclick="vtAutre()">' + (VT.gal && VT_TAILLES.indexOf(VT.gal) < 0 ? fmt(VT.gal) : 'Autre') + '</button></div>';
    h += '<div id="vtAutreBox" class="hidden"><input id="vtAutreIn" type="number" inputmode="decimal" step="0.1" min="0" placeholder="ex. 7,5" oninput="vtAutreLu()"></div>';
    h += '<label>Paiement</label><div class="ph3">' +
      '<button class="' + (VT.mode === 'cash' ? 'fait' : 'sec') + '" onclick="vtMode(\'cash\')">💵 Cash</button>' +
      '<button class="' + (VT.mode === 'credit' ? 'fait' : 'sec') + '" onclick="vtMode(\'credit\')">📒 Crédit</button></div>';
    if (VT.mode === 'credit') {
      var connus = {}; l.forEach(function (r) { if (r.client) connus[r.client] = 1; });
      h += '<label for="vtClient">Client *</label><input id="vtClient" list="vtClients" autocomplete="off" value="' + esc(VT.client || '') + '" oninput="VT.client=this.value">' +
        '<datalist id="vtClients">' + Object.keys(connus).map(function (n) { return '<option value="' + esc(n) + '">'; }).join('') + '</datalist>';
    }
    var recentes = j.ventes.slice().sort(function (a, b) { return b.pris_ms - a.pris_ms; }).slice(0, 15);
    if (recentes.length) {
      h += '<h2>Ventes d\'aujourd\'hui</h2>';
      recentes.forEach(function (r) {
        var s = j.signalees[r.cid];
        var etat = r.statut === 'envoye' ? '✅' : r.statut === 'attente' ? '⏳' : '❌';
        h += '<div class="card"><div class="row"><div><b>' + fmt(r.gallons) + ' gal</b> · ' + (r.paiement === 'credit' ? 'Crédit ' + esc(r.client || '') : 'Cash') +
          '<div class="mut">' + esc(r.heure || '') + ' ' + etat + (s ? ' · ⚠️ erreur signalée' : '') + (r.statut === 'rejete' ? ' · refusée : ' + esc(r.erreur || '') : '') + '</div></div>' +
          (s ? '' : '<div style="text-align:right;flex:none"><button class="sec" style="padding:8px 10px;font-size:14px" onclick="vtSignaler(\'' + esc(r.cid) + '\')">Signaler une erreur</button></div>') + '</div></div>';
      });
    }
    view(h, '<button class="sec" onclick="VT.gal=0;accueil()">🏠 Accueil</button><button id="vtOk" onclick="vtEnregistrer()"' + (VT.gal > 0 ? '' : ' disabled') + '>Enregistrer</button>', 'vente');
  });
}
function vtTaille(g) { VT.gal = g; ecranVente(); }
function vtMode(m) { VT.mode = m; ecranVente(); }
function vtAutre() { var b = $('vtAutreBox'); if (b) { b.classList.remove('hidden'); var i = $('vtAutreIn'); if (i) i.focus(); } }
function vtAutreLu() { var v = nb($('vtAutreIn').value); VT.gal = v > 0 ? v : 0; var b = $('vtOk'); if (b) b.disabled = !(VT.gal > 0); }
function vtEnregistrer() {
  var g = Number(VT.gal) || 0;
  if (!(g > 0)) return toast('Choisissez les gallons.');
  if (g > 500) return alerte('Plus de 500 gallons pour une seule vente : vérifiez le chiffre.');
  var client = '';
  if (VT.mode === 'credit') { client = String((($('vtClient') || {}).value) || VT.client || '').trim().replace(/\s+/g, ' ').slice(0, 60); if (!client) return alerte('Écrivez le nom du client pour une vente à crédit.'); }
  var b = $('vtOk'); if (b) b.disabled = true;
  var ms = Date.now(), p = partiesHaiti(new Date(ms));
  var r = { cid: vtCid(), genre: 'vente', station: stationCode(), date: dateRapport(), pris_ms: ms, heure: p.hour + ':' + p.minute, gallons: g, paiement: VT.mode, client: client, statut: 'attente' };
  vtGarder(r).then(function () {
    var dit = fmt(g) + ' gal ' + (VT.mode === 'credit' ? 'à crédit (' + client + ')' : 'cash');
    VT.gal = 0; VT.mode = 'cash'; VT.client = '';
    ecranVente('Vente enregistrée : ' + dit + '.'); vtEnvoyer();
  }, function () { if (b) b.disabled = false; alerte('Impossible de garder la vente sur le téléphone (mémoire pleine ?). Réessayez.'); });
}
function vtSignaler(cid) {
  var motif = window.prompt('Qu\'est-ce qui est faux dans cette vente ? (ex. 10 gal et pas 5, client, cash au lieu de crédit)');
  if (motif === null) return;
  motif = String(motif).trim().slice(0, 200);
  if (!motif) return alerte('Expliquez l\'erreur pour que le bureau puisse corriger.');
  vtMiennes().then(function (l) {
    var v = l.filter(function (x) { return x.cid === cid; })[0]; if (!v) return;
    return vtGarder({ cid: vtCid(), genre: 'erreur', parent: cid, motif: motif, station: v.station, date: v.date, pris_ms: Date.now(), statut: 'attente' });
  }).then(function () { ecranVente('Erreur signalée au bureau. Il corrigera la vente.'); vtEnvoyer(); });
}

/* Envoi de la file, de la plus ancienne à la plus récente. Le serveur ne garde jamais deux fois la même (cid). */
function vtEnvoyer() {
  if (VT.envoi || !venteOn() || !vtServeurPret()) return Promise.resolve();
  VT.envoi = true;
  return vtMiennes().then(function (l) {
    var file = l.filter(function (r) { return r.statut === 'attente'; }).sort(function (a, b) { return a.pris_ms - b.pris_ms; }), i = 0;
    VT.bouge = file.length > 0;
    function suivant() {
      if (i >= file.length) return;
      var r = file[i++];
      return apiFile(vtCorps(r), r.cid).then(function () { r.statut = 'envoye'; }, function (e) {
        if (e.reseau || e.code === 'NON_ENREGISTRE' || e.code === 'DESACTIVE' || e.code === 'SERVEUR') throw e;   // on réessaiera
        r.statut = 'rejete'; r.erreur = e.erreur || '';
      }).then(function () { return vtGarder(r); }).then(suivant);
    }
    return suivant();
  }).catch(function () { }).then(function () {
    VT.envoi = false;
    if (typeof fondAligner === 'function') fondAligner();
    var tape = document.activeElement && /^vt(AutreIn|Client)$/.test(document.activeElement.id || '');   // on n'efface pas un chiffre en train d'être tapé
    if (VT.bouge && !tape) fond('vente', function () { ecranVente(); });
    return vtMiennes().then(function (l) {   // 3 jours d'historique sur le téléphone ; ce qui n'est pas parti reste
      var lim = Date.now() - 3 * 864e5, vieux = l.filter(function (r) { return r.pris_ms < lim && r.statut !== 'attente'; });
      if (vieux.length) return vtTx('readwrite', function (s) { vieux.forEach(function (r) { s.delete(r.cid); }); });
    }).catch(function () { });
  });
}
/* Pour l'envoi quand l'application est fermée : seulement quand le serveur reçoit les ventes. */
function vtFond() {
  if (!venteOn() || !vtServeurPret()) return Promise.resolve([]);
  return vtMiennes().then(function (l) { return l.filter(function (r) { return r.statut === 'attente'; }).map(function (r) { return { cid: r.cid, ordre: r.pris_ms, corps: vtCorps(r) }; }); }, function () { return []; });
}
/* Contrôle du jour avant d'envoyer le rapport : gallons du compteur contre gallons notés en ventes. */
function vtControleHtml(date, galCompteur) {
  if (!venteOn() || !(galCompteur >= 0)) return Promise.resolve('');
  return vtMiennes().then(function (l) {
    var j = vtDuJour(l, date), manque = Math.round((galCompteur - j.gallons) * 100) / 100;
    var cls = Math.abs(manque) <= Math.max(2, 0.02 * galCompteur) ? 'okb' : 'wait';
    return '<div class="' + cls + '">🧾 Ventes notées : ' + fmt(j.gallons) + ' gal (' + j.ventes.length + ' vente(s)) · Compteur : ' + fmt(galCompteur) + ' gal' +
      (manque > 0 ? ' · ' + fmt(manque) + ' gal sans vente notée' : manque < 0 ? ' · ' + fmt(-manque) + ' gal de plus que le compteur' : '') + '</div>';
  }, function () { return ''; });
}
