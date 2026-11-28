/* ======================= Ventes (pilote POS) =======================
 * Écran « Vente » : l'employé note chaque vente (gallons, Cash ou Crédit, Enregistrer).
 * Pilote (propriétaire, 28/11/2026) : allumé seulement pour la station de formation (code TST, ex. TST-0001).
 * Brochette viendra après l'essai, sur accord séparé du propriétaire : il faudra alors ajouter sa station dans
 * VENTE_PILOTE du serveur (Mobile.gs) ET dans VT_STATIONS_APP ci-dessous.
 * Build 78 (propriétaire, 27/11/2026 : « c'est une caisse ») : le téléphone montre le total de chaque vente (gallons × prix du bureau)
 * et la monnaie à rendre ; reçu au client par WhatsApp ou sur une imprimante Bluetooth. Le prix vient toujours du bureau
 * (init.vente_prix) : il n'est jamais tapé sur le téléphone, et le serveur recalcule le montant avec son propre prix.
 * Chaque vente est gardée sur le téléphone tout de suite, puis envoyée quand il y a du signal. Une vente envoyée
 * ne peut pas être effacée par la station : « Signaler une erreur » prévient le bureau, qui corrige.
 * Build 79 : bouton « Scanner le réservoir » (étiquette du client), facultatif.
 * Tant que le serveur ne connaît pas encore les ventes (pas de vente_pilote dans init), elles restent en attente sur le téléphone. */
var VT = { envoi: false, gal: 0, mode: 'cash', db: null, liste: [], res: '', scan: false };
var VT_TAILLES = [5, 10, 20];
var VT_STATIONS_APP = ['TST'];   // garde-fou côté téléphone : même si le serveur se trompe, l'écran ne s'ouvre que pour ces stations

function vtInit() { return S.init || CFG.get('init', null) || {}; }
function venteOn() {
  if (CFG.get('role', '') || !CFG.get('jeton', '')) return false;
  return VT_STATIONS_APP.indexOf(String(stationCode() || '')) >= 0;
}
function vtServeurPret() { return !!vtInit().vente_pilote; }
function vtPrix() { var p = vtInit().vente_prix; return p && Number(p.prix) > 0 ? Number(p.prix) : 0; }   // prix du bureau, 0 = pas encore reçu
function vtArr(v) { return Math.round(v * 100) / 100; }
/* ---------- Réservoir scanné (build 79, propriétaire 28/11/2026 : « le scan dans la page vente ») ----------
 * Étiquette 2x1 « version B » : QR « RRD:T:000123 », numéro imprimé « T-000123 ». Facultatif : une vente sans réservoir reste possible.
 * Premier scan d'un réservoir : on demande le téléphone du client (et son nom) pour l'enregistrer ; ensuite le réservoir est reconnu. */
function vtResCode(v) {
  var t = String(v || '').toUpperCase().replace(/\s+/g, ''), m = t.match(/^RRD:T:(\d{1,8})$/) || t.match(/^T-?(\d{1,8})$/) || t.match(/^(\d{1,8})$/);
  return m && Number(m[1]) > 0 ? 'T-' + ('000000' + Number(m[1])).slice(-6) : '';
}
function vtTel(v) { var d = String(v || '').replace(/\D/g, ''); if (d.length === 8) d = '509' + d; return d.length >= 10 && d.length <= 15 ? d : ''; }
function vtResConnus(l) {   // réservoirs connus : ceux du bureau (init) puis ceux des ventes de ce téléphone
  var m = {};
  (vtInit().reservoirs || []).forEach(function (x) { if (x && x.r) m[x.r] = { tel: String(x.tel || ''), nom: String(x.nom || '') }; });
  (l || VT.liste || []).slice().sort(function (a, b) { return a.pris_ms - b.pris_ms; }).forEach(function (r) {
    if (r.genre === 'vente' && r.reservoir && r.client_tel && !m[r.reservoir]) m[r.reservoir] = { tel: r.client_tel, nom: r.client_nom || r.client || '' };
  });
  return m;
}
function vtScanPossible() { return !WEB && !!NAT.BarcodeScanner && !(S.natif && S.natif < TT_NATIF); }
function vtScanner() {
  var P = NAT;
  if (!vtScanPossible()) return vtResTaper();
  P.BarcodeScanner.isSupported().then(function (x) { if (!x.supported) throw 'Ce téléphone n\'a pas de caméra utilisable.'; return P.BarcodeScanner.checkPermissions(); })
    .then(function (pm) { return pm.camera === 'granted' ? pm : P.BarcodeScanner.requestPermissions(); })
    .then(function (pm) {
      if (pm.camera !== 'granted') throw 'Permission de la caméra refusée.';
      return P.BarcodeScanner.removeAllListeners().catch(function () { }).then(function () {
        return P.BarcodeScanner.addListener('barcodesScanned', function (ev) {
          var b = ev && ev.barcodes && ev.barcodes[0]; if (!b || !VT.scan) return;
          var lu = String(b.rawValue || b.displayValue || ''), c = vtResCode(lu);
          vtScanFin();
          if (!c) return alerte('Ce code n\'est pas une étiquette de réservoir Red Rebel.');
          vtResChoisi(c);
        });
      });
    }).then(function () {
      VT.scan = true; document.body.classList.add('scan'); document.documentElement.classList.add('scan');
      var bar = document.createElement('div'); bar.id = 'scanBar'; bar.innerHTML = 'Montrez l\'étiquette du réservoir à la caméra<br><button class="sec" style="margin-top:10px" onclick="vtScanFin()">Arrêter</button> <button class="sec" style="margin-top:10px" onclick="vtScanFin();vtResTaper()">Taper le numéro</button>';
      $('app').appendChild(bar);
      VT.scanTimer = setTimeout(function () { if (VT.scan) { vtScanFin(); toast('Aucun code lu.'); } }, 45000);
      return P.BarcodeScanner.startScan({});
    }).catch(function (e) { vtScanFin(); alerte('Lecture du code : ' + ttErr(e)); });
}
function vtScanFin() {
  VT.scan = false; clearTimeout(VT.scanTimer); document.body.classList.remove('scan'); document.documentElement.classList.remove('scan');
  var b = $('scanBar'); if (b && b.parentNode) b.parentNode.removeChild(b);
  try { NAT.BarcodeScanner && NAT.BarcodeScanner.stopScan().catch(function () { }); } catch (e) { }
}
function vtResTaper() {
  var v = window.prompt('Numéro écrit sur l\'étiquette du réservoir (ex. T-000123)');
  if (v === null) return;
  var c = vtResCode(v); if (!c) return alerte('Numéro de réservoir illisible.');
  vtResChoisi(c);
}
function vtResChoisi(c) {
  VT.res = c; VT.resTel = ''; VT.resNom = '';
  var k = vtResConnus()[c];
  if (k && k.nom && VT.mode === 'credit' && !VT.client) VT.client = k.nom;
  ecranVente();
}
function vtResRetirer() { VT.res = ''; VT.resTel = ''; VT.resNom = ''; ecranVente(); }
function vtResHtml(l) {
  if (!VT.res) return '<div style="margin-top:12px;display:flex;gap:8px;flex-wrap:wrap">' + (vtScanPossible() ? '<button class="sec" onclick="vtScanner()">🔳 Scanner le réservoir</button>' : '') +
    '<button class="sec" onclick="vtResTaper()">⌨️ N° du réservoir</button></div><p class="mut">Facultatif : la vente peut se faire sans réservoir.</p>';
  var k = vtResConnus(l)[VT.res];
  var h = '<div class="card" style="margin-top:12px"><div class="row"><div><b>🔳 Réservoir ' + esc(VT.res) + '</b><div class="mut">' +
    (k ? 'Client : ' + esc(k.nom || '(sans nom)') + ' · ' + esc(k.tel) : 'Nouveau réservoir : à enregistrer avec le client') + '</div></div>' +
    '<div><button class="sec" style="padding:6px 10px" onclick="vtResRetirer()">Retirer</button></div></div>';
  if (!k) h += '<label for="vtResTel">Téléphone du client *</label><input id="vtResTel" type="tel" inputmode="tel" placeholder="ex. 3712 3456" autocomplete="off" value="' + esc(VT.resTel || '') + '" oninput="VT.resTel=this.value">' +
    '<label for="vtResNom">Nom du client</label><input id="vtResNom" autocomplete="off" value="' + esc(VT.resNom || '') + '" oninput="VT.resNom=this.value">';
  return h + '</div>';
}
function vtMontant(r) { return Number(r.montant_vu) > 0 ? Number(r.montant_vu) : (Number(r.prix_vu) > 0 ? vtArr((Number(r.gallons) || 0) * Number(r.prix_vu)) : null); }

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
function vtPoste() {   // deux lettres propres à ce téléphone : deux téléphones de la même station ne donnent jamais le même n° de reçu
  var p = CFG.get('vtPoste', '');
  if (!/^[A-Z]{2}$/.test(p)) { var L = 'ABCDEFGHJKLMNPQRSTUVWXYZ'; p = L.charAt(Math.floor(Math.random() * L.length)) + L.charAt(Math.floor(Math.random() * L.length)); CFG.set('vtPoste', p); }
  return p;
}
function vtRecuNo(l, date) {
  var n = l.filter(function (r) { return r.genre === 'vente' && r.date === date && r.recu_no; }).length + 1;
  return String(stationCode() || '') + '-' + date.slice(8, 10) + date.slice(5, 7) + date.slice(2, 4) + '-' + vtPoste() + ('00' + n).slice(-3);
}
function vtCorps(r) {
  return ['vente', { payload: { cid: r.cid, genre: r.genre, date: r.date, pris_ms: r.pris_ms, gallons: r.gallons, paiement: r.paiement, client: r.client, parent: r.parent, motif: r.motif, prix_vu: r.prix_vu || '', recu_no: r.recu_no || '', reservoir: r.reservoir || '', client_tel: r.client_tel || '', client_nom: r.client_nom || '' } }];
}
function vtDuJour(l, date) {
  var v = l.filter(function (r) { return r.genre === 'vente' && r.date === date; });
  var sig = {}; l.forEach(function (r) { if (r.genre === 'erreur') sig[r.parent] = r; });
  var gal = 0, credit = 0, mt = 0, mtCash = 0, mtCredit = 0, sansPrix = 0;
  v.forEach(function (r) {
    var g = Number(r.gallons) || 0, m = vtMontant(r); gal += g; if (r.paiement === 'credit') credit += g;
    if (m === null) sansPrix++; else { mt += m; if (r.paiement === 'credit') mtCredit += m; else mtCash += m; }
  });
  return { ventes: v, signalees: sig, gallons: vtArr(gal), credit: vtArr(credit), montant: vtArr(mt), cash_montant: vtArr(mtCash), credit_montant: vtArr(mtCredit), sans_prix: sansPrix,
    attente: l.filter(function (r) { return r.statut === 'attente'; }).length };
}

/* Bloc sur l'accueil : bouton Vente et total du jour. Rien pour les autres stations. */
function venteAccueilHtml(l) {
  if (!venteOn()) return '';
  var j = vtDuJour(l || [], dateRapport());
  return '<button class="bigbtn" onclick="ecranVente()">🧾 Vente</button>' +
    '<p class="mut center">Aujourd\'hui : ' + j.ventes.length + ' vente(s), ' + fmt(j.gallons) + ' gal' + (j.montant ? ', ' + fmt(j.montant) + ' ' + esc(DEV) : '') + (j.attente ? ' · ⏳ ' + j.attente + ' en attente d\'envoi' : '') + '</p>';
}
function vtAvecListe(x) {   // appelé par accueil() : charge les ventes du téléphone ; ne fait rien pour les autres stations
  if (!venteOn()) { VT.liste = []; return Promise.resolve(x); }
  return vtMiennes().then(function (l) { VT.liste = l; return x; }, function () { VT.liste = []; return x; });
}

function vtTotalHtml() {   // total de la vente en cours et monnaie à rendre (cash)
  var prix = vtPrix(), g = Number(VT.gal) || 0;
  if (!prix) return '<div class="wait">Prix pas encore reçu du bureau : le total n\'est pas montré. Touchez « Envoyer / actualiser » avec du signal.</div>';
  if (!(g > 0)) return '<p class="mut">Prix du bureau : ' + fmt(prix) + ' ' + esc(DEV) + ' le gallon.</p>';
  var t = vtArr(g * prix), h = '<div class="okb" style="text-align:center"><div class="mut">' + fmt(g) + ' gal × ' + fmt(prix) + '</div><div style="font-size:34px;font-weight:800">' + fmt(t) + ' <span style="font-size:16px">' + esc(DEV) + '</span></div>';
  if (VT.mode === 'cash') {
    var r = Number(VT.recu) || 0;
    if (r > 0) h += r >= t ? '<div style="font-size:20px;margin-top:6px">Monnaie à rendre : <b>' + fmt(vtArr(r - t)) + '</b></div>' : '<div style="font-size:18px;margin-top:6px;color:var(--bad)">Il manque ' + fmt(vtArr(t - r)) + '</div>';
  }
  return h + '</div>';
}
function vtMajTotal() { var d = $('vtTotal'); if (d) d.innerHTML = vtTotalHtml(); }

function ecranVente(msg) {
  if (!venteOn()) return accueil();
  if (msg) { VT.msg = msg; VT.msgT = Date.now(); } else if (Date.now() - (VT.msgT || 0) < 10000) msg = VT.msg;   // la confirmation reste 10 s, même si l'écran se rafraîchit
  vtMiennes().then(function (l) { return l; }, function () { return []; }).then(function (l) {
    var j = vtDuJour(l, dateRapport());
    var h = '<h1>🧾 Vente</h1><div class="sub">Journée du ' + dfr(dateRapport()) + ' · ' + j.ventes.length + ' vente(s), ' + fmt(j.gallons) + ' gal' +
      (j.montant ? ' · ' + fmt(j.montant) + ' ' + esc(DEV) + ' (cash ' + fmt(j.cash_montant) + ', crédit ' + fmt(j.credit_montant) + ')' : '') + '</div>';
    if (msg) h += '<div class="okb">' + esc(msg) + '</div>';
    if (j.attente && !vtServeurPret()) h += '<div class="note">Essai : les ventes restent sur ce téléphone tant que le bureau n\'a pas ouvert la réception des ventes.</div>';
    else if (j.attente) h += '<div class="wait">⏳ ' + j.attente + ' en attente d\'envoi. Elles partiront toutes seules dès qu\'il y aura du signal.</div>';
    h += '<label>Gallons</label><div class="ph3">' + VT_TAILLES.map(function (g) {
      return '<button class="' + (VT.gal === g ? 'fait' : 'sec') + '" style="font-size:24px;font-weight:600" onclick="vtTaille(' + g + ')">' + g + '</button>';
    }).join('') + '<button class="' + (VT.gal && VT_TAILLES.indexOf(VT.gal) < 0 ? 'fait' : 'sec') + '" onclick="vtAutre()">' + (VT.gal && VT_TAILLES.indexOf(VT.gal) < 0 ? fmt(VT.gal) : 'Autre') + '</button></div>';
    h += '<div id="vtAutreBox" class="hidden"><input id="vtAutreIn" type="number" inputmode="decimal" step="0.1" min="0" placeholder="ex. 7,5" oninput="vtAutreLu()"></div>';
    h += vtResHtml(l);
    h += '<label>Paiement</label><div class="ph3">' +
      '<button class="' + (VT.mode === 'cash' ? 'fait' : 'sec') + '" onclick="vtMode(\'cash\')">💵 Cash</button>' +
      '<button class="' + (VT.mode === 'credit' ? 'fait' : 'sec') + '" onclick="vtMode(\'credit\')">📒 Crédit</button></div>';
    if (VT.mode === 'credit') {
      var connus = {}; l.forEach(function (r) { if (r.client) connus[r.client] = 1; });
      h += '<label for="vtClient">Client *</label><input id="vtClient" list="vtClients" autocomplete="off" value="' + esc(VT.client || '') + '" oninput="VT.client=this.value">' +
        '<datalist id="vtClients">' + Object.keys(connus).map(function (n) { return '<option value="' + esc(n) + '">'; }).join('') + '</datalist>';
    } else if (vtPrix()) {
      h += '<label for="vtRecu">Argent reçu du client (pour la monnaie, facultatif)</label><input id="vtRecu" type="number" inputmode="decimal" min="0" value="' + (VT.recu ? esc(VT.recu) : '') + '" oninput="VT.recu=nb(this.value);vtMajTotal()">';
    }
    h += '<div id="vtTotal">' + vtTotalHtml() + '</div>';
    var recentes = j.ventes.slice().sort(function (a, b) { return b.pris_ms - a.pris_ms; }).slice(0, 15);
    if (recentes.length) {
      h += '<h2>Ventes d\'aujourd\'hui</h2>';
      recentes.forEach(function (r) {
        var s = j.signalees[r.cid], m = vtMontant(r);
        var etat = r.statut === 'envoye' ? '✅' : r.statut === 'attente' ? '⏳' : '❌';
        h += '<div class="card"><div class="row"><div><b>' + fmt(r.gallons) + ' gal</b>' + (m !== null ? ' · <b>' + fmt(m) + ' ' + esc(DEV) + '</b>' : '') + ' · ' + (r.paiement === 'credit' ? 'Crédit ' + esc(r.client || '') : 'Cash') +
          '<div class="mut">' + esc(r.heure || '') + (r.recu_no ? ' · reçu ' + esc(r.recu_no) : '') + (r.reservoir ? ' · 🔳 ' + esc(r.reservoir) : '') + ' ' + etat + (s ? ' · ⚠️ erreur signalée' : '') + (r.statut === 'rejete' ? ' · refusée : ' + esc(r.erreur || '') : '') + '</div></div>' +
          '<div style="text-align:right;flex:none;display:flex;flex-direction:column;gap:6px"><button class="sec" style="padding:8px 10px;font-size:14px" onclick="vtEcranRecu(\'' + esc(r.cid) + '\')">🧾 Reçu</button>' +
          (s ? '' : '<button class="sec" style="padding:8px 10px;font-size:14px" onclick="vtSignaler(\'' + esc(r.cid) + '\')">Signaler une erreur</button>') + '</div></div></div>';
      });
    }
    view(h, '<button class="sec" onclick="VT.gal=0;VT.res=\'\';accueil()">🏠 Accueil</button><button id="vtOk" onclick="vtEnregistrer()"' + (VT.gal > 0 ? '' : ' disabled') + '>Enregistrer</button>', 'vente');
  });
}
function vtTaille(g) { VT.gal = g; ecranVente(); }
function vtMode(m) { VT.mode = m; ecranVente(); }
function vtAutre() { var b = $('vtAutreBox'); if (b) { b.classList.remove('hidden'); var i = $('vtAutreIn'); if (i) i.focus(); } }
function vtAutreLu() { var v = nb($('vtAutreIn').value); VT.gal = v > 0 ? v : 0; var b = $('vtOk'); if (b) b.disabled = !(VT.gal > 0); vtMajTotal(); }
function vtEnregistrer() {
  var g = Number(VT.gal) || 0;
  if (!(g > 0)) return toast('Choisissez les gallons.');
  if (g > 500) return alerte('Plus de 500 gallons pour une seule vente : vérifiez le chiffre.');
  var client = '';
  if (VT.mode === 'credit') { client = String((($('vtClient') || {}).value) || VT.client || '').trim().replace(/\s+/g, ' ').slice(0, 60); if (!client) return alerte('Écrivez le nom du client pour une vente à crédit.'); }
  var res = VT.res || '', resTel = '', resNom = '';
  if (res) {
    var k = vtResConnus()[res];
    if (k) { resTel = k.tel; resNom = k.nom; }
    else {
      resTel = vtTel((($('vtResTel') || {}).value) || VT.resTel);
      resNom = String((($('vtResNom') || {}).value) || VT.resNom || '').trim().replace(/\s+/g, ' ').slice(0, 60);
      if (!resTel) return alerte('Nouveau réservoir : écrivez le téléphone du client (8 chiffres), ou retirez le réservoir.');
    }
    if (!resNom && client) resNom = client;
  }
  var prix = vtPrix(), total = prix ? vtArr(g * prix) : null, recu = VT.mode === 'cash' && total !== null && Number(VT.recu) >= total ? Number(VT.recu) : 0;
  var b = $('vtOk'); if (b) b.disabled = true;
  var ms = Date.now(), p = partiesHaiti(new Date(ms)), date = dateRapport();
  vtMiennes().then(function (l) { return l; }, function () { return []; }).then(function (l) {
    var r = { cid: vtCid(), genre: 'vente', station: stationCode(), date: date, pris_ms: ms, heure: p.hour + ':' + p.minute, gallons: g, paiement: VT.mode, client: client, statut: 'attente',
      prix_vu: prix || '', montant_vu: total === null ? '' : total, recu_client: recu || '', monnaie: recu ? vtArr(recu - total) : '', recu_no: vtRecuNo(l, date), reservoir: res, client_tel: resTel, client_nom: resNom };
    return vtGarder(r).then(function () {
      VT.gal = 0; VT.mode = 'cash'; VT.client = ''; VT.recu = 0; VT.res = ''; VT.resTel = ''; VT.resNom = '';
      vtEnvoyer(); vtEcranRecu(r.cid, true);
    });
  }).catch(function () { if (b) b.disabled = false; alerte('Impossible de garder la vente sur le téléphone (mémoire pleine ?). Réessayez.'); });
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

/* ---------- Reçu du client : à l'écran, par WhatsApp ou sur une imprimante Bluetooth ---------- */
function vtRecuLignes(r) {
  var st = CFG.get('station', {}) || {}, m = vtMontant(r), d = r.date ? r.date.slice(8, 10) + '/' + r.date.slice(5, 7) + '/' + r.date.slice(0, 4) : '';
  var n = function (v) { return Number(v).toLocaleString('fr-FR', { maximumFractionDigits: 2 }).replace(/\s/g, ' '); };
  var l = ['RED REBEL DISTRIBUTION', 'Propane - ' + (st.nom || r.station || ''), 'Reçu n° ' + (r.recu_no || r.cid), d + ' ' + (r.heure || ''), '--------------------------------',
    'Propane : ' + n(r.gallons) + ' gal' + (Number(r.prix_vu) > 0 ? ' x ' + n(r.prix_vu) : '')];
  if (r.reservoir) l.push('Réservoir : ' + r.reservoir);
  if (m !== null) l.push('TOTAL : ' + n(m) + ' ' + DEV);
  if (r.paiement === 'credit') l.push('Paiement : à crédit - ' + (r.client || ''));
  else { l.push('Paiement : cash'); if (Number(r.recu_client) > 0) { l.push('Reçu : ' + n(r.recu_client)); l.push('Monnaie rendue : ' + n(r.monnaie || 0)); } }
  l.push('--------------------------------', 'Merci de votre confiance !');
  return l;
}
function vtEcranRecu(cid, nouveau) {
  vtMiennes().then(function (l) {
    var r = l.filter(function (x) { return x.cid === cid; })[0];
    if (!r) return ecranVente();
    VT.recuCid = cid;
    var imp = CFG.get('imprimante', null);
    var h = '<h1>🧾 Reçu</h1>' + (nouveau ? '<div class="okb">Vente enregistrée.</div>' : '') +
      '<div class="card" style="font-family:monospace;white-space:pre-wrap;font-size:15px">' + esc(vtRecuLignes(r).join('\n')) + '</div>' +
      '<label for="vtWa">Numéro WhatsApp du client (facultatif)</label><input id="vtWa" type="tel" inputmode="tel" placeholder="+509" autocomplete="off">' +
      '<div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap"><button onclick="vtRecuWa()">📲 Envoyer par WhatsApp</button>' +
      (WEB ? '' : '<button class="sec" onclick="vtImprimer()">🖨 Imprimer</button>') + '</div>' +
      (WEB ? '' : '<p class="mut">Imprimante : ' + (imp ? esc(imp.nom || imp.address) + ' · <a href="#" onclick="vtChoisirImprimante();return false">changer</a>' : 'aucune choisie · <a href="#" onclick="vtChoisirImprimante();return false">choisir</a>') + '</p>');
    view(h, '<button class="sec" onclick="ecranVente()">Retour</button><button onclick="VT.msgT=0;ecranVente()">➕ Nouvelle vente</button>', 'recu');
  });
}
function vtRecuCourant() { return vtMiennes().then(function (l) { return l.filter(function (x) { return x.cid === VT.recuCid; })[0]; }); }
function vtRecuWa() {
  var num = String(($('vtWa') || {}).value || '').replace(/[^0-9]/g, '');
  if (num.length === 8) num = '509' + num;   // numéro haïtien sans indicatif
  vtRecuCourant().then(function (r) {
    if (!r) return;
    var url = 'https://wa.me/' + (num.length >= 10 ? num : '') + '?text=' + encodeURIComponent(vtRecuLignes(r).join('\n'));
    var a = document.createElement('a'); a.href = url; a.target = '_blank'; a.rel = 'noopener'; document.body.appendChild(a); a.click(); a.remove();
  });
}
/* Imprimante de reçus Bluetooth (ESC/POS, 58 mm). Le texte est envoyé sans accents : beaucoup d'imprimantes n'ont pas les lettres accentuées. */
function vtBT() { return (window.Capacitor && Capacitor.Plugins && Capacitor.Plugins.BluetoothSerial) || null; }
function vtChoisirImprimante() {
  var bt = vtBT();
  if (!bt) return alerte('Impression Bluetooth : pas disponible sur cette application. Il faut installer la nouvelle application (lien du bureau).');
  view('<h1>🖨 Imprimante</h1><p class="mut center" style="margin-top:20px">Recherche des appareils Bluetooth associés…</p>', '<button class="sec" onclick="vtEcranRecu(VT.recuCid)">Retour</button>', 'imprimante');
  bt.isEnabled().then(function (e) { return e.enabled ? null : bt.enable(); }).then(function () { return bt.scan(); }).then(function (res) {
    var l = ((res && res.devices) || []).filter(function (d) { return d.address; });
    VT.btListe = l;
    var h = '<h1>🖨 Imprimante</h1><p class="mut">Associez d\'abord l\'imprimante dans les réglages Bluetooth du téléphone, puis choisissez-la ici.</p>';
    if (!l.length) h += '<div class="wait">Aucun appareil trouvé. Allumez l\'imprimante et associez-la dans les réglages Bluetooth.</div>';
    l.forEach(function (d, i) { h += '<button class="bigbtn sec" onclick="vtImprimanteChoisie(' + i + ')">' + esc(d.name || d.address) + '<div class="mut">' + esc(d.address) + '</div></button>'; });
    view(h, '<button class="sec" onclick="vtEcranRecu(VT.recuCid)">Retour</button>', 'imprimante');
  }).catch(function (e) { view('<h1>🖨 Imprimante</h1><div class="err">' + esc('Bluetooth : ' + ((e && (e.message || e.errorMessage)) || e || 'erreur')) + '</div>', '<button class="sec" onclick="vtEcranRecu(VT.recuCid)">Retour</button>', 'imprimante'); });
}
function vtImprimanteChoisie(i) { var d = (VT.btListe || [])[i]; if (!d) return; CFG.set('imprimante', { address: d.address, nom: d.name || '' }); vtEcranRecu(VT.recuCid); }
function vtOctets(lignes) {
  var sans = function (t) { return String(t).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/°/g, 'o').replace(/[^\x20-\x7E]/g, ' '); };
  var o = [27, 64], txt = function (t) { for (var i = 0; i < t.length; i++) o.push(t.charCodeAt(i)); o.push(10); };
  o.push(27, 97, 1, 27, 69, 1); txt(sans(lignes[0])); o.push(27, 69, 0); txt(sans(lignes[1])); o.push(27, 97, 0);
  for (var i = 2; i < lignes.length; i++) {
    if (/^TOTAL/.test(lignes[i])) { o.push(27, 69, 1, 29, 33, 17); txt(sans(lignes[i])); o.push(29, 33, 0, 27, 69, 0); } else txt(sans(lignes[i]));
  }
  o.push(10, 10, 10, 29, 86, 66, 0);   // avance et coupe (sans effet si l'imprimante n'a pas de couteau)
  return o;
}
function vtImprimer() {
  var bt = vtBT(), imp = CFG.get('imprimante', null);
  if (!bt) return alerte('Impression Bluetooth : pas disponible sur cette application. Il faut installer la nouvelle application (lien du bureau).');
  if (!imp) return vtChoisirImprimante();
  vtRecuCourant().then(function (r) {
    if (!r) return;
    toast('Impression…');
    var a = { address: imp.address };
    return bt.connect(a).catch(function () { return bt.connectInsecure(a); })
      .then(function () { return bt.write({ address: imp.address, value: vtOctets(vtRecuLignes(r)) }); })
      .then(function () { return new Promise(function (res) { setTimeout(res, 1500); }); })
      .then(function () { return bt.disconnect(a).catch(function () { }); })
      .then(function () { toast('Reçu imprimé.'); });
  }).catch(function (e) { alerte('Impression impossible : ' + ((e && (e.message || e.errorMessage)) || 'imprimante éteinte ou trop loin') + '.'); });
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
    var tape = document.activeElement && /^vt(AutreIn|Client|Recu)$/.test(document.activeElement.id || '');   // on n'efface pas un chiffre en train d'être tapé
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
/* Contrôle du jour avant d'envoyer le rapport : gallons du compteur contre gallons notés en ventes, et l'argent des ventes. */
function vtControleHtml(date, galCompteur) {
  if (!venteOn() || !(galCompteur >= 0)) return Promise.resolve('');
  return vtMiennes().then(function (l) {
    var j = vtDuJour(l, date), manque = Math.round((galCompteur - j.gallons) * 100) / 100;
    var cls = Math.abs(manque) <= Math.max(2, 0.02 * galCompteur) ? 'okb' : 'wait';
    return '<div class="' + cls + '">🧾 Ventes notées : ' + fmt(j.gallons) + ' gal (' + j.ventes.length + ' vente(s)) · Compteur : ' + fmt(galCompteur) + ' gal' +
      (manque > 0 ? ' · ' + fmt(manque) + ' gal sans vente notée' : manque < 0 ? ' · ' + fmt(-manque) + ' gal de plus que le compteur' : '') +
      (j.montant ? '<br>Argent des ventes : cash ' + fmt(j.cash_montant) + ' · crédit ' + fmt(j.credit_montant) + ' ' + esc(DEV) + (j.sans_prix ? ' · ' + j.sans_prix + ' vente(s) sans prix' : '') : '') + '</div>';
  }, function () { return ''; });
}
