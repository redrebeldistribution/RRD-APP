/**
 * RRD Notifications — relais des notifications poussées vers les téléphones (Firebase Cloud Messaging).
 *
 * Programme à part du rapport de fermeture : s'il s'arrête, rien n'est perdu, les messages s'affichent à la prochaine
 * ouverture de l'application comme avant.
 *
 * Ce qu'il fait, toutes les 5 minutes (et tout de suite quand la direction écrit un message pour maintenant) :
 *   1. lit l'onglet « Messages (app) » : les messages actifs dont l'heure est arrivée et qui n'ont pas encore été traités ;
 *   2. lit l'onglet « Appareils » : les téléphones de station qui ont une adresse de notification (colonne « push ») ;
 *   3. pousse le message aux téléphones visés qui ne l'ont pas encore pris (« msg_vu_le » plus ancien que le message) —
 *      les autres l'ont déjà, programmé sur le téléphone ;
 *   4. note « pousse_le » et « pousse_a » sur la ligne du message. Rien d'autre n'est modifié, aucune ligne n'est supprimée.
 *
 * Il n'écrit que dans ces deux onglets : « pousse_le », « pousse_a », et la colonne « push » (vidée quand Firebase dit que
 * l'adresse n'existe plus).
 */
var SHEET_ID = '__SHEET_ID__';   // classeur privé du rapport de fermeture (jamais dans le dépôt public)
var PROJET = 'rrd-app-94e2c';
var TZ = 'America/Port-au-Prince';
var ONGLET_MSG = 'Messages (app)', ONGLET_APP = 'Appareils', CANAL = 'rrd-messages';

function doPost(e) {
  var b = {}; try { b = JSON.parse((e && e.postData && e.postData.contents) || '{}') || {}; } catch (x) { }
  if (b.action === 'chatMembres') return sortie_(chatMembres_());
  if (b.action === 'chatAlertes') return sortie_(chatAlertes_());
  return sortie_(traiter_());
}
function doGet() { return sortie_({ ok: true, service: 'RRD Notifications' }); }
function sortie_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function toutesLes5Minutes() { traiter_(); try { chatMembres_(); chatAlertes_(); } catch (e) { console.error(e && e.stack || e); } }

/** À lancer une fois : installe le passage toutes les 5 minutes. */
function installer() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'toutesLes5Minutes') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('toutesLes5Minutes').timeBased().everyMinutes(5).create();
  Logger.log('Installé. ' + verifier());
}

/** Vérifie l'accès à Firebase sans rien envoyer à personne (essai « à blanc » avec une adresse qui n'existe pas). */
function verifier() {
  var c = corps_('adresse-inexistante', { id: 0, type: 'message', texte: 'Essai' }); c.validate_only = true;
  var r = fcm_(c);
  var ok = r.code === 400 && /not a valid FCM registration token/.test(r.texte);   // Firebase a bien lu toute la demande et ne refuse que l'adresse
  var t = (ok ? 'Accès à Firebase : OK' : 'Accès à Firebase : REFUSÉ') + ' (réponse ' + r.code + ', mode ' + r.mode + ') ' + r.texte.slice(0, 300);
  Logger.log(t);
  return t;
}

function entetes_(sh, aAjouter) {
  var h = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(function (x) { return String(x).trim(); });
  aAjouter.forEach(function (c) { if (h.indexOf(c) < 0) { sh.getRange(1, h.length + 1).setValue(c); h.push(c); } });
  var m = {}; h.forEach(function (c, i) { if (c && m[c] === undefined) m[c] = i; });
  return m;
}

/** La notification telle qu'elle est envoyée à Firebase (la même pour la vérification et pour les vrais envois). */
function corps_(adresse, m) {
  return { message: { token: adresse,
    notification: { title: m.type === 'conseil' ? '💡 Conseil du jour' : '✉️ Message de la direction', body: m.texte },
    data: { msg_id: String(m.id), type: String(m.type) },
    android: { priority: 'HIGH', ttl: '86400s', notification: { channel_id: CANAL, tag: 'msg-' + m.id, default_sound: true, notification_priority: 'PRIORITY_HIGH' } } } };
}

function fcm_(corps) {
  var props = PropertiesService.getScriptProperties(), prefere = props.getProperty('MODE') || 'projet';
  var modes = prefere === 'simple' ? ['simple', 'projet'] : ['projet', 'simple'], dernier = null;
  for (var i = 0; i < modes.length; i++) {
    var entetes = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
    if (modes[i] === 'projet') entetes['X-Goog-User-Project'] = PROJET;
    var r = UrlFetchApp.fetch('https://fcm.googleapis.com/v1/projects/' + PROJET + '/messages:send',
      { method: 'post', contentType: 'application/json', headers: entetes, payload: JSON.stringify(corps), muteHttpExceptions: true });
    dernier = { code: r.getResponseCode(), texte: String(r.getContentText() || ''), mode: modes[i] };
    if (dernier.code !== 401 && dernier.code !== 403) { if (modes[i] !== prefere) props.setProperty('MODE', modes[i]); return dernier; }
  }
  return dernier;
}

/** « TOUS » = toutes les stations (pas les camions) ; « CAMIONS » = tous les camions ; sinon un code exact (station ou « CAM:1 »). */
function pourTel_(cible, station) {
  cible = String(cible).trim().toUpperCase();
  if (station.indexOf('CAM:') === 0) return cible === 'CAMIONS' || cible === station;
  return cible === 'TOUS' || cible === station;
}
function traiter_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return { ok: true, occupe: true };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID), fz = ss.getSpreadsheetTimeZone();
    var shM = ss.getSheetByName(ONGLET_MSG);
    if (!shM || shM.getLastRow() < 2) return { ok: true, messages: 0, envoyes: 0 };
    var txt = function (v, f) {
      return v instanceof Date ? Utilities.formatDate(v, fz, f) : String(v == null ? '' : v).trim().replace(/^(\d{4}-\d{2}-\d{2}) (\d):/, '$1 0$2:');
    };
    var hM = entetes_(shM, ['pousse_le', 'pousse_a']);
    ['id', 'cree_le', 'cible', 'date', 'heure', 'type', 'texte', 'actif'].forEach(function (c) { if (hM[c] === undefined) throw new Error('Colonne absente dans « ' + ONGLET_MSG + ' » : ' + c); });
    var n = shM.getLastRow(), de = Math.max(2, n - 499);
    var lignes = shM.getRange(de, 1, n - de + 1, shM.getLastColumn()).getValues();
    var now = new Date(), maint = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm'), hier = Utilities.formatDate(new Date(now.getTime() - 864e5), TZ, 'yyyy-MM-dd HH:mm');
    var dus = [];
    lignes.forEach(function (r, i) {
      if (!/^(oui|true)$/i.test(String(r[hM.actif]).trim()) || String(r[hM.pousse_le] == null ? '' : r[hM.pousse_le]).trim()) return;
      var date = txt(r[hM.date], 'yyyy-MM-dd').slice(0, 10), heure = txt(r[hM.heure], 'HH:mm').replace('h', ':').slice(0, 5);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^\d{2}:\d{2}$/.test(heure)) return;
      var quand = date + ' ' + heure;
      if (quand > maint) return;   // pas encore l'heure
      dus.push({ ligne: de + i, id: Number(r[hM.id]) || 0, cree: txt(r[hM.cree_le], 'yyyy-MM-dd HH:mm:ss'), cible: String(r[hM.cible]).trim(),
        type: String(r[hM.type]).trim(), texte: String(r[hM.texte]).trim(), tropTard: quand < hier });
    });
    if (!dus.length) return { ok: true, messages: 0, envoyes: 0 };

    var shA = ss.getSheetByName(ONGLET_APP), tel = [];
    if (shA && shA.getLastRow() >= 2) {
      var hA = entetes_(shA, []);
      if (hA.push !== undefined && hA.station !== undefined) {
        shA.getRange(2, 1, shA.getLastRow() - 1, shA.getLastColumn()).getValues().forEach(function (r, i) {
          var st = String(r[hA.station]).trim().toUpperCase(), jeton = String(r[hA.push] || '').trim();
          if (!jeton || !st || st === 'DIRECTION' || st === 'COMPTA' || st.indexOf('TR:') === 0) return;   // conseils et messages : stations et camions (27/11/2026)
          if (hA.actif !== undefined && !/^(oui|true|yes|1)$/i.test(String(r[hA.actif]).trim())) return;
          if (hA.jeton !== undefined && !String(r[hA.jeton] || '').trim()) return;   // code libéré
          tel.push({ ligne: i + 2, station: st, push: jeton, vu: hA.msg_vu_le === undefined ? '' : txt(r[hA.msg_vu_le], 'yyyy-MM-dd HH:mm:ss'), col: hA.push + 1 });
        });
      }
    }

    var envoyes = 0, erreurs = 0, traites = 0, stop = '';
    for (var k = 0; k < dus.length && !stop; k++) {
      var m = dus[k], recus = [];
      if (!m.tropTard && m.texte) {
        var cibles = tel.filter(function (t) { return pourTel_(m.cible, t.station) && t.vu && (!m.cree || t.vu < m.cree); });   // t.vu vide : le téléphone n'a encore jamais pris ses messages, il le fait à l'ouverture
        for (var j = 0; j < cibles.length; j++) {
          var t = cibles[j], r = fcm_(corps_(t.push, m));
          if (r.code === 200) { envoyes++; if (recus.indexOf(t.station) < 0) recus.push(t.station); }
          else if (r.code === 401 || r.code === 403) { stop = 'Firebase refuse l\'accès (' + r.code + ') : ' + r.texte.slice(0, 300); break; }   // rien n'est marqué : nouvel essai au prochain passage
          else {
            erreurs++;
            if (/UNREGISTERED/.test(r.texte) || (r.code === 400 && /not a valid FCM registration token/.test(r.texte))) { shA.getRange(t.ligne, t.col).setValue(''); t.push = ''; }   // l'application a été désinstallée
            console.warn('Message ' + m.id + ' → ' + t.station + ' : ' + r.code + ' ' + r.texte.slice(0, 200));
          }
        }
        tel = tel.filter(function (x) { return x.push; });
      }
      if (stop) break;
      shM.getRange(m.ligne, hM.pousse_le + 1).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
      shM.getRange(m.ligne, hM.pousse_a + 1).setValue(recus.join(', '));
      SpreadsheetApp.flush();
      traites++;
    }
    var props = PropertiesService.getScriptProperties();
    if (stop) { props.setProperty('DERNIERE_ERREUR', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') + ' ' + stop); console.error(stop); }
    else if (props.getProperty('DERNIERE_ERREUR')) props.deleteProperty('DERNIERE_ERREUR');
    return { ok: !stop, messages: traites, envoyes: envoyes, erreurs: erreurs };
  } catch (e) {
    console.error(e && e.stack || e);
    return { ok: false };
  } finally { lock.releaseLock(); }
}

/* ============================ Chat : liste des membres (27/11/2026) ============================
 * Le serveur du rapport de fermeture note dans l'onglet « Chat (membres) » quel téléphone (identité Firebase anonyme « uid »)
 * parle sous quel nom. Ce relais recopie cette liste dans la base Firebase (« membres/<uid> »), avec le compte Google du propriétaire
 * (pas de clé secrète), et retire les téléphones désactivés ou libérés dans l'onglet « Appareils ». La base n'accepte que ses membres.
 * Il n'écrit dans le classeur que les colonnes « base » et « base_le » de cet onglet. */
var ONGLET_CHM = 'Chat (membres)', BASE_CHAT = 'https://' + PROJET + '-default-rtdb.firebaseio.com/';
function chatMembres_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return { ok: true, occupe: true };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID), sh = ss.getSheetByName(ONGLET_CHM);
    if (!sh || sh.getLastRow() < 2) return { ok: true, membres: 0 };
    var h = entetes_(sh, ['base', 'base_le']);
    ['uid', 'code', 'nom', 'role', 'appareil', 'actif'].forEach(function (c) { if (h[c] === undefined) throw new Error('Colonne absente dans « ' + ONGLET_CHM + ' » : ' + c); });
    var actifs = {}, shA = ss.getSheetByName(ONGLET_APP);
    if (shA && shA.getLastRow() >= 2) {
      var hA = entetes_(shA, []);
      shA.getRange(2, 1, shA.getLastRow() - 1, shA.getLastColumn()).getValues().forEach(function (r) {
        var c = String(r[hA.code_activation]).trim(); if (!c) return;
        actifs[c] = (hA.actif === undefined || /^(oui|true|yes|1)$/i.test(String(r[hA.actif]).trim())) && (hA.jeton === undefined || String(r[hA.jeton] || '').trim() !== '');
      });
    }
    var lignes = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues(), faits = 0, stop = '';
    var entetesHttp = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
    for (var i = 0; i < lignes.length && !stop; i++) {
      var r = lignes[i], uid = String(r[h.uid]).trim(), code = String(r[h.code]).trim();
      if (!/^[A-Za-z0-9]{20,40}$/.test(uid) || !/^[A-Z0-9-]{2,30}$/.test(code)) continue;
      var voulu = /^(oui|true)$/i.test(String(r[h.actif]).trim()) && actifs[String(r[h.appareil]).trim()] === true, etat = String(r[h.base]).trim();
      if (voulu === (etat === 'present')) continue;
      var rep = UrlFetchApp.fetch(BASE_CHAT + 'membres/' + uid + '.json', voulu
        ? { method: 'put', contentType: 'application/json', headers: entetesHttp, muteHttpExceptions: true, payload: JSON.stringify({ code: code, nom: String(r[h.nom]).slice(0, 60), role: String(r[h.role]) }) }
        : { method: 'delete', headers: entetesHttp, muteHttpExceptions: true });
      if (rep.getResponseCode() === 200) {
        sh.getRange(i + 2, h.base + 1).setValue(voulu ? 'present' : 'retire');
        sh.getRange(i + 2, h.base_le + 1).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
        faits++;
      } else stop = 'Base du chat : réponse ' + rep.getResponseCode() + ' ' + String(rep.getContentText() || '').slice(0, 200);
    }
    if (stop) console.error(stop);
    return { ok: !stop, changes: faits, erreur: stop };
  } finally { lock.releaseLock(); }
}

/** À lancer une fois depuis l'éditeur après l'ajout du chat : Google demande l'autorisation d'écrire dans la base, puis on vérifie l'accès. */
function verifierChat() {
  var r = UrlFetchApp.fetch(BASE_CHAT + 'membres.json?shallow=true', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  var t = 'Accès à la base du chat : ' + (r.getResponseCode() === 200 ? 'OK' : 'REFUSÉ') + ' (réponse ' + r.getResponseCode() + ')';
  Logger.log(t + ' — ' + JSON.stringify(chatMembres_()));
  return t;
}

/* Notifications du chat (27/11/2026). Le serveur du rapport note une ligne par message signalé dans « Chat (alertes) » (qui, à qui,
 * quelle conversation — jamais le texte). Ici : on pousse « <nom> vous a écrit » aux téléphones visés qui ont une adresse de notification,
 * puis on note « pousse_le » et « pousse_a ». Une alerte de plus d'une heure n'est plus envoyée (seulement notée). */
var ONGLET_CHA = 'Chat (alertes)', CANAL_CHAT = 'rrd-chat';
function codeChat_(station, activation) {   // le nom de chat d'un téléphone, comme le serveur le donne
  var s = String(station || '').trim().toUpperCase();
  if (s === 'DIRECTION' || s === 'COMPTA') return String(activation || '').toUpperCase().replace(/[^A-Z0-9-]/g, '');
  if (s.indexOf('TR:') === 0) return '';
  if (s.indexOf('CAM:') === 0) return 'CAM-' + s.slice(4);
  return s.replace(/[^A-Z0-9-]/g, '');
}
function chatAlertes_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return { ok: true, occupe: true };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID), sh = ss.getSheetByName(ONGLET_CHA);
    if (!sh || sh.getLastRow() < 2) return { ok: true, alertes: 0 };
    var h = entetes_(sh, ['pousse_le', 'pousse_a']), n = sh.getLastRow(), de = Math.max(2, n - 199);
    var lignes = sh.getRange(de, 1, n - de + 1, sh.getLastColumn()).getValues(), fz = ss.getSpreadsheetTimeZone();
    var dus = [];
    lignes.forEach(function (r, i) { if (!String(r[h.pousse_le] == null ? '' : r[h.pousse_le]).trim()) dus.push({ ligne: de + i, r: r }); });
    if (!dus.length) return { ok: true, alertes: 0 };
    var tel = [], shA = ss.getSheetByName(ONGLET_APP);
    if (shA && shA.getLastRow() >= 2) {
      var hA = entetes_(shA, []);
      shA.getRange(2, 1, shA.getLastRow() - 1, shA.getLastColumn()).getValues().forEach(function (r, i) {
        var push = hA.push === undefined ? '' : String(r[hA.push] || '').trim(), act = String(r[hA.code_activation] || '').trim();
        if (!push || (hA.actif !== undefined && !/^(oui|true|yes|1)$/i.test(String(r[hA.actif]).trim())) || (hA.jeton !== undefined && !String(r[hA.jeton] || '').trim())) return;
        var code = codeChat_(r[hA.station], act); if (code) tel.push({ ligne: i + 2, code: code, act: act, push: push, col: hA.push + 1 });
      });
    }
    var maint = new Date(), envoyes = 0, stop = '';
    for (var k = 0; k < dus.length && !stop; k++) {
      var r = dus[k].r, conv = String(r[h.conv]), cible = String(r[h.cible]), deAct = String(r[h.appareil]), nom = String(r[h.de_nom]), recus = [];
      var le = r[h.le] instanceof Date ? r[h.le] : new Date(String(r[h.le]).replace(' ', 'T') + '-05:00');
      var vieux = !(le instanceof Date) || isNaN(le) || maint - le > 3600e3;
      if (!vieux) {
        var cibles = tel.filter(function (x) { return x.act !== deAct && (cible === 'TOUS' || x.code === cible); });
        for (var j = 0; j < cibles.length; j++) {
          var x = cibles[j], rep = fcm_({ message: { token: x.push,
            notification: { title: '💬 ' + nom, body: conv === 'general' ? 'Nouveau message dans « Général »' : 'Vous a écrit un message' },
            data: { chat: conv },
            android: { priority: 'HIGH', ttl: '3600s', notification: { channel_id: CANAL_CHAT, tag: 'chat-' + conv, default_sound: true, notification_priority: 'PRIORITY_HIGH' } } } });
          if (rep.code === 200) { envoyes++; if (recus.indexOf(x.code) < 0) recus.push(x.code); }
          else if (rep.code === 401 || rep.code === 403) { stop = 'Firebase refuse l\'accès (' + rep.code + ')'; break; }
          else if (/UNREGISTERED/.test(rep.texte) || (rep.code === 400 && /not a valid FCM registration token/.test(rep.texte))) { shA.getRange(x.ligne, x.col).setValue(''); x.push = ''; }
        }
        tel = tel.filter(function (y) { return y.push; });
      }
      if (stop) break;
      sh.getRange(dus[k].ligne, h.pousse_le + 1).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
      sh.getRange(dus[k].ligne, h.pousse_a + 1).setValue(vieux ? '(trop ancien, pas envoyé)' : recus.join(', '));
    }
    if (stop) console.error(stop);
    return { ok: !stop, alertes: dus.length, envoyes: envoyes };
  } finally { lock.releaseLock(); }
}

