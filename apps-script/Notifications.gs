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

function doPost() { return sortie_(traiter_()); }
function doGet() { return sortie_({ ok: true, service: 'RRD Notifications' }); }
function sortie_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function toutesLes5Minutes() { traiter_(); }

/** À lancer une fois : installe le passage toutes les 5 minutes. */
function installer() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'toutesLes5Minutes') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('toutesLes5Minutes').timeBased().everyMinutes(5).create();
  Logger.log('Installé. ' + verifier());
}

/** Vérifie l'accès à Firebase sans rien envoyer à personne (essai « à blanc » avec une adresse qui n'existe pas). */
function verifier() {
  var r = fcm_({ validate_only: true, message: { token: 'adresse-inexistante', notification: { title: 'Essai', body: 'Essai' } } });
  var ok = r.code === 400 && /INVALID_ARGUMENT/.test(r.texte);   // Firebase a bien lu la demande et refuse seulement l'adresse
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
          if (!jeton || !st || st === 'DIRECTION' || st === 'COMPTA' || st.indexOf('TR:') === 0) return;
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
        var cibles = tel.filter(function (t) { return (m.cible === 'TOUS' || m.cible.toUpperCase() === t.station) && t.vu && (!m.cree || t.vu < m.cree); });   // t.vu vide : le téléphone n'a encore jamais pris ses messages, il le fait à l'ouverture
        for (var j = 0; j < cibles.length; j++) {
          var t = cibles[j], r = fcm_({ message: { token: t.push,
            notification: { title: m.type === 'conseil' ? '💡 Conseil du jour' : '✉️ Message de la direction', body: m.texte },
            data: { msg_id: String(m.id), type: m.type },
            android: { priority: 'HIGH', ttl: '86400s', notification: { channel_id: CANAL, tag: 'msg-' + m.id, default_sound: true, notification_priority: 'PRIORITY_HIGH' } } } });
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
