/**
 * RRD – Rapport de fermeture : accès pour l'application Android (hors Telegram).
 * À AJOUTER comme nouveau fichier « Mobile » dans le projet Apps Script « RRD – Rapport de fermeture ».
 * Ne modifie aucun autre fichier : réutilise validerRapport, rapportsDuJour_, recuEmploye_, enregistrerPhoto_, etc.
 *
 * Principe :
 *  - Chaque téléphone de station est enregistré UNE fois avec un code d'activation que la direction écrit
 *    dans l'onglet « Appareils » de la feuille privée (colonne code_activation + station + actif = oui).
 *  - Le téléphone reçoit alors un jeton secret ; chaque envoi est vérifié avec ce jeton.
 *  - Mettre actif = non dans « Appareils » coupe immédiatement le téléphone.
 *  - L'application peut préparer le rapport hors connexion ; la référence (ex. ESA-20261125-KXMP) est créée
 *    sur le téléphone et réutilisée à l'envoi : renvoyer le même rapport ne crée jamais de doublon.
 *
 * Installation (une seule fois) : exécuter installerMobile(), puis Déployer › Gérer les déploiements › ✏️ › Nouvelle version.
 */

var MOBILE_VERSION = 'm1.0';
var JOURS_DIFFERE_MAX = 3;   // un rapport fait hors connexion peut arriver jusqu'à 3 jours plus tard

var APPAREILS = {
  name: 'Appareils',
  headers: ['code_activation', 'station', 'actif', 'appareil_id', 'jeton', 'modele', 'enregistre_le', 'dernier_contact', 'note']
};

function installerMobile() {
  var ss = ss_();
  var sh = ss.getSheetByName(APPAREILS.name);
  if (!sh) {
    sh = ss.insertSheet(APPAREILS.name);
    sh.appendRow(APPAREILS.headers);
    sh.setFrozenRows(1);
    sh.getRange('A:I').setNumberFormat('@');
  }
  return 'Onglet « Appareils » prêt. Ajoutez une ligne par téléphone : code_activation (ex. ESA-4821), station (ex. ESA), actif = oui.';
}

/* ============================ Point d'entrée HTTP ============================ */

function doPost(e) {
  var rep;
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    rep = apiMobile_(String(b.action || ''), b);
  } catch (err) {
    console.error(err && err.stack || err);
    rep = { ok: false, erreur: 'Requête invalide.' };
  }
  return ContentService.createTextOutput(JSON.stringify(rep)).setMimeType(ContentService.MimeType.JSON);
}

function apiMobile_(action, b) {
  try {
    if (action === 'ping') return ok_({ version: MOBILE_VERSION, heure: maintenant_() });
    if (action === 'enregistrer') return ok_(enregistrerAppareil_(b));
    var app = appareil_(b.jeton);
    var payload = b.payload || {};
    switch (action) {
      case 'init': return ok_(initMobile_(app));
      case 'envoyer': return ok_(envoyerMobile_(app, payload));
      default: throw new UserError('Action inconnue.');
    }
  } catch (e) {
    if (e instanceof UserError) return { ok: false, erreur: e.message, code: e.code || '' };
    console.error(e && e.stack || e);
    return { ok: false, erreur: 'Erreur du serveur. Le rapport reste sur le téléphone et sera renvoyé.', code: 'SERVEUR' };
  }
}

/* ============================ Appareils ============================ */

function jetonAleatoire_() {
  var L = 'abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 40; i++) s += L.charAt(Math.floor(Math.random() * L.length));
  return s;
}

function enregistrerAppareil_(b) {
  var code = String(b.code || '').trim().toUpperCase();
  if (!code) throw new UserError('Entrez le code d\'activation donné par la direction.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ligne = lire_(APPAREILS).filter(function (a) { return String(a.code_activation).trim().toUpperCase() === code; })[0];
    if (!ligne || !estVrai_(ligne.actif)) throw new UserError('Code d\'activation inconnu ou désactivé.', 'CODE');
    if (String(ligne.jeton || '').trim()) throw new UserError('Ce code a déjà été utilisé sur un autre téléphone. Demandez un nouveau code à la direction.', 'CODE_UTILISE');
    var st = stations_().filter(function (s) { return String(s.code) === String(ligne.station).trim() && estVrai_(s.actif); })[0];
    if (!st) throw new UserError('La station de ce code n\'existe pas ou est inactive.', 'CODE');
    var jeton = jetonAleatoire_();
    majCellule_(APPAREILS, ligne._ligne, 'jeton', jeton);
    majCellule_(APPAREILS, ligne._ligne, 'appareil_id', String(b.appareil_id || '').slice(0, 64));
    majCellule_(APPAREILS, ligne._ligne, 'modele', String(b.modele || '').slice(0, 80));
    majCellule_(APPAREILS, ligne._ligne, 'enregistre_le', maintenant_());
    journal_({ id: 'APP', nom: 'Appareil ' + code }, 'appareil+', st.code);
    return { jeton: jeton, station: { code: st.code, nom: st.nom } };
  } finally { lock.releaseLock(); }
}

function appareil_(jeton) {
  jeton = String(jeton || '').trim();
  if (!jeton) throw new UserError('Téléphone pas encore enregistré.', 'NON_ENREGISTRE');
  var a = lire_(APPAREILS).filter(function (x) { return String(x.jeton) === jeton; })[0];
  if (!a) throw new UserError('Ce téléphone n\'est plus enregistré. Demandez un nouveau code à la direction.', 'NON_ENREGISTRE');
  if (!estVrai_(a.actif)) throw new UserError('Ce téléphone a été désactivé par la direction.', 'DESACTIVE');
  var st = stations_().filter(function (s) { return String(s.code) === String(a.station).trim() && estVrai_(s.actif); })[0];
  if (!st) throw new UserError('Station inactive.', 'DESACTIVE');
  try { majCellule_(APPAREILS, a._ligne, 'dernier_contact', maintenant_()); } catch (e) { }
  return { ligne: a, station: st, id: 'APP-' + String(a.code_activation), nom: 'Téléphone ' + st.code };
}

/* ============================ Écran de départ ============================ */

function initMobile_(app) {
  var st = app.station, date = dateRapport();
  var jour = rapportsDuJour_(st.code, date);
  var base = {
    version: MOBILE_VERSION, devise: CURRENCY, station: { code: st.code, nom: st.nom }, date: date,
    champs: FIELDS, envoyes: jour.map(recuEmploye_), periodes: periodesPossibles(jour), noms_connus: nomsConnus_(st.code)
  };
  var periode = base.periodes[0];
  var prec = periode === 'PM' ? (jour.filter(function (r) { return String(r.periode) === 'AM'; })[0] || null) : null;
  prec = prec || dernierRapport_(st.code, date);
  if (prec) base.prerempli = { opening_meter: num(prec.closing_meter), starting_cash: num(prec.counted_cash) };
  return base;
}

/* ============================ Envoi ============================ */

function dateValideMobile_(d) {
  d = String(d || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new UserError('Date du rapport invalide.');
  var auj = dateRapport();
  if (d > auj) throw new UserError('La date du téléphone est en avance. Corrigez l\'heure du téléphone.', 'DATE');
  var min = Utilities.formatDate(new Date(Date.now() - JOURS_DIFFERE_MAX * 24 * 3600 * 1000), TZ, 'yyyy-MM-dd');
  if (d < min) throw new UserError('Rapport trop ancien pour être envoyé par l\'application (plus de ' + JOURS_DIFFERE_MAX + ' jours). Prévenez le bureau.', 'TROP_VIEUX');
  return { date: d, differe: d !== auj };
}

function envoyerMobile_(app, p) {
  var st = app.station;
  var dv = dateValideMobile_(p.date);
  var date = dv.date;
  var val = validerRapport(p);
  if (val.erreurs.length) throw new UserError(val.erreurs.join('\n'), 'INVALIDE');
  var noms = decouperNoms(p.employes_service);
  if (!noms.length) throw new UserError('Employé(s) de service : écrivez le nom de chaque personne qui a travaillé.', 'INVALIDE');

  var refAttendue = new RegExp('^' + st.code + '-' + date.replace(/-/g, '') + '-[A-Z]{4}$');
  var ref = String(p.ref || '');
  if (!refAttendue.test(ref)) ref = refAleatoire(st.code, date);

  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    // Même rapport renvoyé (coupure pendant l'envoi) : on renvoie le reçu existant, sans doublon.
    var existant = rapportsEffectifs_().filter(function (r) { return String(r.ref) === ref; })[0];
    if (existant) {
      if (String(existant.employe_id) === app.id) return { recu: recuEmploye_(existant), deja: true };
      ref = refAleatoire(st.code, date);   // collision avec une autre référence : on en crée une nouvelle
    }
    var possibles = periodesPossibles(rapportsDuJour_(st.code, date));
    var attendue = String(p.periode === undefined || p.periode === null ? (possibles[0] || '') : p.periode);
    if (possibles.indexOf(attendue) < 0) {
      throw new UserError('Un rapport « ' + (LIBELLE_PERIODE[attendue] || attendue) + ' » existe déjà pour le ' + date.split('-').reverse().join('/') +
        '. Ce rapport n\'a pas été enregistré. Prévenez le bureau.', 'DEJA');
    }
    var v = val.valeurs;
    var notes = v.notes;
    if (dv.differe) notes = (notes ? notes + ' ' : '') + '[envoyé en différé par l\'application le ' + maintenant_() + ']';
    var photoId = p.photo ? enregistrerPhoto_(p.photo, ref) : '';
    var ligne = ajouter_(SHEETS.RAPPORTS, {
      ref: ref, station: st.code, date_rapport: date, soumis_le: maintenant_(), employe_id: app.id, employe: noms[0], username: 'app',
      opening_meter: v.opening_meter, closing_meter: v.closing_meter, gallons: v.gallons, starting_cash: v.starting_cash, counted_cash: v.counted_cash,
      bank_deposit: v.bank_deposit, other_cash_in: v.other_cash_in, generator_fuel: v.generator_fuel, withdrawals: v.withdrawals,
      other_expenses: v.other_expenses, other_expenses_detail: v.other_expenses_detail, notes: notes, photo_id: photoId,
      chat_instance: 'app:' + String(app.ligne.code_activation), gas_received: v.gas_received, periode: attendue, employes_service: noms.join(', ')
    });
    journal_(app, 'rapport-app', ref + ' ' + (attendue || 'journée') + (dv.differe ? ' (différé)' : '') + ' (ligne ' + ligne + ')');
    if (st.chat_id) {
      try {
        telegram_('sendMessage', { chat_id: String(st.chat_id), text: '✅ Rapport de fermeture reçu — ' + st.nom + (attendue ? ' (' + LIBELLE_PERIODE[attendue].toLowerCase() + ')' : '') +
          ', ' + date.split('-').reverse().join('/') + ', par ' + noms[0] + '. Réf. ' + ref });
      } catch (e) { }
    }
    var r = rapportsEffectifs_().filter(function (x) { return String(x.ref) === ref; })[0];
    var reste = periodesPossibles(rapportsDuJour_(st.code, date));
    return { recu: recuEmploye_(r), reste: reste.length ? reste[0] : null };
  } finally { lock.releaseLock(); }
}
