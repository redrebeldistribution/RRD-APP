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

var MOBILE_VERSION = 'm1.1';
var ROLES_APPAREIL = { DIRECTION: 'management', COMPTA: 'accounting' };   // station = DIRECTION ou COMPTA → téléphone de la direction / comptabilité
var JOURS_DIFFERE_MAX = 3;   // un rapport fait hors connexion peut arriver jusqu'à 3 jours plus tard

var APPAREILS = {
  name: 'Appareils',
  headers: ['code_activation', 'station', 'actif', 'appareil_id', 'jeton', 'modele', 'enregistre_le', 'dernier_contact', 'note', 'nom']
};

function installerMobile() {
  var ss = ss_();
  var sh = ss.getSheetByName(APPAREILS.name);
  if (!sh) {
    sh = ss.insertSheet(APPAREILS.name);
    sh.appendRow(APPAREILS.headers);
    sh.setFrozenRows(1);
    sh.getRange('A:I').setNumberFormat('@');
  } else {
    var h = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
    APPAREILS.headers.forEach(function (c) { if (h.indexOf(c) < 0) { sh.getRange(1, sh.getLastColumn() + 1).setValue(c); h.push(c); } });
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
      case 'dir': return ok_(direction_(app, String(b.sous_action || ''), payload));
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
    var role = ROLES_APPAREIL[String(ligne.station).trim().toUpperCase()] || '';
    var st = role ? null : stations_().filter(function (s) { return String(s.code) === String(ligne.station).trim() && estVrai_(s.actif); })[0];
    if (!role && !st) throw new UserError('La station de ce code n\'existe pas ou est inactive.', 'CODE');
    var jeton = jetonAleatoire_();
    majCellule_(APPAREILS, ligne._ligne, 'jeton', jeton);
    majCellule_(APPAREILS, ligne._ligne, 'appareil_id', String(b.appareil_id || '').slice(0, 64));
    majCellule_(APPAREILS, ligne._ligne, 'modele', String(b.modele || '').slice(0, 80));
    majCellule_(APPAREILS, ligne._ligne, 'enregistre_le', maintenant_());
    journal_({ id: 'APP', nom: 'Appareil ' + code }, 'appareil+', role || st.code);
    if (role) return { jeton: jeton, role: role, nom: String(ligne.nom || ligne.note || code) };
    return { jeton: jeton, station: { code: st.code, nom: st.nom } };
  } finally { lock.releaseLock(); }
}

function appareil_(jeton) {
  jeton = String(jeton || '').trim();
  if (!jeton) throw new UserError('Téléphone pas encore enregistré.', 'NON_ENREGISTRE');
  var a = lire_(APPAREILS).filter(function (x) { return String(x.jeton) === jeton; })[0];
  if (!a) throw new UserError('Ce téléphone n\'est plus enregistré. Demandez un nouveau code à la direction.', 'NON_ENREGISTRE');
  if (!estVrai_(a.actif)) throw new UserError('Ce téléphone a été désactivé par la direction.', 'DESACTIVE');
  var role = ROLES_APPAREIL[String(a.station).trim().toUpperCase()] || '';
  if (role) {
    try { majCellule_(APPAREILS, a._ligne, 'dernier_contact', maintenant_()); } catch (e) { }
    return { ligne: a, station: null, role: role, id: 'APP-' + String(a.code_activation), nom: String(a.nom || a.note || a.code_activation), username: 'app', enGroupe: false, chatInstance: '' };
  }
  var st = stations_().filter(function (s) { return String(s.code) === String(a.station).trim() && estVrai_(s.actif); })[0];
  if (!st) throw new UserError('Station inactive.', 'DESACTIVE');
  try { majCellule_(APPAREILS, a._ligne, 'dernier_contact', maintenant_()); } catch (e) { }
  return { ligne: a, station: st, id: 'APP-' + String(a.code_activation), nom: 'Téléphone ' + st.code };
}

/* ============================ Écran de départ ============================ */

function initMobile_(app) {
  if (app.role) return initDirection_(app);
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
  if (app.role) throw new UserError('Ce téléphone est celui de la direction : il ne fait pas de rapport de station.');
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


/* ============================ Direction / comptabilité ============================ */

function initDirection_(app) {
  return {
    version: MOBILE_VERSION, devise: CURRENCY, mode: 'tableau',
    utilisateur: { id: app.id, nom: app.nom, role: app.role },
    stations: stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return { code: s.code, nom: s.nom }; }),
    aujourdhui: dateRapport(),
    champs: [{ k: 'employes_service', label: 'Employé(s) de service', texte: true }].concat(FIELDS)
      .concat([{ k: 'other_expenses_detail', label: 'Détail autres dépenses', texte: true }, { k: 'notes', label: 'Notes', texte: true }])
  };
}

/** Mêmes fonctions que le tableau de bord Telegram, avec l'identité du téléphone de la direction. */
function direction_(app, sous, p) {
  if (!app.role) throw new UserError('Accès réservé à la direction.', 'FORBIDDEN');
  var ctx = { id: app.id, nom: app.nom, username: 'app', role: app.role, chatType: 'private', chatInstance: '', enGroupe: false };
  switch (sous) {
    case 'rapports': return listeRapports_(ctx, p);
    case 'detail': return detailRapport_(ctx, p);
    case 'corriger': return corriger_(ctx, p);
    case 'annuler': return annuler_(ctx, p);
    case 'resume': return resumeJour_(ctx, p);
    case 'prix': return listePrix_(ctx);
    case 'ajouterPrix': return ajouterPrix_(ctx, p);
    case 'appareils': exiger_(ctx, ['management']); return listeAppareils_();
    case 'ajouterAppareil': exiger_(ctx, ['management']); return ajouterAppareil_(ctx, p);
    case 'activerAppareil': exiger_(ctx, ['management']); return activerAppareil_(ctx, p);
    default: throw new UserError('Action inconnue.');
  }
}

function listeAppareils_() {
  var noms = {}; stations_().forEach(function (s) { noms[String(s.code)] = s.nom; });
  noms.DIRECTION = 'Direction'; noms.COMPTA = 'Comptabilité';
  return {
    stations: stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return { code: s.code, nom: s.nom }; }),
    appareils: lire_(APPAREILS).map(function (a) {
      return { code: String(a.code_activation), station: String(a.station), station_nom: noms[String(a.station).trim()] || String(a.station),
        actif: estVrai_(a.actif), utilise: !!String(a.jeton || '').trim(), enregistre_le: String(a.enregistre_le || ''),
        dernier_contact: String(a.dernier_contact || ''), nom: String(a.nom || '') };
    })
  };
}

function ajouterAppareil_(ctx, p) {
  var station = String(p.station || '').trim().toUpperCase();
  var valides = stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return String(s.code); }).concat(['DIRECTION', 'COMPTA']);
  if (valides.indexOf(station) < 0) throw new UserError('Station inconnue.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var existants = lire_(APPAREILS).map(function (a) { return String(a.code_activation).trim().toUpperCase(); });
    var prefixe = station === 'DIRECTION' ? 'DIR' : station === 'COMPTA' ? 'CPT' : station, code;
    do { code = prefixe + '-' + (1000 + Math.floor(Math.random() * 9000)); } while (existants.indexOf(code) >= 0);
    ajouter_(APPAREILS, { code_activation: code, station: station, actif: 'oui', note: 'créé par ' + ctx.nom + ' le ' + maintenant_(), nom: String(p.nom || '').trim().slice(0, 60) });
    journal_(ctx, 'appareil-code', code + ' → ' + station);
  } finally { lock.releaseLock(); }
  var l = listeAppareils_(); l.nouveau = code; return l;
}

function activerAppareil_(ctx, p) {
  var a = lire_(APPAREILS).filter(function (x) { return String(x.code_activation) === String(p.code); })[0];
  if (!a) throw new UserError('Code introuvable.');
  if (String(a.code_activation) === ctx.id.replace(/^APP-/, '') && !p.actif) throw new UserError('Vous ne pouvez pas désactiver votre propre téléphone.');
  majCellule_(APPAREILS, a._ligne, 'actif', p.actif ? 'oui' : 'non');
  journal_(ctx, p.actif ? 'appareil-on' : 'appareil-off', String(a.code_activation));
  return listeAppareils_();
}
