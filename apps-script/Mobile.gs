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

var MOBILE_VERSION = 'm1.39';
var TRANSFERTS_SHEET_ID = '1TmjlsDmogSWElqWNmPfON9S-dvouCbEgOHnWl8iPRnA';   // feuille des transferts (onglet « Agents transferts »)
var STATION_FORMATION = 'TST', MAX_APPAREILS_FORMATION = 50;   // les codes de la station de test servent à la formation : plusieurs téléphones avec le même code (propriétaire, 26/11/2026)
var MAX_APPAREILS_ROLE = 5;   // un code DIR ou CPT peut servir sur 5 appareils en même temps (demande du propriétaire, 26/11/2026)
var ROLES_APPAREIL = { DIRECTION: 'management', COMPTA: 'accounting' };   // station = DIRECTION ou COMPTA → téléphone de la direction / comptabilité
/* Contrôles avant d'accepter un rapport (chiffres « visiblement faux »). Réglables ici. */
var CTRL = {
  PLAFOND_MIN: 25000,       // HTG : un montant au-dessus du plafond de la station est signalé (plafond = max(PLAFOND_MIN, 4 × caisse habituelle))
  PLAFOND_FOIS: 4,
  GALLONS_MIN: 150,         // gallons : signalé au-dessus de max(GALLONS_MIN, 3 × gallons habituels)
  GALLONS_FOIS: 3,
  ECART_MIN: 2000,          // HTG : écart de caisse signalé s'il dépasse max(ECART_MIN, ECART_PCT × ventes)
  ECART_PCT: 0.30
};
var JOURS_DIFFERE_MAX = 7;   // une station peut faire un rapport manquant jusqu'à 7 jours en arrière (jamais modifier un rapport déjà envoyé)

var APPAREILS = {
  name: 'Appareils',
  headers: ['code_activation', 'station', 'actif', 'appareil_id', 'jeton', 'modele', 'enregistre_le', 'dernier_contact', 'note', 'nom', 'push', 'msg_vu_le']
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
    if (action === 'demandeStation') return ok_(demandeStation_(b.payload || {}));   // formulaire public du site redrebel.us (pas de code)
    if (action === 'honneur') return ok_(honneurPublic_());   // site redrebel.us : équipe et employé(e) du mois confirmés (pas de code)
    var app = appareil_(b.jeton);
    var payload = b.payload || {};
    switch (action) {
      case 'init': return ok_(initMobile_(app));
      case 'envoyer': return ok_(envoyerMobile_(app, payload));
      case 'photo': return ok_(photoHoraire_(app, payload));
      case 'photos': return ok_(photosStation_(app));
      case 'bordereau': return ok_(bordereau_(app, payload));
      case 'messages': return ok_(messagesStation_(app));
      case 'pushToken': return ok_(pushToken_(app, payload));
      case 'reponse': return ok_(reponseStation_(app, payload));
      case 'camion': return ok_(String(payload.genre) === 'rapport' ? camRapport_(app, payload) : camEnvoyer_(app, payload));
      case 'camSuggestions': return ok_(sugPourCamion_(app));
      case 'chatEntrer': return ok_(chatEntrer_(app, payload));
      case 'chatSignal': return ok_(chatSignal_(app, payload));
      case 'votes': return ok_(votesTel_(app));
      case 'vote': return ok_(vote_(app, payload));
      case 'accordPhoto': return ok_(accordPhoto_(app, payload));
      case 'dir': return ok_(direction_(app, String(b.sous_action || ''), payload));
      default: throw new UserError('Action inconnue.');
    }
  } catch (e) {
    if (e instanceof UserError) return { ok: false, erreur: e.message, code: e.code || '', controles: e.controles || undefined };
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
  // Sans identifiant d'appareil, on ne peut ni réserver un code de station pour un téléphone, ni compter les appareils d'un code partagé.
  if (!String(b.appareil_id || '').split('|').filter(function (x) { return x.trim(); }).length) throw new UserError('Activation impossible : identifiant du téléphone manquant. Fermez l\'application, rouvrez-la et réessayez.', 'CODE');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ligne = lire_(APPAREILS).filter(function (a) { return String(a.code_activation).trim().toUpperCase() === code; })[0];
    if (!ligne || !estVrai_(ligne.actif)) throw new UserError('Code d\'activation inconnu ou désactivé.', 'CODE');
    // Un code déjà utilisé ne peut resservir que sur le MÊME téléphone (ex. après « Changer de code » pour tester).
    // Les identifiants sont séparés par « | » (identifiant de l'app + identifiant fixe du téléphone) : un seul en commun suffit.
    var role = roleAppareil_(ligne.station);
    if (role === 'management' || role === 'accounting') return enregistrerPartage_(ligne, b, role, code);
    if (role === 'camion') {
      var cam = camDe_(ligne.station);
      if (!cam) throw new UserError('Le camion de ce code n\'existe pas ou est inactif.', 'CODE');
      if (b.web) throw new UserError('La version web est réservée à la direction et à la comptabilité. Pour un camion, installez l\'application Android.', 'CODE');
      if (cam.code === CAM_TEST) { var rp = enregistrerPartage_(ligne, b, 'camion', code, MAX_APPAREILS_FORMATION); return { jeton: rp.jeton, role: 'camion', nom: cam.nom, camion: { code: cam.code, nom: cam.nom } }; }
    }
    if (String(ligne.station).trim().toUpperCase() === STATION_FORMATION) return enregistrerFormation_(ligne, b, code);
    var ancien = String(ligne.appareil_id || '').split('|').filter(String), envoye = String(b.appareil_id || '').split('|').filter(String);
    var memeTel = envoye.some(function (x) { return ancien.indexOf(x) >= 0; });
    if (String(ligne.jeton || '').trim() && !memeTel) throw new UserError('Ce code a déjà été utilisé sur un autre téléphone. Demandez un nouveau code à la direction.', 'CODE_UTILISE');
    var st = role ? null : stations_().filter(function (s) { return String(s.code) === String(ligne.station).trim() && estVrai_(s.actif); })[0];
    if (!role && !st) throw new UserError('La station de ce code n\'existe pas ou est inactive.', 'CODE');
    if (!role && b.web) throw new UserError('La version web est réservée à la direction et à la comptabilité. Pour une station, installez l\'application Android.', 'CODE');
    var jeton = jetonAleatoire_();
    majCellule_(APPAREILS, ligne._ligne, 'jeton', jeton);
    majCellule_(APPAREILS, ligne._ligne, 'appareil_id', String(b.appareil_id || '').slice(0, 160));
    majCellule_(APPAREILS, ligne._ligne, 'modele', String(b.modele || '').slice(0, 80));
    majCellule_(APPAREILS, ligne._ligne, 'enregistre_le', maintenant_());
    journal_({ id: 'APP', nom: 'Appareil ' + code }, 'appareil+', role || st.code);
    if (role === 'camion') return { jeton: jeton, role: role, nom: cam.nom, camion: { code: cam.code, nom: cam.nom } };
    if (role) return { jeton: jeton, role: role, nom: String(ligne.nom || ligne.note || code) };
    return { jeton: jeton, station: { code: st.code, nom: st.nom } };
  } finally { lock.releaseLock(); }
}

/**
 * Codes DIR et CPT : le même code sert sur plusieurs appareils (MAX_APPAREILS_ROLE), sans déconnecter les autres.
 * Tous les appareils du code partagent le même jeton ; « appareil_id » garde la liste des appareils, séparés par « ; »
 * (chaque appareil = ses identifiants séparés par « | »). « Libérer le code » déconnecte tous les appareils d'un coup.
 */
function appListe_(v) { return String(v || '').split(';').map(function (x) { return x.trim(); }).filter(String); }
function enregistrerFormation_(ligne, b, code) {
  var st = stations_().filter(function (s) { return String(s.code) === STATION_FORMATION && estVrai_(s.actif); })[0];
  if (!st) throw new UserError('La station de ce code n\'existe pas ou est inactive.', 'CODE');
  if (b.web) throw new UserError('La version web est réservée à la direction et à la comptabilité. Pour une station, installez l\'application Android.', 'CODE');
  var r = enregistrerPartage_(ligne, b, 'formation', code, MAX_APPAREILS_FORMATION);
  return { jeton: r.jeton, station: { code: st.code, nom: st.nom } };
}
function enregistrerPartage_(ligne, b, role, code, max) {
  max = max || MAX_APPAREILS_ROLE;
  var appareils = appListe_(ligne.appareil_id), modeles = String(ligne.modele || '').split(' || ');   // un modèle par appareil, même ordre
  var envoye = String(b.appareil_id || '').slice(0, 80).split('|').filter(String);
  var deja = String(ligne.jeton || '').trim(), i = -1;
  appareils.forEach(function (a, k) { if (i < 0 && envoye.some(function (x) { return a.split('|').indexOf(x) >= 0; })) i = k; });
  if (envoye.length) {
    if (i < 0) {
      if (deja && appareils.length >= max) throw new UserError('Ce code est déjà utilisé sur ' + max + ' appareils. Demandez à la direction de le libérer ou de créer un autre code.', 'CODE_UTILISE');
      i = appareils.length; appareils.push(envoye.join('|'));
    } else appareils[i] = appareils[i].split('|').concat(envoye).filter(function (x, k, l) { return l.indexOf(x) === k; }).join('|');
    while (modeles.length < appareils.length) modeles.push('');
    modeles[i] = String(b.modele || '').replace(/\|\|/g, '|').slice(0, 80);
  }
  var jeton = deja || jetonAleatoire_();
  if (!deja) majCellule_(APPAREILS, ligne._ligne, 'jeton', jeton);
  majCellule_(APPAREILS, ligne._ligne, 'appareil_id', appareils.join(' ; '));
  majCellule_(APPAREILS, ligne._ligne, 'modele', modeles.slice(0, appareils.length).join(' || '));
  majCellule_(APPAREILS, ligne._ligne, 'enregistre_le', maintenant_());
  journal_({ id: 'APP', nom: 'Appareil ' + code }, 'appareil+', role + ' (' + appareils.length + '/' + max + ')');
  return { jeton: jeton, role: role, nom: String(ligne.nom || ligne.note || code) };
}

/** DIRECTION / COMPTA = tableau de bord ; « TR:CODE » = agent de transfert sans station (transferts seulement). */
function roleAppareil_(station) {
  station = String(station).trim().toUpperCase();
  return ROLES_APPAREIL[station] || (station.indexOf('TR:') === 0 ? 'transferts' : station.indexOf('CAM:') === 0 ? 'camion' : '');
}

/** Agents de transfert actifs qui ne sont pas des stations du rapport de fermeture : ils reçoivent un code « transferts seulement ». */
function agentsTransfertsSeuls_() {
  try {
    var sh = SpreadsheetApp.openById(TRANSFERTS_SHEET_ID).getSheetByName('Agents transferts');
    var st = stations_().map(function (s) { return String(s.code); }), alias = { ARC: 'A32', MAH: 'M83' };
    return sh.getDataRange().getValues().slice(1).filter(function (r) {
      var c = String(r[0]).trim().toUpperCase();
      return c && /^(oui|yes|true|1)$/i.test(String(r[7]).trim()) && st.indexOf(alias[c] || c) < 0;
    }).map(function (r) { return { code: 'TR:' + String(r[0]).trim().toUpperCase(), nom: String(r[1]).trim() + ' (transferts seulement)' }; });
  } catch (e) { return []; }
}

function appareil_(jeton) {
  jeton = String(jeton || '').trim();
  if (!jeton) throw new UserError('Téléphone pas encore enregistré.', 'NON_ENREGISTRE');
  var a = lire_(APPAREILS).filter(function (x) { return String(x.jeton) === jeton; })[0];
  if (!a) throw new UserError('Ce téléphone n\'est plus enregistré. Demandez un nouveau code à la direction.', 'NON_ENREGISTRE');
  if (!estVrai_(a.actif)) throw new UserError('Ce téléphone a été désactivé par la direction.', 'DESACTIVE');
  var role = roleAppareil_(a.station);
  var cam = role === 'camion' ? camDe_(a.station) : null;
  if (role === 'camion' && !cam) throw new UserError('Ce camion n\'existe pas ou est inactif.', 'DESACTIVE');
  if (role) {
    try { majCellule_(APPAREILS, a._ligne, 'dernier_contact', maintenant_()); } catch (e) { }
    return { ligne: a, station: null, role: role, camion: cam, id: 'APP-' + String(a.code_activation), nom: String(a.nom || a.note || a.code_activation), username: 'app', enGroupe: false, chatInstance: '' };
  }
  var st = stations_().filter(function (s) { return String(s.code) === String(a.station).trim() && estVrai_(s.actif); })[0];
  if (!st) throw new UserError('Station inactive.', 'DESACTIVE');
  try { majCellule_(APPAREILS, a._ligne, 'dernier_contact', maintenant_()); } catch (e) { }
  return { ligne: a, station: st, id: 'APP-' + String(a.code_activation), nom: 'Téléphone ' + st.code };
}

/* ============================ Écran de départ ============================ */

function initMobile_(app) {
  if (app.role === 'transferts') return { version: MOBILE_VERSION, mode: 'transferts', utilisateur: { id: app.id, nom: app.nom, role: app.role } };
  if (app.role === 'camion') return camInit_(app);
  if (app.role) return initDirection_(app);
  var st = app.station, date = dateRapport();
  var tous = rapportsEffectifs_().filter(function (r) { return String(r.station) === st.code && !estAnnule(r); });
  var duJour = function (d) {
    return tous.filter(function (r) { return iso(r.date_rapport) === d; })
      .sort(function (a, b) { return ORDRE_PERIODE[String(a.periode || '')] - ORDRE_PERIODE[String(b.periode || '')]; });
  };
  var prerempliPour = function (d, jour, periode) {
    var prec = periode === 'PM' ? (jour.filter(function (r) { return String(r.periode) === 'AM'; })[0] || null) : null;
    if (!prec) {
      var avant = tous.filter(function (r) { return iso(r.date_rapport) < d; }).sort(function (a, b) {
        var da = iso(a.date_rapport), db = iso(b.date_rapport);
        if (da !== db) return da < db ? -1 : 1;
        return ORDRE_PERIODE[String(a.periode || '')] - ORDRE_PERIODE[String(b.periode || '')];
      });
      prec = avant.length ? avant[avant.length - 1] : null;
    }
    return prec ? { opening_meter: num(prec.closing_meter), starting_cash: num(prec.counted_cash) } : null;
  };
  var jour = duJour(date);
  var base = {
    version: MOBILE_VERSION, devise: CURRENCY, station: { code: st.code, nom: st.nom }, date: date, jours_max: JOURS_DIFFERE_MAX, argent: ARG_STATIONS ? 1 : 0,
    champs: FIELDS, envoyes: jour.map(recuEmploye_), periodes: periodesPossibles(jour), noms_connus: nomsConnus_(st.code)
  };
  base.prerempli = prerempliPour(date, jour, base.periodes[0]);
  base.plafonds = plafonds_(tous);
  // Rapports de cette station annulés par le bureau (8 derniers jours) : le téléphone l'affiche et propose de le refaire.
  var limite = Utilities.formatDate(new Date(new Date(date + 'T12:00:00Z').getTime() - 8 * 24 * 3600 * 1000), 'UTC', 'yyyy-MM-dd');
  base.annules = rapportsEffectifs_().filter(function (r) { return String(r.station) === st.code && estAnnule(r) && iso(r.date_rapport) >= limite; })
    .map(function (r) { return { ref: String(r.ref), date: iso(r.date_rapport), periode: String(r.periode || ''), motif: String(r.motif_annulation || '') }; });
  // Les jours passés : ce qui est déjà envoyé (non modifiable par la station) et ce qui manque encore.
  base.jours = [];
  for (var i = 1; i <= JOURS_DIFFERE_MAX; i++) {
    var d = Utilities.formatDate(new Date(new Date(date + 'T12:00:00Z').getTime() - i * 24 * 3600 * 1000), 'UTC', 'yyyy-MM-dd');
    var j = duJour(d), per = periodesPossibles(j);
    base.jours.push({ date: d, envoyes: j.map(function (r) { return { ref: r.ref, periode: String(r.periode || '') }; }), periodes: per,
      prerempli: per.length ? prerempliPour(d, j, per[0]) : null });
  }
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
  var arg = argStation_(p);
  if (arg && arg.bureau > 0) {   // l'argent remis au bureau sort de la caisse : il est compté avec les prélèvements dans le rapport et les anciennes feuilles
    var prel = String(p.withdrawals === undefined || p.withdrawals === null ? '' : p.withdrawals).trim(), pn = prel === '' ? '0' : phNombre_(prel);
    if (pn !== '') { p = Object.assign({}, p); p.withdrawals = String(Number(pn) + arg.bureau); }
  }
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
        '. Un rapport déjà envoyé ne peut pas être refait ni modifié par la station : seule la direction peut le corriger ou l\'annuler. Ce rapport n\'a pas été enregistré.', 'DEJA');
    }
    var v = val.valeurs;
    var notes = v.notes;
    if (arg) {
      if (v.bank_deposit > 0 && !arg.depot_par) throw new UserError('Dépôt en banque : écrivez le nom de la personne qui a fait le dépôt.', 'INVALIDE');
      if (v.counted_cash > 0 && !arg.caisse_par) throw new UserError('Caisse comptée : écrivez le nom de la personne qui garde la caisse.', 'INVALIDE');
      var ou = [];
      if (arg.bureau > 0) ou.push('remis au bureau ' + arg.bureau + ', reçu par ' + arg.bureau_qui + ' (compté dans les prélèvements)');
      if (v.bank_deposit > 0) ou.push('dépôt fait par ' + arg.depot_par);
      if (v.counted_cash > 0) ou.push('caisse gardée par ' + arg.caisse_par);
      if (ou.length) notes = (notes ? notes + ' ' : '') + '[Argent : ' + ou.join(' ; ') + ']';
    }
    // Chiffres visiblement faux : refusés tant que l'employé n'a pas revérifié et expliqué.
    var ctrls = controles_(st, date, attendue, v);
    var confirmes = [].concat(p.confirme_codes || []).map(String), expl = String(p.explication_controle || '').trim().slice(0, 400);
    var nonConfirmes = ctrls.filter(function (c) { return confirmes.indexOf(c.code) < 0; });
    if (nonConfirmes.length || (ctrls.length && !expl)) {
      var err = new UserError('Chiffres à revérifier :\n• ' + ctrls.map(function (c) { return c.msg; }).join('\n• '), 'VERIF');
      err.controles = ctrls; throw err;
    }
    if (ctrls.length) notes = (notes ? notes + ' ' : '') + '[⚠️ Contrôle : ' + ctrls.map(function (c) { return c.court; }).join(', ') + ' — explication de l\'employé : ' + expl + ']';
    if (p.en_retard) {
      var raison = String(p.raison_retard || '').trim().slice(0, 300);
      if (!raison) throw new UserError('Rapport en retard : écrivez la raison du retard.', 'INVALIDE');
      notes = (notes ? notes + ' ' : '') + '[Rapport en retard : ' + raison + ']';
    }
    if (p.photo && String(p.photo_source || '') === 'galerie') {   // photo du registre prise dans la galerie (propriétaire, 27/11/2026)
      var pg = String(p.photo_galerie || '').replace(/\s+/g, ' ').trim().slice(0, 200);
      if (pg.length < 3) throw new UserError('Photo du registre choisie dans la galerie : écrivez pourquoi.', 'INVALIDE');
      notes = (notes ? notes + ' ' : '') + '[Photo du registre choisie dans la galerie : ' + pg + ']';
    }
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
    if (arg) try {
      argEntrees_(st.code, st.nom, date, ref, [{ lieu: 'depot', montant: v.bank_deposit, tenu_par: arg.depot_par }, { lieu: 'bureau', montant: arg.bureau, tenu_par: arg.bureau_qui }, { lieu: 'caisse', montant: v.counted_cash, tenu_par: arg.caisse_par }], app.id, '');
    } catch (e) { console.error(e && e.stack || e); }   // le rapport est déjà enregistré : le registre d'argent ne doit pas le faire échouer
    var r = rapportsEffectifs_().filter(function (x) { return String(x.ref) === ref; })[0];
    var reste = periodesPossibles(rapportsDuJour_(st.code, date));
    return { recu: recuEmploye_(r), reste: reste.length ? reste[0] : null };
  } finally { lock.releaseLock(); }
}


/* ============================ Rapports de gestion ============================ */

/** Tous les rapports (corrections appliquées, avec ventes attendues et écart) d'une période jusqu'à 1 an. Les calculs se font sur le téléphone. */
function gestion_(p) {
  var du = iso(p.du || dateRapport()), au = iso(p.au || du);
  if (au < du) throw new UserError('Dates inversées.');
  if ((new Date(au) - new Date(du)) / 86400000 > 366) throw new UserError('Période trop longue (1 an maximum).');
  var prix = lire_(SHEETS.PRIX);
  var rapports = rapportsEffectifs_().filter(function (r) { var d = iso(r.date_rapport); return d >= du && d <= au; }).map(function (r) {
    var o = vueDirection_(r, prix); delete o.photo; return o;
  });
  return { du: du, au: au, aujourdhui: dateRapport(), rapports: rapports,
    stations: stations_().map(function (s) { return { code: String(s.code), nom: String(s.nom), actif: estVrai_(s.actif) }; }) };
}

/* ============================ Contrôles ============================ */

function mediane_(a) { a = a.filter(function (x) { return x > 0; }).sort(function (x, y) { return x - y; }); return a.length ? a[Math.floor(a.length / 2)] : 0; }

/** Plafonds de la station, calculés sur ses 30 derniers rapports (sans prix : rien de secret pour l'employé). */
function plafonds_(rapportsStation) {
  var r = rapportsStation.slice().sort(function (a, b) { return iso(a.date_rapport) < iso(b.date_rapport) ? 1 : -1; }).slice(0, 30);
  var caisse = mediane_(r.map(function (x) { return Math.max(num(x.starting_cash), num(x.counted_cash)); }));
  var gal = mediane_(r.map(function (x) { return num(x.closing_meter) - num(x.opening_meter); }));
  return { montant: Math.round(Math.max(CTRL.PLAFOND_MIN, CTRL.PLAFOND_FOIS * caisse)), gallons: Math.round(Math.max(CTRL.GALLONS_MIN, CTRL.GALLONS_FOIS * gal)) };
}

/** Montants trop grands. La caisse s'accumule tant qu'il n'y a pas de dépôt : le solde initial n'est jamais plafonné
 *  (il est contrôlé par la caisse du rapport précédent) ; caisse comptée et dépôt sont plafonnés à « solde initial + entrées + plafond ». */
function trop_(v, plafond, val) {
  var base = (Number(val('starting_cash')) || 0) + (Number(val('other_cash_in')) || 0);
  return FIELDS.filter(function (x) {
    if (!x.money || x.k === 'starting_cash') return false;
    var lim = (x.k === 'counted_cash' || x.k === 'bank_deposit') ? base + plafond : plafond;
    return Number(val(x.k)) > lim;
  });
}

/** Liste des anomalies d'un rapport. Le message pour l'employé ne montre jamais le prix, la caisse attendue ni l'écart. */
function controles_(st, date, periode, v) {
  var tous = rapportsEffectifs_().filter(function (r) { return String(r.station) === st.code && !estAnnule(r); });
  var out = [], pl = plafonds_(tous), f = function (n) { return Number(n).toLocaleString('fr-FR'); };
  var cle = function (d, per) { return d + '|' + ORDRE_PERIODE[String(per || '')]; };
  var moi = cle(date, periode);
  var avant = tous.filter(function (r) { return cle(iso(r.date_rapport), r.periode) < moi; })
    .sort(function (a, b) { return cle(iso(a.date_rapport), a.periode) < cle(iso(b.date_rapport), b.periode) ? 1 : -1; })[0];
  if (avant) {
    var dPrec = iso(avant.date_rapport).split('-').reverse().join('/');
    if (num(avant.closing_meter) !== v.opening_meter) out.push({ code: 'COMPTEUR', court: 'compteur ouverture ≠ fermeture précédente',
      msg: 'Compteur d\'ouverture ' + f(v.opening_meter) + ' : le dernier rapport (' + dPrec + ') a fermé à ' + f(num(avant.closing_meter)) + '. S\'il manque un jour, faites d\'abord le « Rapport manquant ».' });
    if (num(avant.counted_cash) !== v.starting_cash) out.push({ code: 'CAISSE_DEPART', court: 'solde initial ≠ caisse précédente',
      msg: 'Solde initial ' + f(v.starting_cash) + ' : la caisse comptée au dernier rapport (' + dPrec + ') était ' + f(num(avant.counted_cash)) + '.' });
  }
  var gros = trop_(v, pl.montant, function (k) { return v[k]; });
  if (gros.length) out.push({ code: 'MONTANT', court: 'montant trop grand (' + gros.map(function (x) { return x.label; }).join(', ') + ')',
    msg: gros.map(function (x) { return x.label + ' ' + f(v[x.k]); }).join(', ') + ' : montant beaucoup trop grand pour cette station. Un seul montant par case : additionnez les transferts et écrivez le total.' });
  if (v.gallons > pl.gallons) out.push({ code: 'GALLONS', court: v.gallons + ' gal vendus', msg: f(v.gallons) + ' gallons vendus : beaucoup plus que d\'habitude. Vérifiez les deux compteurs.' });
  var prix = prixPour(lire_(SHEETS.PRIX), st.code, date);
  if (prix && !gros.length) {
    var c = calculsDirection(v, prix);
    if (c.ecart !== null && Math.abs(c.ecart) > Math.max(CTRL.ECART_MIN, CTRL.ECART_PCT * c.ventes_attendues)) out.push({ code: 'ECART', court: 'grand écart de caisse',
      msg: 'La caisse comptée ne correspond pas du tout aux gallons vendus et aux montants entrés. Recomptez la caisse et vérifiez chaque case (dépôt, dépenses, transferts, ventes à crédit). Si c\'est juste (ex. vente à crédit), expliquez-le.' });
  }
  return out;
}

/* ============================ Argent : qui le tient, jusqu'à la preuve (m1.28, 27/11/2026) ============================
 * Demande du propriétaire (27/11/2026) : pour les camions et les stations, l'argent de chaque rapport va à l'un de trois endroits —
 * déposé à la banque, remis au bureau, gardé en main — chaque fois avec le nom de la personne qui l'a, et la trace doit être fermée.
 * Onglet « Argent (app) », où rien n'est jamais effacé :
 *   genre « entree »  : une ligne par endroit et par rapport (lieu = depot, bureau, main, ou caisse pour la caisse comptée d'une station) ;
 *   genre « cloture » : quelqu'un du bureau confirme, avec une pièce justificative (photo), que l'argent est bien arrivé
 *                       (mode = recu_bureau ou depot_verifie). Une confirmation peut être partielle : le reste demeure ouvert.
 * Une entrée ne compte que si son rapport existe toujours et n'est pas annulé. La caisse d'une station n'est pas une somme à fermer :
 * elle passe au rapport suivant (solde initial) ; on montre seulement qui la garde d'après le dernier rapport.
 */
var ARG_ONGLET = 'Argent (app)';
var ARG_ENTETES = ['id', 'genre', 'origine', 'origine_nom', 'date', 'rapport', 'lieu', 'montant', 'devise', 'tenu_par', 'pour', 'fichier_id', 'parent', 'mode', 'par', 'le', 'note', 'cid', 'annule'];
var ARG_STATIONS = true;    // stations activées le 27/11/2026 à la demande du propriétaire (« turn on cash box today »)
var ARG_MODES = { depot: ['depot_verifie'], bureau: ['recu_bureau'], main: ['recu_bureau', 'depot_verifie'] };
function argFeuille_() { var sh = camFeuille_(ARG_ONGLET, ARG_ENTETES); camEntetes_(sh, ARG_ENTETES); return sh; }
function argLire_(max) {
  var sh = argFeuille_(), n = sh.getLastRow();
  if (n < 2) return [];
  var from = Math.max(2, n - (max || 10000) + 1);
  return sh.getRange(from, 1, n - from + 1, ARG_ENTETES.length).getValues().map(function (r, i) {
    return { ligne: from + i, id: String(r[0]), genre: String(r[1]), origine: String(r[2]), origine_nom: String(r[3]), date: phJour_(r[4]), rapport: String(r[5]), lieu: String(r[6]), montant: Number(r[7]) || 0, devise: String(r[8]),
      tenu_par: String(r[9] || ''), pour: String(r[10] || ''), fichier_id: String(r[11] || ''), parent: String(r[12] || ''), mode: String(r[13] || ''), par: String(r[14] || ''), le: phTexte_(r[15]), note: String(r[16] || ''), cid: String(r[17] || ''), annule: estVrai_(r[18]) };
  });
}
/** Écrit les entrées d'un rapport (à appeler sous le verrou du script). parts : [{ lieu, montant, tenu_par, pour }] ; les montants nuls sont laissés de côté. */
function argEntrees_(origine, nom, date, rapport, parts, appId, fid) {
  var sh = argFeuille_(), deja = {}, le = maintenant_();
  argLire_(10000).forEach(function (x) { if (x.genre === 'entree') deja[x.cid] = 1; });
  parts.forEach(function (x) {
    var m = Number(x.montant) || 0, cid = rapport + ':' + x.lieu;
    if (!(m > 0) || deja[cid]) return;
    sh.appendRow(['A' + date.replace(/-/g, '').slice(2) + '-' + sh.getLastRow(), 'entree', origine, nom, date, rapport, x.lieu, m, CURRENCY, String(x.tenu_par || '').slice(0, 60), String(x.pour || ''), x.lieu === 'depot' ? (fid || '') : '', '', '', appId, le, '', cid, '']);
  });
}
/** Les entrées valables avec ce qui reste ouvert. */
function argEtat_() {
  var l = argLire_(10000), auj = msgJour_(0), okCam = {}, okSt = {}, ordre = {};
  camRapports_(3000).forEach(function (r) { if (!r.annule) okCam[r.id] = 1; });
  rapportsEffectifs_().forEach(function (r) { if (!estAnnule(r)) { okSt[String(r.ref)] = 1; ordre[String(r.ref)] = iso(r.date_rapport) + '|' + (ORDRE_PERIODE[String(r.periode || '')] || 0) + '|' + String(r.soumis_le || ''); } });
  var regle = {}, clot = {};
  l.forEach(function (x) { if (x.genre === 'cloture' && !x.annule) { regle[x.parent] = Math.round(((regle[x.parent] || 0) + x.montant) * 100) / 100; (clot[x.parent] = clot[x.parent] || []).push(x); } });
  var entrees = l.filter(function (x) { return x.genre === 'entree' && !x.annule && (x.origine.indexOf('CAM:') === 0 ? okCam[x.rapport] : okSt[x.rapport]) && x.origine !== 'CAM:' + CAM_TEST; });
  var derniere = {};   // caisse des stations : seule celle du dernier rapport de chaque station compte
  entrees.forEach(function (x) { if (x.lieu === 'caisse' && (!derniere[x.origine] || (ordre[x.rapport] || '') > (ordre[derniere[x.origine].rapport] || ''))) derniere[x.origine] = x; });
  var jours = function (d) { return Math.max(0, Math.round((new Date(auj + 'T12:00:00Z') - new Date(d + 'T12:00:00Z')) / 864e5)); };
  return { toutes: l, clot: clot,
    caisses: Object.keys(derniere).map(function (k) { var x = derniere[k]; return { origine: x.origine, nom: x.origine_nom, date: x.date, montant: x.montant, tenu_par: x.tenu_par, rapport: x.rapport }; }),
    lignes: entrees.filter(function (x) { return x.lieu !== 'caisse'; }).map(function (x) {
      return { id: x.id, origine: x.origine, nom: x.origine_nom, date: x.date, rapport: x.rapport, lieu: x.lieu, montant: x.montant, devise: x.devise, tenu_par: x.tenu_par, pour: x.pour, fichier: x.fichier_id,
        regle: regle[x.id] || 0, ouvert: Math.round((x.montant - (regle[x.id] || 0)) * 100) / 100, jours: jours(x.date) };
    }) };
}
/** Direction / comptabilité : ce qui est ouvert (par personne), la caisse des stations, et les confirmations des 30 derniers jours. */
function argListe_() {
  var e = argEtat_(), ouverts = e.lignes.filter(function (x) { return x.ouvert > 0; }).sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.id < b.id ? -1 : 1); });
  var tot = { depot: 0, bureau: 0, main: 0 }, pers = {};
  ouverts.forEach(function (x) {
    tot[x.lieu] = Math.round(((tot[x.lieu] || 0) + x.ouvert) * 100) / 100;
    var k = x.tenu_par.toLowerCase(), p = pers[k] = pers[k] || { nom: x.tenu_par, total: 0, lignes: [] };
    p.total = Math.round((p.total + x.ouvert) * 100) / 100; p.lignes.push(x);
  });
  var parId = {}; e.lignes.forEach(function (x) { parId[x.id] = x; });
  var lim = msgJour_(-30), faites = e.toutes.filter(function (x) { return x.genre === 'cloture' && phJour_(x.le) >= lim; }).sort(function (a, b) { return a.le < b.le ? 1 : -1; }).slice(0, 200).map(function (c) {
    var x = parId[c.parent] || {};
    return { id: c.id, parent: c.parent, le: c.le, par: c.par, mode: c.mode, montant: c.montant, note: c.note, fichier: c.fichier_id, annule: c.annule, nom: x.nom || c.origine_nom, date: x.date || c.date, lieu: x.lieu || c.lieu, tenu_par: x.tenu_par || c.tenu_par };
  });
  return { devise: CURRENCY, aujourdhui: msgJour_(0), stations_actives: ARG_STATIONS, totaux: tot, ouverts: ouverts,
    personnes: Object.keys(pers).map(function (k) { return pers[k]; }).sort(function (a, b) { return b.total - a.total; }),
    caisses: e.caisses.sort(function (a, b) { return a.nom < b.nom ? -1 : 1; }), faites: faites };
}
/** Fermer la trace : quelqu'un du bureau confirme, pièce à l'appui, que l'argent est arrivé (au bureau ou à la banque). Toujours avec une pièce. */
function argConfirmer_(ctx, p) {
  var montant = Number(phNombre_(String(p.montant === undefined || p.montant === null ? '' : p.montant).trim()));
  if (!(montant > 0)) throw new UserError('Montant illisible. Écrivez seulement les chiffres.');
  var note = String(p.note || '').trim().slice(0, 200), doc = String(p.doc || ''), m = doc.match(/^data:image\/(jpeg|png);base64,(.+)$/), octets = null;
  if (doc && !m) throw new UserError('Pièce invalide : il faut une photo (image).');
  if (m) { octets = Utilities.base64Decode(m[2]); if (octets.length > 3000000) throw new UserError('Photo de la pièce trop lourde.'); }
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var x = argEtat_().lignes.filter(function (y) { return y.id === String(p.id || ''); })[0];
    if (!x) throw new UserError('Ligne introuvable (rapport annulé ?). Actualisez la page.');
    var mode = String(p.mode || '');
    if ((ARG_MODES[x.lieu] || []).indexOf(mode) < 0) throw new UserError('Choisissez où l\'argent est arrivé.');
    if (montant > x.ouvert) throw new UserError('Il ne reste que ' + x.ouvert + ' ' + x.devise + ' à confirmer sur cette ligne.');
    var memePiece = !octets && estVrai_(p.piece_du_rapport) && x.fichier;
    if (!octets && !memePiece) throw new UserError('Ajoutez la pièce justificative : une photo du reçu signé, du bordereau ou du relevé.');
    var fid = memePiece ? x.fichier : '', le = maintenant_(), auj = msgJour_(0);
    if (octets) {
      var f = phDossier_({ code: 'ARGENT', nom: 'Pièces justificatives' }, auj, 'ARG', 'RRD – Argent, pièces justificatives (app, privé)').createFile(Utilities.newBlob(octets, 'image/' + m[1], 'ARGENT_' + x.id + '_' + le.replace(/[^0-9]/g, '') + '.' + (m[1] === 'png' ? 'png' : 'jpg')));
      f.setDescription('Pièce justificative (application RRD) — ' + x.nom + ' — ' + x.date + ' — ' + montant + ' ' + x.devise + ' — confirmé par ' + ctx.nom); fid = f.getId();
    }
    var sh = argFeuille_(), id = 'F' + auj.replace(/-/g, '').slice(2) + '-' + sh.getLastRow();
    sh.appendRow([id, 'cloture', x.origine, x.nom, auj, x.rapport, x.lieu, montant, x.devise, x.tenu_par, x.pour, fid, x.id, mode, ctx.nom + ' (' + (ctx.role === 'accounting' ? 'comptabilité' : 'direction') + ')', le, note, id, '']);
    journal_(ctx, 'argent-confirme', x.id + ' ' + x.nom + ' ' + x.date + ' ' + x.lieu + ' ' + montant + ' → ' + mode + (note ? ' : ' + note : ''));
  } finally { lock.releaseLock(); }
  return argListe_();
}
/** La direction annule une confirmation faite par erreur : la ligne reste, marquée, et l'argent redevient ouvert. */
function argAnnuler_(ctx, p) {
  var motif = String(p.motif || '').trim().slice(0, 150);
  if (!motif) throw new UserError('Écrivez la raison.');
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var c = argLire_(10000).filter(function (x) { return x.genre === 'cloture' && x.id === String(p.id || ''); })[0];
    if (!c) throw new UserError('Confirmation introuvable.');
    if (c.annule) throw new UserError('Déjà annulée.');
    var sh = argFeuille_();
    sh.getRange(c.ligne, ARG_ENTETES.indexOf('annule') + 1).setValue('oui');
    sh.getRange(c.ligne, ARG_ENTETES.indexOf('note') + 1).setValue((c.note ? c.note + ' ; ' : '') + 'Annulée le ' + maintenant_() + ' par ' + ctx.nom + ' : ' + motif);
    journal_(ctx, 'argent-annule', c.id + ' (' + c.parent + ') : ' + motif);
  } finally { lock.releaseLock(); }
  return argListe_();
}
/** Station (quand ARG_STATIONS est activé) : lit et vérifie les cases « où est l'argent » d'un rapport de fermeture. null = ancien téléphone, rien à tracer. */
function argStation_(p) {
  if (!ARG_STATIONS || !estVrai_(p.argent_v)) return null;
  var mt = function (v, quoi) { var t = String(v === undefined || v === null ? '' : v).trim(); if (t === '' || t === '0') return 0; var x = phNombre_(t); if (x === '') throw new UserError(quoi + ' : montant illisible. Écrivez seulement les chiffres.', 'INVALIDE'); return Number(x); };
  var nom = function (k) { return String(p[k] || '').trim().replace(/\s+/g, ' ').slice(0, 60); };
  var a = { bureau: mt(p.remis_bureau, 'Remis au bureau'), bureau_qui: nom('remis_bureau_a'), depot_par: nom('depot_par'), caisse_par: nom('caisse_par') };
  if (a.bureau > 0 && !a.bureau_qui) throw new UserError('Remis au bureau : écrivez le nom de la personne du bureau qui a reçu l\'argent.', 'INVALIDE');
  return a;
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
  if (!app.role || app.role === 'transferts' || app.role === 'camion') throw new UserError('Accès réservé à la direction.', 'FORBIDDEN');
  var ctx = { id: app.id, nom: app.nom, username: 'app', role: app.role, chatType: 'private', chatInstance: '', enGroupe: false };
  switch (sous) {
    case 'rapports': return listeRapports_(ctx, p);
    case 'gestion': exiger_(ctx, ['management', 'accounting']); return gestion_(p);
    case 'detail': return detailRapport_(ctx, p);
    case 'corriger': return corriger_(ctx, p);
    case 'annuler': return annuler_(ctx, p);
    case 'resume': return resumeJour_(ctx, p);
    case 'prix': return listePrix_(ctx);
    case 'ajouterPrix': return ajouterPrix_(ctx, p);
    case 'appareils': exiger_(ctx, ['management']); return listeAppareils_();
    case 'ajouterAppareil': exiger_(ctx, ['management']); return ajouterAppareil_(ctx, p);
    case 'activerAppareil': exiger_(ctx, ['management']); return activerAppareil_(ctx, p);
    case 'libererAppareil': exiger_(ctx, ['management']); return libererAppareil_(ctx, p);
    case 'photosSuivi': exiger_(ctx, ['management', 'accounting']); return photosSuivi_(p);
    case 'photosListe': exiger_(ctx, ['management', 'accounting']); return photosListe_(p);
    case 'photoVoir': exiger_(ctx, ['management', 'accounting']); return photoVoir_(p);
    case 'bordereaux': exiger_(ctx, ['management', 'accounting']); return bordereauxListe_(p);
    case 'messages': exiger_(ctx, ['management', 'accounting']); return messagesListe_();
    case 'messageAjouter': exiger_(ctx, ['management']); return messageAjouter_(ctx, p);
    case 'messageSupprimer': exiger_(ctx, ['management']); return messageSupprimer_(ctx, p);
    case 'stock': exiger_(ctx, ['management', 'accounting']); return stockEstime_();
    case 'stockRegler': exiger_(ctx, ['management']); return stockRegler_(ctx, p);
    case 'camion': exiger_(ctx, ['management', 'accounting']); return camListe_(p);
    case 'camClient': exiger_(ctx, ['management', 'accounting']); return camClient_(ctx, p);
    case 'camReleve': exiger_(ctx, ['management', 'accounting']); return camReleve_(p);
    case 'argent': exiger_(ctx, ['management', 'accounting']); return argListe_();
    case 'argentConfirmer': exiger_(ctx, ['management', 'accounting']); return argConfirmer_(ctx, p);
    case 'argentAnnuler': exiger_(ctx, ['management']); return argAnnuler_(ctx, p);
    case 'livrer': exiger_(ctx, ['management', 'accounting']); return sugListe_(p);
    case 'vracRelier': exiger_(ctx, ['management', 'accounting']); return vracRelier_(ctx, p);
    case 'camAnnuler': exiger_(ctx, ['management']); return camAnnuler_(ctx, p);
    case 'camRapportAnnuler': exiger_(ctx, ['management']); return camRapportAnnuler_(ctx, p);
    case 'demandes': exiger_(ctx, ['management', 'accounting']); return demandesListe_();
    case 'demandeStatut': exiger_(ctx, ['management']); return demandeStatut_(ctx, p);
    case 'honneur': exiger_(ctx, ['management', 'accounting']); return honneurBureau_(p);
    case 'honneurConfirmer': exiger_(ctx, ['management']); return honneurConfirmer_(ctx, p);
    case 'honneurRetirer': exiger_(ctx, ['management']); return honneurRetirer_(ctx, p);
    default: throw new UserError('Action inconnue.');
  }
}

function listeAppareils_() {
  var noms = {}; stations_().forEach(function (s) { noms[String(s.code)] = s.nom; });
  noms.DIRECTION = 'Direction'; noms.COMPTA = 'Comptabilité';
  var tr = agentsTransfertsSeuls_(); tr.forEach(function (t) { noms[t.code] = t.nom; });
  var cams = camions_().filter(function (c) { return c.actif; }); cams.forEach(function (c) { noms['CAM:' + c.code] = '🚚 ' + c.nom; });
  var partage = function (st) { st = String(st).trim().toUpperCase(); return !!ROLES_APPAREIL[st] || st === STATION_FORMATION || st === 'CAM:' + CAM_TEST; };
  var grand = function (st) { st = String(st).trim().toUpperCase(); return st === STATION_FORMATION || st === 'CAM:' + CAM_TEST; };
  return {
    transferts: tr, camions: cams.map(function (c) { return { code: 'CAM:' + c.code, nom: c.nom }; }),
    stations: stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return { code: s.code, nom: s.nom }; }),
    appareils: lire_(APPAREILS).map(function (a) {
      return { code: String(a.code_activation), station: String(a.station), station_nom: noms[String(a.station).trim()] || String(a.station),
        actif: estVrai_(a.actif), utilise: !!String(a.jeton || '').trim(), enregistre_le: String(a.enregistre_le || ''),
        partage: partage(a.station), nb: appListe_(a.appareil_id).length,
        max: grand(a.station) ? MAX_APPAREILS_FORMATION : MAX_APPAREILS_ROLE,
        dernier_contact: String(a.dernier_contact || ''), nom: String(a.nom || ''), push: !!String(a.push || '').trim() };
    })
  };
}

function ajouterAppareil_(ctx, p) {
  var station = String(p.station || '').trim().toUpperCase();
  var valides = stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return String(s.code); }).concat(['DIRECTION', 'COMPTA'])
    .concat(agentsTransfertsSeuls_().map(function (t) { return t.code; })).concat(camions_().filter(function (c) { return c.actif; }).map(function (c) { return 'CAM:' + c.code; }));
  if (valides.indexOf(station) < 0) throw new UserError('Station inconnue.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var existants = lire_(APPAREILS).map(function (a) { return String(a.code_activation).trim().toUpperCase(); });
    var prefixe = station === 'DIRECTION' ? 'DIR' : station === 'COMPTA' ? 'CPT' : station.indexOf('CAM:') === 0 ? 'CM' + station.slice(4, 5) : station.replace(/^TR:/, ''), code;
    if (station === 'CAM:' + CAM_TEST && existants.indexOf(CAM_CODE_TEST) < 0) code = CAM_CODE_TEST;   // le code du camion de test, toujours le même
    else do { code = prefixe + '-' + (1000 + Math.floor(Math.random() * 9000)); } while (existants.indexOf(code) >= 0);
    ajouter_(APPAREILS, { code_activation: code, station: station, actif: 'oui', note: 'créé par ' + ctx.nom + ' le ' + maintenant_(), nom: String(p.nom || '').trim().slice(0, 60) });
    journal_(ctx, 'appareil-code', code + ' → ' + station);
  } finally { lock.releaseLock(); }
  var l = listeAppareils_(); l.nouveau = code; return l;
}

/** Libère un code : le téléphone qui l'utilisait est déconnecté, et le code peut être enregistré une fois sur n'importe quel téléphone. */
function libererAppareil_(ctx, p) {
  var a = lire_(APPAREILS).filter(function (x) { return String(x.code_activation) === String(p.code); })[0];
  if (!a) throw new UserError('Code introuvable.');
  if (String(a.code_activation) === ctx.id.replace(/^APP-/, '')) throw new UserError('Vous ne pouvez pas libérer le code de votre propre téléphone.');
  majCellule_(APPAREILS, a._ligne, 'jeton', '');
  majCellule_(APPAREILS, a._ligne, 'appareil_id', '');
  try { if (String(a.push || '')) majCellule_(APPAREILS, a._ligne, 'push', ''); } catch (e) { }
  majCellule_(APPAREILS, a._ligne, 'note', String(a.note || '') + ' | libéré le ' + maintenant_() + ' par ' + ctx.nom);
  journal_(ctx, 'appareil-libere', String(a.code_activation));
  return listeAppareils_();
}

function activerAppareil_(ctx, p) {
  var a = lire_(APPAREILS).filter(function (x) { return String(x.code_activation) === String(p.code); })[0];
  if (!a) throw new UserError('Code introuvable.');
  if (String(a.code_activation) === ctx.id.replace(/^APP-/, '') && !p.actif) throw new UserError('Vous ne pouvez pas désactiver votre propre téléphone.');
  majCellule_(APPAREILS, a._ligne, 'actif', p.actif ? 'oui' : 'non');
  journal_(ctx, p.actif ? 'appareil-on' : 'appareil-off', String(a.code_activation));
  return listeAppareils_();
}

/* ============================ Photos horaires (m1.11, 26/11/2026) ============================
 * Remplace les groupes photos de Telegram : chaque heure (7 h à 16 h) la station prend 3 photos dans l'application
 * (compteur, station, citerne). Les photos vont dans un dossier Drive privé, une ligne par photo dans l'onglet
 * « Photos horaires » de la feuille privée. L'heure retenue est celle du téléphone au moment de la photo
 * (jamais dans le futur) ; l'heure de réception est notée à côté. Rien n'est jamais supprimé ici.
 */
var PH_ONGLET = 'Photos horaires';
var PH_ENTETES = ['date', 'station', 'heure', 'type', 'pris_le', 'recu_le', 'compteur', 'fichier_id', 'appareil', 'cid', 'note'];
var PH_TYPES = ['COMPTEUR', 'STATION', 'CITERNE'];
var PH_LETTRE = { COMPTEUR: 'C', STATION: 'S', CITERNE: 'T' };   // lettres de la grille du tableau de bord
var PH_HEURES = [7, 8, 9, 10, 11, 12, 13, 14, 15, 16];   // heures comptées
var PH_HEURES_VUES = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19];   // heures montrées (5 h à 20 h) : une photo prise y est marquée même si l'heure n'est pas comptée (propriétaire, 26/11/2026)
var PH_JAUGES = ['JAUGE_OUV', 'JAUGE_FER'];   // manomètre : ouverture = avant midi, fermeture = à partir de 15 h (règles du 09/11/2026)
var PH_NON_SUIVIES = ['M83', 'LEO'];   // décision du propriétaire (21/11/2026) : Mahotière 83 et Léogâne ne sont pas suivies
var PH_SANS_CITERNE = ['DIQ'];         // décision du propriétaire (23/11/2026) : la citerne n'est pas comptée à Diquini
var PH_MAX_OCTETS = 1500000;

function phFeuille_() {
  var sh = ss_().getSheetByName(PH_ONGLET);
  if (!sh) {
    sh = ss_().insertSheet(PH_ONGLET);
    sh.getRange(1, 1, sh.getMaxRows(), PH_ENTETES.length).setNumberFormat('@');   // texte : les dates et heures restent telles qu'écrites
    sh.appendRow(PH_ENTETES);
    sh.setFrozenRows(1);
  }
  return sh;
}
function phJour_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '').slice(0, 10); }
function phTexte_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss') : String(v || '').replace(/^(\d{4}-\d{2}-\d{2}) (\d):/, '$1 0$2:'); }
function phLire_(max) {
  var sh = ss_().getSheetByName(PH_ONGLET);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = sh.getLastRow(), from = Math.max(2, n - (max || 3000) + 1);
  return sh.getRange(from, 1, n - from + 1, PH_ENTETES.length).getValues().map(function (r) {
    return { date: phJour_(r[0]), station: String(r[1]), heure: Number(r[2]), type: String(r[3]), pris_le: phTexte_(r[4]), recu_le: phTexte_(r[5]),
      compteur: r[6] === '' || r[6] === null ? '' : String(r[6]), fichier_id: String(r[7]), appareil: String(r[8]), cid: String(r[9]), note: String(r[10] || '') };
  });
}
/** Même règle que les rapports : « 1 788 » ou « 1,788 » = mille sept cent quatre-vingt-huit ; le point est la décimale. '' si illisible. */
function phNombre_(v) {
  var s = String(v === undefined || v === null ? '' : v).trim();
  if (!/^(\d{1,3}([ \u00a0\u202f,]\d{3})+|\d+)(\.\d+)?$/.test(s)) return '';
  return s.replace(/[ \u00a0\u202f,]/g, '');
}
function phTypesDe_(code) { return PH_SANS_CITERNE.indexOf(code) >= 0 ? ['COMPTEUR', 'STATION'] : PH_TYPES; }

/** Dossier Drive privé : « RRD – Photos horaires (app, privé) » / station / mois. */
function phDossier_(st, date, prefixe, nomRacine) {
  prefixe = prefixe || 'PH'; nomRacine = nomRacine || 'RRD – Photos horaires (app, privé)';
  var props = PropertiesService.getScriptProperties(), mois = date.slice(0, 7), cle = prefixe + '_DOSSIER_' + st.code + '_' + mois;
  var id = props.getProperty(cle);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  var racine = null, rid = props.getProperty(prefixe + '_RACINE');
  if (rid) { try { racine = DriveApp.getFolderById(rid); } catch (e) { } }
  if (!racine) { racine = DriveApp.createFolder(nomRacine); props.setProperty(prefixe + '_RACINE', racine.getId()); }
  var nom = st.code + ' – ' + st.nom, it = racine.getFoldersByName(nom), ds = it.hasNext() ? it.next() : racine.createFolder(nom);
  var it2 = ds.getFoldersByName(mois), dm = it2.hasNext() ? it2.next() : ds.createFolder(mois);
  props.setProperty(cle, dm.getId());
  return dm;
}

/** Une photo envoyée par le téléphone d'une station. Le même « cid » renvoyé deux fois n'est enregistré qu'une fois. */
function photoHoraire_(app, p) {
  if (app.role || !app.station) throw new UserError('Les photos se font depuis le téléphone de la station.', 'FORBIDDEN');
  var st = app.station, type = String(p.type || '').toUpperCase();
  if (PH_TYPES.indexOf(type) < 0 && PH_JAUGES.indexOf(type) < 0) throw new UserError('Type de photo inconnu.');
  var cid = String(p.cid || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 60);
  if (!cid) throw new UserError('Photo sans identifiant.');
  var m = String(p.photo || '').match(/^data:image\/jpeg;base64,(.+)$/);
  if (!m) throw new UserError('Photo invalide.');
  var octets = Utilities.base64Decode(m[1]);
  if (octets.length > PH_MAX_OCTETS) throw new UserError('Photo trop lourde.');
  var recu = new Date(), pris = new Date(Number(p.pris_ms) || 0), note = '';
  if (!(pris.getTime() > 0) || pris.getTime() > recu.getTime() + 10 * 60000) {
    note = 'Heure du téléphone refusée (' + (pris.getTime() > 0 ? Utilities.formatDate(pris, TZ, 'yyyy-MM-dd HH:mm') : 'absente') + ') : heure de réception utilisée';
    pris = recu;
  } else if (recu.getTime() - pris.getTime() > 31 * 864e5) throw new UserError('Photo trop ancienne (plus de 31 jours).');
  var date = Utilities.formatDate(pris, TZ, 'yyyy-MM-dd'), heure = Number(Utilities.formatDate(pris, TZ, 'H'));
  var compteur = '', jauge = PH_JAUGES.indexOf(type) >= 0;   // colonne « compteur » : index du compteur, ou niveau du manomètre en %
  if ((type === 'COMPTEUR' || jauge) && p.compteur !== undefined && p.compteur !== null) {
    var c = phNombre_(String(p.compteur).replace(/%/g, ''));
    if (c && (!jauge || Number(c) <= 100)) compteur = c;
    else if (String(p.compteur).trim()) note = (note ? note + ' ; ' : '') + (jauge ? 'Niveau saisi illisible : ' : 'Index saisi illisible : ') + String(p.compteur).slice(0, 30);
  }
  if (type === 'JAUGE_OUV' && heure >= 12) note = (note ? note + ' ; ' : '') + 'Ouverture prise après midi';
  if (type === 'JAUGE_FER' && heure < 15) note = (note ? note + ' ; ' : '') + 'Fermeture prise avant 15 h';
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var sh = phFeuille_(), n = sh.getLastRow();
    if (n > 1) {
      var from = Math.max(2, n - 1999), cids = sh.getRange(from, 10, n - from + 1, 1).getValues();
      for (var i = 0; i < cids.length; i++) if (String(cids[i][0]) === cid) return { recu: true, deja: true, cid: cid, date: date, heure: heure, type: type };
    }
    var nomF = st.code + '_' + date + '_' + Utilities.formatDate(pris, TZ, 'HHmm') + '_' + type + '.jpg';
    var f = phDossier_(st, date).createFile(Utilities.newBlob(octets, 'image/jpeg', nomF));
    f.setDescription('Photo horaire (application RRD) — ' + st.nom + ' — ' + type + ' — ' + cid);
    sh.appendRow([date, st.code, String(heure), type, Utilities.formatDate(pris, TZ, 'yyyy-MM-dd HH:mm:ss'), Utilities.formatDate(recu, TZ, 'yyyy-MM-dd HH:mm:ss'),
      compteur, f.getId(), app.id, cid, note]);
    return { recu: true, cid: cid, date: date, heure: heure, type: type };
  } finally { lock.releaseLock(); }
}

/** Pour le téléphone de la station : ce qui est déjà reçu aujourd'hui (aussi depuis un autre téléphone de la même station). */
function photosStation_(app) {
  if (app.role || !app.station) throw new UserError('Réservé au téléphone de la station.', 'FORBIDDEN');
  var code = String(app.station.code), now = new Date(), date = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), faits = {};
  phLire_(3000).forEach(function (r) {
    if (r.station !== code || r.date !== date) return;
    (faits[r.heure] = faits[r.heure] || []).indexOf(r.type) < 0 && faits[r.heure].push(r.type);
  });
  return { date: date, heure: Number(Utilities.formatDate(now, TZ, 'H')), heures: PH_HEURES, heures_vues: PH_HEURES_VUES, types: phTypesDe_(code), faits: faits, suivie: PH_NON_SUIVIES.indexOf(code) < 0 };
}

/** Bilan d'une journée pour une station : heures sans photo et types manquants (seulement pour les heures déjà passées). */
function phBilanJour_(parHeure, types, date, auj, hMaintenant) {
  var manq = [], typesManq = [], couvertes = 0, echues = 0;
  PH_HEURES.forEach(function (h) {
    var passee = date < auj || (date === auj && hMaintenant > h);
    var la = parHeure[h] || {};
    var nb = PH_TYPES.filter(function (t) { return la[t]; }).length;
    if (nb) couvertes++;
    if (!passee) return;
    echues++;
    if (!nb) { manq.push(h); return; }
    types.forEach(function (t) { if (!la[t]) typesManq.push({ h: h, type: t }); });
  });
  return { manquantes: manq, typesManquants: typesManq, couvertes: couvertes, echues: echues };
}

/** Direction / comptabilité : suivi du jour pour toutes les stations + total du mois. */
function photosSuivi_(p) {
  var now = new Date(), auj = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), hM = Number(Utilities.formatDate(now, TZ, 'H'));
  var date = /^\d{4}-\d{2}-\d{2}$/.test(String(p.date || '')) ? String(p.date) : auj, mois = date.slice(0, 7);
  var lignes = phLire_(15000), idx = {}, premier = {}, jg = {};
  lignes.forEach(function (r) {
    var k = r.station + '|' + r.date;
    if (PH_JAUGES.indexOf(r.type) >= 0) {   // ouverture = la première de la journée ; fermeture = la dernière
      var j = jg[k] = jg[k] || {}, hm = r.pris_le.slice(11, 16);
      if (r.type === 'JAUGE_OUV' && (!j.ouv || hm < j.ouv)) { j.ouv = hm; j.ouvNiveau = r.compteur; }
      if (r.type === 'JAUGE_FER' && (!j.fer || hm > j.fer)) { j.fer = hm; j.ferNiveau = r.compteur; }
    }
    ((idx[k] = idx[k] || {})[r.heure] = idx[k][r.heure] || {})[r.type] = (idx[k][r.heure][r.type] || 0) + 1;
    if (!premier[r.station] || r.date < premier[r.station]) premier[r.station] = r.date;
  });
  var finMois = date < auj ? (auj.slice(0, 7) === mois ? auj : mois + '-31') : auj;
  var out = stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) {
    var code = String(s.code), types = phTypesDe_(code), jour = idx[code + '|' + date] || {};
    var b = phBilanJour_(jour, types, date, auj, hM), grille = {}, nbPhotos = 0;
    PH_HEURES_VUES.forEach(function (h) { grille[h] = PH_TYPES.filter(function (t) { return (jour[h] || {})[t]; }).map(function (t) { return PH_LETTRE[t]; }).join(''); });
    Object.keys(jour).forEach(function (h) { Object.keys(jour[h]).forEach(function (t) { nbPhotos += jour[h][t]; }); });
    // Mois : compté à partir du premier jour où la station a envoyé une photo par l'application.
    var m = { depuis: '', jours: 0, heuresManquantes: 0, typesManquants: 0 };
    if (premier[code] && premier[code].slice(0, 7) <= mois) {
      var d = premier[code] > mois + '-01' ? premier[code] : mois + '-01';
      m.depuis = d;
      while (d <= finMois && d.slice(0, 7) === mois) {
        var bj = phBilanJour_(idx[code + '|' + d] || {}, types, d, auj, hM);
        m.jours++; m.heuresManquantes += bj.manquantes.length; m.typesManquants += bj.typesManquants.length;
        d = Utilities.formatDate(new Date(new Date(d + 'T12:00:00Z').getTime() + 864e5), 'UTC', 'yyyy-MM-dd');
      }
    }
    var j = jg[code + '|' + date] || {};
    var jauge = { ouv: j.ouv || '', ouvOk: !!j.ouv && j.ouv < '12:00', ouvNiveau: j.ouvNiveau || '', fer: j.fer || '', ferOk: !!j.fer && j.fer >= '15:00', ferNiveau: j.ferNiveau || '',
      ouvManque: !(j.ouv && j.ouv < '12:00') && (date < auj || hM >= 12), ferManque: !(j.fer && j.fer >= '15:00') && (date < auj || hM >= 18) };
    return { code: code, nom: String(s.nom), suivie: PH_NON_SUIVIES.indexOf(code) < 0, debut: premier[code] || '', types: types, grille: grille, nbPhotos: nbPhotos, jauge: jauge,
      couvertes: b.couvertes, echues: b.echues, manquantes: b.manquantes, typesManquants: b.typesManquants, mois: m };
  });
  return { date: date, aujourdhui: auj, heures: PH_HEURES_VUES, comptees: PH_HEURES, stations: out };
}

/** Liste des photos d'une station pour un jour. */
function photosListe_(p) {
  var date = String(p.date || ''), code = String(p.station || '');
  var l = phLire_(15000).filter(function (r) { return r.station === code && r.date === date; })
    .sort(function (a, b) { return a.pris_le < b.pris_le ? -1 : 1; })
    .map(function (r) { return { heure: r.heure, type: r.type, pris_le: r.pris_le, recu_le: r.recu_le, compteur: r.compteur, id: r.fichier_id, note: r.note }; });
  return { date: date, station: code, photos: l };
}

/** Renvoie une photo (seulement un fichier listé dans l'onglet « Photos horaires »). */
function photoVoir_(p) {
  var id = String(p.id || '');
  if (!id || !(phLire_(15000).some(function (r) { return r.fichier_id === id; }) || bdLire_(5000).some(function (r) { return r.fichier_id === id; }) || camOps_(8000).some(function (r) { return r.fichier_id === id; }) || argLire_(10000).some(function (r) { return r.fichier_id === id; }) || camRapports_(3000).some(function (r) { return r.fichier_id === id; }))) throw new UserError('Photo introuvable.');
  var f = DriveApp.getFileById(id), b = f.getBlob();
  return { nom: f.getName(), photo: 'data:image/jpeg;base64,' + Utilities.base64Encode(b.getBytes()) };
}

/* ============================ Bordereaux de dépôt (m1.13, 26/11/2026) ============================
 * La station photographie le bordereau (fiche de dépôt, reçu de caisse) et écrit le montant tel qu'il est sur le papier.
 * Une ligne par bordereau dans l'onglet « Bordereaux (app) », photo dans un dossier Drive privé. Rien n'est envoyé aux fournisseurs d'ici.
 */
var BD_ONGLET = 'Bordereaux (app)';
var BD_ENTETES = ['date_depot', 'station', 'montant', 'devise', 'beneficiaire', 'numero', 'note', 'pris_le', 'recu_le', 'fichier_id', 'appareil', 'cid'];
var BD_BENEF = ['Haiti Propane', 'SODIGAZ', 'Compte Red Rebel', 'Autre'];
var BD_DEVISES = ['HTG', 'USD'];
function bdLire_(max) {
  var sh = ss_().getSheetByName(BD_ONGLET);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = sh.getLastRow(), from = Math.max(2, n - (max || 5000) + 1);
  return sh.getRange(from, 1, n - from + 1, BD_ENTETES.length).getValues().map(function (r) {
    return { date_depot: phJour_(r[0]), station: String(r[1]), montant: Number(r[2]) || 0, devise: String(r[3]), beneficiaire: String(r[4]), numero: String(r[5] || ''), note: String(r[6] || ''),
      pris_le: phTexte_(r[7]), recu_le: phTexte_(r[8]), fichier_id: String(r[9]), appareil: String(r[10]), cid: String(r[11]) };
  });
}
function bordereau_(app, p) {
  if (app.role || !app.station) throw new UserError('Les bordereaux se font depuis le téléphone de la station.', 'FORBIDDEN');
  var st = app.station, cid = String(p.cid || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 60);
  if (!cid) throw new UserError('Bordereau sans identifiant.');
  var m = String(p.photo || '').match(/^data:image\/jpeg;base64,(.+)$/);
  if (!m) throw new UserError('Photo du bordereau invalide.');
  var octets = Utilities.base64Decode(m[1]);
  if (octets.length > 3000000) throw new UserError('Photo trop lourde.');
  var mt = phNombre_(p.montant);
  if (!mt || !(Number(mt) > 0)) throw new UserError('Montant du bordereau illisible. Écrivez seulement les chiffres.');
  var devise = BD_DEVISES.indexOf(String(p.devise)) >= 0 ? String(p.devise) : 'HTG';
  var benef = BD_BENEF.indexOf(String(p.beneficiaire)) >= 0 ? String(p.beneficiaire) : 'Autre';
  var recu = new Date(), auj = Utilities.formatDate(recu, TZ, 'yyyy-MM-dd'), pris = new Date(Number(p.pris_ms) || 0);
  if (!(pris.getTime() > 0) || pris.getTime() > recu.getTime() + 10 * 60000) pris = recu;
  var date = String(p.date_depot || ''), note = String(p.note || '').slice(0, 200);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date > auj) date = auj;   // jamais dans le futur
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var sh = ss_().getSheetByName(BD_ONGLET);
    if (!sh) { sh = ss_().insertSheet(BD_ONGLET); sh.appendRow(BD_ENTETES); sh.setFrozenRows(1); }
    var n = sh.getLastRow();
    if (n > 1) {
      var from = Math.max(2, n - 1999), cids = sh.getRange(from, 12, n - from + 1, 1).getValues();
      for (var i = 0; i < cids.length; i++) if (String(cids[i][0]) === cid) return { recu: true, deja: true, cid: cid };
    }
    var nomF = st.code + '_bordereau_' + date + '_' + mt.replace('.', '-') + devise + '_' + Utilities.formatDate(pris, TZ, 'HHmmss') + '.jpg';
    var f = phDossier_(st, date, 'BD', 'RRD – Bordereaux de dépôt (app, privé)').createFile(Utilities.newBlob(octets, 'image/jpeg', nomF));
    f.setDescription('Bordereau de dépôt (application RRD) — ' + st.nom + ' — ' + mt + ' ' + devise + ' — ' + benef + ' — ' + cid);
    sh.appendRow([date, st.code, Number(mt), devise, benef, String(p.numero || '').slice(0, 40), note, Utilities.formatDate(pris, TZ, 'yyyy-MM-dd HH:mm:ss'),
      Utilities.formatDate(recu, TZ, 'yyyy-MM-dd HH:mm:ss'), f.getId(), app.id, cid]);
    return { recu: true, cid: cid };
  } finally { lock.releaseLock(); }
}
/** Direction / comptabilité : bordereaux reçus entre deux dates de dépôt (7 derniers jours par défaut), avec les totaux par devise. */
function bordereauxListe_(p) {
  var auj = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), ok = function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d || '')); };
  var au = ok(p.au) ? String(p.au) : auj, du = ok(p.du) ? String(p.du) : Utilities.formatDate(new Date(new Date(au + 'T12:00:00Z').getTime() - 6 * 864e5), 'UTC', 'yyyy-MM-dd');
  var noms = {}; stations_().forEach(function (s) { noms[String(s.code)] = String(s.nom); });
  var totaux = {}, l = bdLire_(5000).filter(function (r) { return r.date_depot >= du && r.date_depot <= au && (!p.station || r.station === String(p.station)); })
    .sort(function (a, b) { return a.date_depot === b.date_depot ? (a.recu_le < b.recu_le ? 1 : -1) : (a.date_depot < b.date_depot ? 1 : -1); });
  l.forEach(function (r) { totaux[r.devise] = (totaux[r.devise] || 0) + r.montant; });
  return { du: du, au: au, aujourdhui: auj, totaux: totaux, bordereaux: l.map(function (r) {
    return { date: r.date_depot, station: r.station, nom: noms[r.station] || r.station, montant: r.montant, devise: r.devise, beneficiaire: r.beneficiaire, numero: r.numero, note: r.note, recu_le: r.recu_le, id: r.fichier_id };
  }) };
}

/** Texte ajouté au résumé du soir (courriel) : photos horaires, manomètre et bordereaux du jour. */
function resumeApp_(date) {
  var t = '', s = photosSuivi_({ date: date }).stations.filter(function (x) { return x.debut && x.debut <= date; });
  if (s.length) {
    t += '\n\n📷 Photos — ' + date.split('-').reverse().join('/') + '\n' + s.map(function (x) {
      var pb = x.jauge.ouvManque || x.jauge.ferManque || (x.suivie && (x.manquantes.length || x.typesManquants.length));
      return (pb ? '⚠️ ' : '✅ ') + x.nom + ' : ' + (x.suivie ? x.manquantes.length + ' heure(s) sans photo' +
        (x.manquantes.length ? ' (' + x.manquantes.map(function (h) { return h + ' h'; }).join(', ') + ')' : '') + ', ' + x.typesManquants.length + ' photo(s) manquante(s) ; ' : '') +
        'manomètre ouverture ' + (x.jauge.ouv ? x.jauge.ouv + (x.jauge.ouvOk ? '' : ' (après midi)') + (x.jauge.ouvNiveau ? ' ' + x.jauge.ouvNiveau + ' %' : '') : 'manquante') +
        ', fermeture ' + (x.jauge.fer ? x.jauge.fer + (x.jauge.ferOk ? '' : ' (avant 15 h)') + (x.jauge.ferNiveau ? ' ' + x.jauge.ferNiveau + ' %' : '') : 'manquante') +
        (x.suivie ? ' — mois : ' + x.mois.heuresManquantes + ' heure(s) sans photo, ' + x.mois.typesManquants + ' photo(s) manquante(s)' : '');
    }).join('\n');
  }
  var noms = {}; stations_().forEach(function (x) { noms[String(x.code)] = String(x.nom); });
  var b = bdLire_(5000).filter(function (r) { return r.recu_le.slice(0, 10) === date; });
  if (b.length) t += '\n\n🧾 Bordereaux reçus le ' + date.split('-').reverse().join('/') + '\n' + b.map(function (r) {
    return '• ' + (noms[r.station] || r.station) + ' : ' + r.montant.toLocaleString('fr-FR') + ' ' + r.devise + ' — ' + r.beneficiaire + (r.numero ? ' — n° ' + r.numero : '') +
      (r.date_depot !== date ? ' — dépôt du ' + r.date_depot.split('-').reverse().join('/') : '');
  }).join('\n');
  try {
    var stk = stockEstime_().stations.filter(function (x) { return x.dernier; });
    if (stk.length) t += '\n\n🛢 Stock estimé (reçu − vendu, d\'après les rapports)\n' + stk.map(function (x) {
      return (x.alerte ? '⚠️ ' : '• ') + x.nom + ' : ' + (x.methode === 'plus_bas' ? 'au moins ' : 'environ ') + x.stock.toLocaleString('fr-FR') + ' gal (' + x.pct + ' %) — rapports jusqu\'au ' +
        x.dernier.split('-').reverse().join('/') + (x.alerte ? ' — à vérifier' : '');
    }).join('\n');
  } catch (e) { console.error(e && e.stack || e); }   // le stock ne doit jamais empêcher le résumé de partir
  try {
    var jr = date, cr = camRapportsVue_(jr, jr, '');
    if (cr.length) t += '\n\n🚚 Camions (' + jr.split('-').reverse().join('/') + ', ' + CURRENCY + ')\n' + cr.map(function (x) {
      var k = x.totaux, n = function (v) { return Number(v).toLocaleString('fr-FR'); };
      return '• ' + x.camion_nom + ' : chargé ' + n(k.charge) + ' gal, livré ' + n(k.livre) + ' gal (' + k.nb + ' livraison(s)), reçu ' + n(k.recu + k.paiements) + ', crédit ' + n(k.credit) + ', achats ' + n(k.achats) +
        (k.sans_prix ? ', ' + k.sans_prix + ' livraison(s) sans prix' : '') + (x.fait ? ' — rapport fait : argent ' + n(x.remis) + (x.ou ? ' (' + [x.ou.depose ? 'déposé ' + n(x.ou.depose) + ' par ' + x.ou.depose_par : '', x.ou.bureau ? 'au bureau ' + n(x.ou.bureau) + ' reçu par ' + x.ou.bureau_qui : '', x.ou.main ? 'en main ' + n(x.ou.main) + ' gardé par ' + x.ou.main_qui : ''].filter(String).join(' ; ') + ')' : '') + (x.difference ? ', différence ' + n(x.difference) : '') : ' — ⚠️ rapport du jour pas encore fait');
    }).join('\n');
  } catch (e) { console.error(e && e.stack || e); }   // les camions non plus
  return t;
}

/* ============================ Stock estimé de propane par station (m1.18, 26/11/2026) ============================
 * Stock = propane reçu − propane vendu depuis la dernière fois que la station était à sec (décision du propriétaire, 26/11/2026).
 * Sources, jour par jour : les rapports de l'application (compteur et « gaz reçu ») ; pour les jours sans rapport dans l'application,
 * l'historique de l'Activity Log (« Daily Station Reports » : gallons vendus, gaz acheté ; « Truck Deliveries » : livraisons du camion).
 * Pour les livraisons d'un jour, on garde le plus grand des chiffres connus (rapport, colonne automatique, camion).
 * Sans date « à sec » connue, le calcul part du point le plus bas des rapports depuis STK_DEBUT : c'est alors un minimum.
 * Réglages (taille de la citerne, date « à sec ») : onglet « Stock (réglages) », modifiable par la direction dans le tableau de bord.
 * Rien n'est inventé : si les rapports ne notent pas une livraison, le calcul est faux et l'alerte le dit. */
var STK_ONGLET = 'Stock (réglages)';
var STK_ENTETES = ['station', 'capacite_gal', 'vide_le', 'source', 'modifie_le', 'modifie_par'];
var STK_DEBUT = '2026-06-01';
var STK_DEFAUT = {   // tailles données par le propriétaire le 26/11/2026 ; dates « à sec » trouvées dans l'historique (à confirmer par la direction)
  BRO: [1000, '2026-10-05', 'Historique : aucune vente le 05/10/2026 (à confirmer)'],
  A32: [2000, '2026-08-08', 'Historique : aucune vente le 08/08/2026 (à confirmer)'],
  DED: [1000, '', ''],
  M83: [1000, '', ''],
  LEO: [1000, '', ''],
  DIQ: [1000, '2026-10-04', 'Historique, note du rapport du 04/10/2026 : station à sec à 12 h 06'],
  ESA: [1000, '', '']
};
function stkCode_(nom) {
  var n = String(nom || '').toLowerCase();
  if (/naupi/.test(n)) return '';   // client extérieur
  if (/broch/.test(n)) return 'BRO';
  if (/arcachon/.test(n)) return 'A32';
  if (/ded|d\.e\.d|lamentin|raile|\bray\b/.test(n)) return 'DED';
  if (/mahot|rrd ?83|waney/.test(n)) return 'M83';
  if (/l[ée]og|aka|bon zanmi/.test(n)) return 'LEO';
  if (/diquini/.test(n)) return 'DIQ';
  if (/esans|thor 67/.test(n)) return 'ESA';
  return '';
}
function stkNum_(v) { var n = Number(String(v === null || v === undefined ? '' : v).replace(/[\s\u00a0\u202f,]/g, '')); return isFinite(n) ? n : 0; }
function stkFeuille_() {
  var sh = ss_().getSheetByName(STK_ONGLET);
  if (!sh) {
    sh = ss_().insertSheet(STK_ONGLET); sh.appendRow(STK_ENTETES); sh.setFrozenRows(1); sh.getRange('C:C').setNumberFormat('@');
    Object.keys(STK_DEFAUT).forEach(function (c) { sh.appendRow([c, STK_DEFAUT[c][0], STK_DEFAUT[c][1], STK_DEFAUT[c][2], maintenant_(), 'Claude (historique)']); });
  }
  return sh;
}
function stkReglages_() {
  var sh = stkFeuille_(), o = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r, i) {
    var c = String(r[0]).trim(); if (!c) return;
    o[c] = { ligne: i + 2, capacite: stkNum_(r[1]) || 1000, vide_le: /^\d{4}-\d{2}-\d{2}/.test(phJour_(r[2])) ? phJour_(r[2]) : '', source: String(r[3] || '') };
  });
  return o;
}
/** Historique de l'Activity Log depuis STK_DEBUT : { code: { date: [vendu, recu_rapport, recu_auto, recu_camion] } } (gardé 6 h en mémoire). */
function stkHistorique_() {
  var cache = CacheService.getScriptCache(), c = cache.get('stk_hist_v1');
  if (c) return JSON.parse(c);
  var log = SpreadsheetApp.openById(DSR_ACTIVITY_LOG_ID), tz = log.getSpreadsheetTimeZone(), h = {};
  var jour = function (v) { return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').slice(0, 10); };
  var cas = function (code, d) { var s = h[code] = h[code] || {}; return s[d] = s[d] || [0, 0, 0, 0]; };
  var sh = log.getSheetByName('Daily Station Reports'), n = sh ? sh.getLastRow() : 0;
  if (n > 1) sh.getRange(2, 1, n - 1, 16).getValues().forEach(function (r) {
    var d = jour(r[0]), code = stkCode_(r[1]);
    if (!code || !/^\d{4}-\d{2}-\d{2}$/.test(d) || d < STK_DEBUT) return;
    var x = cas(code, d); x[0] += stkNum_(r[4]); x[1] += stkNum_(r[14]); x[2] = Math.max(x[2], stkNum_(r[15]));
  });
  sh = log.getSheetByName('Truck Deliveries'); n = sh ? sh.getLastRow() : 0;
  if (n > 1) sh.getRange(2, 1, n - 1, 5).getValues().forEach(function (r) {
    var d = jour(r[0]), code = stkCode_(r[3]);
    if (!code || !/^\d{4}-\d{2}-\d{2}$/.test(d) || d < STK_DEBUT) return;
    cas(code, d)[3] += stkNum_(r[4]);
  });
  Object.keys(h).forEach(function (k) { Object.keys(h[k]).forEach(function (d) { h[k][d] = h[k][d].map(function (v) { return Math.round(v * 10) / 10; }); }); });
  try { cache.put('stk_hist_v1', JSON.stringify(h), 21600); } catch (e) { }
  return h;
}
function stockEstime_() {
  var regl = stkReglages_(), hist = stkHistorique_(), app = {}, noms = {};
  stations_().forEach(function (s) { noms[String(s.code)] = String(s.nom); });
  rapportsEffectifs_().forEach(function (r) {
    if (estAnnule(r)) return;
    var code = String(r.station), d = iso(r.date_rapport);
    if (!regl[code] || !/^\d{4}-\d{2}-\d{2}$/.test(d) || d < STK_DEBUT) return;
    var x = (app[code] = app[code] || {})[d] = (app[code][d] || [0, 0]);
    var g = num(r.closing_meter) - num(r.opening_meter);
    x[0] += isNaN(g) ? 0 : g; x[1] += stkNum_(r.gas_received);
  });
  var camRecu = {};   // livraisons des camions (application) à un client relié à une station
  try {
    var liens = {}; camClients_().forEach(function (c) { if (c.station) liens[c.id] = c.station; });
    camOps_(8000).forEach(function (o) {
      var stn = liens[o.client_id];
      if (o.annule || o.genre !== 'livraison' || o.camion === CAM_TEST || !stn || !regl[stn]) return;
      (camRecu[stn] = camRecu[stn] || {})[o.date] = ((camRecu[stn] || {})[o.date] || 0) + (Number(o.gallons) || 0);
    });
  } catch (e) { }
  var jauges = {};   // dernier niveau du manomètre tapé dans l'application
  phLire_(3000).forEach(function (r) {
    if (PH_JAUGES.indexOf(r.type) < 0 || !regl[r.station] || !(Number(r.compteur) > 0)) return;
    var j = jauges[r.station];
    if (!j || r.pris_le > j.pris_le) jauges[r.station] = { pris_le: r.pris_le, pct: Number(r.compteur) };
  });
  var out = Object.keys(regl).map(function (code) {
    var g = regl[code], hh = hist[code] || {}, aa = app[code] || {}, cc = camRecu[code] || {}, jours = {};
    Object.keys(hh).forEach(function (d) { jours[d] = 1; }); Object.keys(aa).forEach(function (d) { jours[d] = 1; }); Object.keys(cc).forEach(function (d) { jours[d] = 1; });
    var ds = Object.keys(jours).sort(), serie = ds.map(function (d) {
      var x = hh[d] || [0, 0, 0, 0], a = aa[d];
      return { d: d, vendu: a ? a[0] : x[0], recu: Math.max(a ? a[1] : 0, x[1], x[2], x[3], cc[d] || 0) };
    });
    var methode = g.vide_le ? 'vide' : 'plus_bas', depuis = g.vide_le, cum = 0, mn = 0, mx = 0, recu = 0, vendu = 0;
    if (methode === 'plus_bas') {   // point le plus bas du solde (reçu − vendu) : on suppose la citerne vide ce jour-là
      var c0 = 0, bas = 0; depuis = ds.length ? ds[0] : '';
      serie.forEach(function (x) { c0 += x.recu - x.vendu; if (c0 <= bas) { bas = c0; depuis = x.d; } });
    }
    serie.forEach(function (x) {
      if (x.d <= depuis) return;
      recu += x.recu; vendu += x.vendu; cum += x.recu - x.vendu;
      if (cum < mn) mn = cum; if (cum > mx) mx = cum;
    });
    var alerte = '';
    if (mx > g.capacite + 50) alerte = 'Le calcul dépasse la taille de la citerne : une livraison est peut-être comptée deux fois, ou des ventes manquent.';
    else if (mn < -50) alerte = 'Le calcul passe sous zéro : des livraisons manquent dans les rapports, ou la date « à sec » est à revoir.';
    var j = jauges[code], av = serie.filter(function (x) { return x.vendu > 0; }).slice(-7), vj = 0;   // ventes moyennes des 7 derniers jours avec ventes
    av.forEach(function (x) { vj += x.vendu; });
    return { code: code, nom: noms[code] || code, vente_jour: av.length ? Math.round(vj / av.length) : 0, jours_vente: av.length, capacite: g.capacite, methode: methode, depuis: depuis, source: g.source, recu: Math.round(recu), vendu: Math.round(vendu),
      stock: Math.round(cum), pct: Math.round(100 * cum / g.capacite), alerte: alerte, dernier: ds.length ? ds[ds.length - 1] : '',
      jauge: j ? { le: j.pris_le, pct: j.pct, gal: Math.round(j.pct * g.capacite / 100) } : null };
  });
  return { calcule_le: maintenant_(), stations: out };
}
function stockRegler_(ctx, p) {
  var regl = stkReglages_(), code = String(p.station || ''), g = regl[code];
  if (!g) throw new UserError('Station inconnue.');
  var d = String(p.vide_le || '').trim();
  if (d && (!/^\d{4}-\d{2}-\d{2}$/.test(d) || d > msgJour_(0))) throw new UserError('Date invalide.');
  var cap = p.capacite === undefined || p.capacite === '' ? g.capacite : stkNum_(p.capacite);
  if (!(cap >= 100 && cap <= 100000)) throw new UserError('Taille de citerne invalide.');
  var sh = stkFeuille_();
  sh.getRange(g.ligne, 2, 1, 5).setValues([[cap, d, d ? 'Saisi par la direction' : '', maintenant_(), ctx.nom]]);
  journal_(ctx, 'stock-reglage', code + ' vide_le=' + (d || '(aucune)') + ' capacite=' + cap + ' (avant : ' + (g.vide_le || '(aucune)') + ', ' + g.capacite + ')');
  return stockEstime_();
}

/* ============================ Messages et conseils aux stations (m1.14, 26/11/2026) ============================
 * La direction écrit un conseil du jour ou un message, pour toutes les stations ou une seule, avec une date et une heure.
 * Le téléphone de la station le prend à sa prochaine ouverture de l'application, puis l'affiche en notification à l'heure prévue
 * (même sans signal). Une ligne par message dans l'onglet « Messages (app) » ; « supprimer » le désactive, la ligne reste.
 */
var MSG_ONGLET = 'Messages (app)';
var MSG_ENTETES = ['id', 'cree_le', 'par', 'cible', 'date', 'heure', 'type', 'texte', 'actif', 'pousse_le', 'pousse_a'];   // les deux dernières : remplies par le relais des notifications
function msgFeuille_(creer) {
  var sh = ss_().getSheetByName(MSG_ONGLET);
  if (!sh && creer) { sh = ss_().insertSheet(MSG_ONGLET); sh.appendRow(MSG_ENTETES); sh.setFrozenRows(1); }
  else if (sh && creer && sh.getLastColumn() < MSG_ENTETES.length) sh.getRange(1, 1, 1, MSG_ENTETES.length).setValues([MSG_ENTETES]);
  return sh;
}
function msgLire_() {
  var sh = msgFeuille_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = sh.getLastRow(), from = Math.max(2, n - 1999);
  return sh.getRange(from, 1, n - from + 1, MSG_ENTETES.length).getValues().map(function (r, i) {
    return { ligne: from + i, id: Number(r[0]) || 0, cree_le: phTexte_(r[1]), par: String(r[2]), cible: String(r[3]).trim(), date: phJour_(r[4]),
      heure: String(r[5]).replace('h', ':'), type: String(r[6]), texte: String(r[7]).trim(), actif: /^(oui|true)$/i.test(String(r[8]).trim()),
      pousse_a: String(r[10] || '').split(',').map(function (x) { return x.trim(); }).filter(String) };
  });
}
function msgJour_(n) { return Utilities.formatDate(new Date(Date.now() + n * 864e5), TZ, 'yyyy-MM-dd'); }
/** m1.38 (27/11/2026) : les conseils aussi pour les chauffeurs. Pour qui est un message : « TOUS » = toutes les stations (pas les camions),
 * « CAMIONS » = tous les camions, « CAM:1 » = un camion, sinon le code d'une station. msgCode_ : le code qui reçoit et répond, ou ''. */
function msgCode_(app) {
  if (app.role === 'camion' && app.camion) return 'CAM:' + app.camion.code;
  return (!app.role && app.station) ? String(app.station.code) : '';
}
function msgPour_(m, code) {
  if (code.indexOf('CAM:') === 0) return m.cible === 'CAMIONS' || m.cible === code;
  return m.cible === 'TOUS' || m.cible === code;
}
/** Pour le téléphone d'une station ou d'un camion : ses messages actifs, d'avant-hier à dans 14 jours. */
function messagesStation_(app) {
  var code = msgCode_(app), du = msgJour_(-2), au = msgJour_(14);
  if (!code) throw new UserError('Réservé au téléphone d\'une station ou d\'un camion.', 'FORBIDDEN');
  var miennes = {}; repLire_().forEach(function (r) { if (r.station === code) miennes[r.msg_id] = r.texte; });   // la dernière réponse de cette station à chaque message
  var l = msgLire_().filter(function (m) { return m.actif && msgPour_(m, code) && m.date >= du && m.date <= au; })
    .map(function (m) { return { id: m.id, date: m.date, heure: m.heure, type: m.type, texte: m.texte, pousse: m.pousse_a.indexOf(code) >= 0, reponse: miennes[m.id] || '' }; });
  try { majColonne_(APPAREILS, app.ligne._ligne, 'msg_vu_le', maintenant_()); } catch (e) { }   // ce téléphone a maintenant ces messages : inutile de les lui pousser
  return { date: msgJour_(0), messages: l };
}

/* ============================ Notifications poussées (m1.15, 26/11/2026) ============================
 * Le téléphone d'une station donne son adresse de notification (Firebase) ; elle est gardée dans la colonne « push » de l'onglet
 * « Appareils ». Un petit programme à part, « RRD Notifications », lit l'onglet « Messages (app) » et pousse à l'heure prévue les
 * messages qu'un téléphone n'a pas encore pris (« msg_vu_le » plus ancien que le message), puis note « pousse_le » et « pousse_a ».
 * Si ce relais est arrêté, rien n'est perdu : le message s'affiche à la prochaine ouverture de l'application, comme avant. */
function majColonne_(def, ligne, champ, valeur) {   // comme majCellule_, mais ajoute la colonne si elle manque
  var sh = sh_(def), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String), j = h.indexOf(champ);
  if (j < 0) { j = h.length; sh.getRange(1, j + 1).setValue(champ); }
  sh.getRange(ligne, j + 1).setValue(valeur);
}
function pushToken_(app, p) {   // m1.34 : aussi camions, direction et comptabilité (notifications du chat) ; pas les agents de transfert
  if (app.role === 'transferts' || (!app.role && !app.station)) return { garde: false };
  var t = String(p.token || '').trim();
  if (t && (t.length > 1024 || !/^[A-Za-z0-9_:.\-]+$/.test(t))) throw new UserError('Adresse de notification invalide.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    lire_(APPAREILS).forEach(function (a) {   // une adresse = un seul téléphone : on la retire d'un ancien code
      if (t && a._ligne !== app.ligne._ligne && String(a.push || '') === t) majColonne_(APPAREILS, a._ligne, 'push', '');
    });
    if (String(app.ligne.push || '') !== t) majColonne_(APPAREILS, app.ligne._ligne, 'push', t);
  } finally { lock.releaseLock(); }
  return { garde: !!t };
}
/** Réveille le relais tout de suite (sinon il passe toutes les 5 minutes). Sans effet tant que l'adresse du relais n'est pas réglée. */
function pushReveiller_() {
  try {
    var url = prop_('PUSH_RELAIS_URL');
    if (url) UrlFetchApp.fetch(url, { method: 'post', payload: '{}', contentType: 'text/plain', muteHttpExceptions: true });
  } catch (e) { }
}
function messagesListe_() {
  var noms = { TOUS: 'Toutes les stations', CAMIONS: '🚚 Tous les camions' }; stations_().forEach(function (s) { noms[String(s.code)] = String(s.nom); });
  var cams = camions_().filter(function (c) { return c.actif; }); cams.forEach(function (c) { noms['CAM:' + c.code] = '🚚 ' + c.nom; });
  var du = msgJour_(-30), parMsg = {};
  repLire_().forEach(function (r) { (parMsg[r.msg_id] = parMsg[r.msg_id] || []).push({ station: r.station, nom: noms[r.station] || r.station, ecrit_le: r.ecrit_le, texte: r.texte }); });
  return { aujourdhui: msgJour_(0), stations: stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return { code: String(s.code), nom: String(s.nom) }; }),
    camions: cams.map(function (c) { return { code: 'CAM:' + c.code, nom: '🚚 ' + c.nom }; }),
    messages: msgLire_().filter(function (m) { return m.actif && m.date >= du; }).sort(function (a, b) { return (a.date + a.heure) < (b.date + b.heure) ? 1 : -1; })
      .map(function (m) { return { id: m.id, date: m.date, heure: m.heure, type: m.type, texte: m.texte, cible: m.cible, cible_nom: noms[m.cible] || m.cible, par: m.par, reponses: parMsg[m.id] || [] }; }) };
}

/* ============================ Réponses des stations aux messages (m1.17, 26/11/2026) ============================
 * La station répond par une phrase au conseil du jour ou à un message (c'est le point « ✅ » du RRD Challenge, et la réponse
 * à la question du samedi). Une ligne par réponse dans l'onglet « Réponses (app) » ; rien n'est jamais modifié ni supprimé.
 * « ecrit_le » = l'heure du téléphone (heure d'Haïti) au moment où la réponse a été écrite, même si elle part plus tard faute de signal. */
var REP_ONGLET = 'Réponses (app)';
var REP_ENTETES = ['id', 'recu_le', 'station', 'msg_id', 'date_msg', 'ecrit_le', 'texte', 'appareil', 'cid', 'note'];
function repFeuille_(creer) {
  var sh = ss_().getSheetByName(REP_ONGLET);
  if (!sh && creer) { sh = ss_().insertSheet(REP_ONGLET); sh.appendRow(REP_ENTETES); sh.setFrozenRows(1); }
  return sh;
}
function repLire_() {
  var sh = repFeuille_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = sh.getLastRow(), from = Math.max(2, n - 2999);
  return sh.getRange(from, 1, n - from + 1, REP_ENTETES.length).getValues().map(function (r) {
    return { id: Number(r[0]) || 0, recu_le: phTexte_(r[1]), station: String(r[2]).trim(), msg_id: Number(r[3]) || 0, date_msg: phJour_(r[4]), ecrit_le: phTexte_(r[5]),
      texte: String(r[6]).trim(), cid: String(r[8]) };
  });
}
function reponseStation_(app, p) {
  var code = msgCode_(app);
  if (!code) throw new UserError('Réservé au téléphone d\'une station ou d\'un camion.', 'FORBIDDEN');
  var texte = String(p.texte || '').trim().slice(0, 500), cid = String(p.cid || '').slice(0, 40);
  if (texte.length < 3) throw new UserError('Écrivez votre réponse.');
  var m = msgLire_().filter(function (x) { return x.id === Number(p.msg_id); })[0];
  if (!m || !m.actif || !msgPour_(m, code)) throw new UserError('Ce message n\'existe plus.');
  var recu = new Date(), ecrit = new Date(Number(p.ecrit_ms) || 0), note = '';
  if (!(ecrit.getTime() > 0) || ecrit.getTime() > recu.getTime() + 10 * 60000) { note = 'heure du téléphone refusée : heure de réception utilisée'; ecrit = recu; }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var l = repLire_();
    if (cid && l.some(function (r) { return r.cid === cid; })) return { recu: true };   // déjà reçue (nouvel essai du téléphone)
    var id = l.reduce(function (mx, r) { return Math.max(mx, r.id); }, 0) + 1;
    repFeuille_(true).appendRow([id, Utilities.formatDate(recu, TZ, 'yyyy-MM-dd HH:mm:ss'), code, m.id, m.date, Utilities.formatDate(ecrit, TZ, 'yyyy-MM-dd HH:mm:ss'),
      (/^[=+\-@]/.test(texte) ? ' ' : '') + texte, app.id, cid, note]);   // jamais une formule
  } finally { lock.releaseLock(); }
  return { recu: true };
}
function messageAjouter_(ctx, p) {
  var texte = String(p.texte || '').trim();
  if (!texte) throw new UserError('Écrivez le texte du message.');
  if (texte.length > 1500) throw new UserError('Message trop long (1 500 caractères au plus).');
  var auj = msgJour_(0), date = /^\d{4}-\d{2}-\d{2}$/.test(String(p.date || '')) ? String(p.date) : auj;
  var heure = /^([01]\d|2[0-3]):[0-5]\d$/.test(String(p.heure || '')) ? String(p.heure) : '06:30';
  if (p.maintenant === true) { date = auj; heure = Utilities.formatDate(new Date(), TZ, 'HH:mm'); }   // « Envoyer maintenant » : l'heure du bureau, pas celle de l'appareil
  if (date < auj) throw new UserError('La date est déjà passée.');
  var cible = String(p.cible || 'TOUS').trim();
  if (cible !== 'TOUS' && cible !== 'CAMIONS' && !(cible.indexOf('CAM:') === 0 ? !!camDe_(cible) : stations_().some(function (s) { return String(s.code) === cible && estVrai_(s.actif); }))) throw new UserError('Station ou camion inconnu.');
  var type = p.type === 'conseil' ? 'conseil' : 'message';
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var id = msgLire_().reduce(function (mx, m) { return Math.max(mx, m.id); }, 0) + 1;
    msgFeuille_(true).appendRow([id, maintenant_(), ctx.nom, cible, date, heure.replace(':', 'h'), type, (/^[=+\-@]/.test(texte) ? ' ' : '') + texte, 'oui']);   // jamais une formule
    journal_(ctx, 'message+', id + ' ' + cible + ' ' + date + ' ' + heure);
  } finally { lock.releaseLock(); }
  if (p.maintenant === true || (date === auj && heure <= Utilities.formatDate(new Date(Date.now() + 60000), TZ, 'HH:mm'))) pushReveiller_();   // message pour maintenant : poussé tout de suite
  return messagesListe_();
}
function messageSupprimer_(ctx, p) {
  var m = msgLire_().filter(function (x) { return x.id === Number(p.id); })[0];
  if (!m) throw new UserError('Message introuvable.');
  msgFeuille_(false).getRange(m.ligne, 9).setValue('non');
  journal_(ctx, 'message-', String(m.id));
  return messagesListe_();
}

/* ============================ Camions de livraison (m1.22, 26/11/2026) ============================
 * Décisions du propriétaire (26/11/2026) : un code par camion (3 camions pour commencer) et un camion de test qui s'ouvre sur
 * plusieurs téléphones ; le chauffeur tape seulement les gallons, le prix est fixé par le bureau ; avant de livrer, il voit les
 * 5 dernières opérations du client ; les achats du camion sont suivis aussi.
 * Onglets : « Camions » (liste), « Clients camion » (clients et prix au gallon, fixés par le bureau), « Camion (app) » (une ligne par
 * livraison, paiement reçu, chargement de propane ou autre achat). Les photos vont dans un dossier Drive privé. Rien n'est supprimé :
 * « annuler » marque la ligne.
 */
var CAM_ONGLET = 'Camions', CAM_ENTETES = ['code', 'nom', 'actif', 'note'];
var CAM_DEFAUT = [['1', 'Camion 1'], ['2', 'Camion 2'], ['3', 'Camion 3'], ['TEST', 'Camion test (formation)']];
var CAM_TEST = 'TEST', CAM_CODE_TEST = 'CMT-0001';
var CCL_ONGLET = 'Clients camion', CCL_ENTETES = ['id', 'nom', 'prix_gal', 'station', 'actif', 'modifie_le', 'modifie_par', 'note', 'groupe', 'solde_depart', 'solde_depart_le', 'solde_depart_source'];
var CCL_DEFAUT = 'DEFAUT';   // ligne spéciale : prix pour un client qui n'a pas encore son prix
var COP_ONGLET = 'Camion (app)';
var COP_ENTETES = ['id', 'genre', 'date', 'camion', 'client_id', 'client', 'gallons', 'prix_gal', 'montant', 'paye', 'credit', 'devise', 'bon_no', 'fournisseur', 'detail',
  'lat', 'lon', 'precision_m', 'pris_le', 'recu_le', 'fichier_id', 'appareil', 'cid', 'annule', 'note'];
var COP_GENRES = ['livraison', 'paiement', 'chargement', 'achat', 'photo'];   // photo = preuve (compteur ou station du client, autre) pour une réclamation
var CAM_PHOTOS = ['Compteur du client', 'Station du client', 'Autre'];
var CAM_FOURNISSEURS = ['Sodigaz', 'Haiti Propane', 'Autre'];
/* Prix et montants des livraisons : dans la monnaie des rapports de fermeture (CURRENCY), comme sur les bons du camion.
 * Un paiement ou un achat peut aussi être noté en gourdes (HTG) ou en dollars US. */
function camDevises_() { return [CURRENCY, 'HTG', 'USD'].filter(function (x, i, l) { return l.indexOf(x) === i; }); }

function camFeuille_(nom, entetes, lignes) {
  var sh = ss_().getSheetByName(nom);
  if (!sh) { sh = ss_().insertSheet(nom); sh.appendRow(entetes); sh.setFrozenRows(1); (lignes || []).forEach(function (l) { sh.appendRow(l); }); }
  return sh;
}
function camLignes_(sh, n) { return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, n).getValues(); }
function camions_() {
  var sh = camFeuille_(CAM_ONGLET, CAM_ENTETES, CAM_DEFAUT.map(function (c) { return [c[0], c[1], 'oui', 'créé le ' + maintenant_() + ' (décision du propriétaire, 26/11/2026)']; }));
  return camLignes_(sh, 3).map(function (r) { return { code: String(r[0]).trim().toUpperCase(), nom: String(r[1] || r[0]), actif: estVrai_(r[2]) }; }).filter(function (c) { return c.code; });
}
function camDe_(station) {   // « CAM:1 » → le camion, ou null
  var code = String(station || '').trim().toUpperCase().replace(/^CAM:/, '');
  return camions_().filter(function (c) { return c.code === code && c.actif; })[0] || null;
}
/* Ajoute à droite les colonnes qui manquent dans un onglet déjà créé. N'écrit que dans des cases d'en-tête vides : rien n'est écrasé. */
function camEntetes_(sh, entetes) {
  if (sh.getLastColumn() >= entetes.length) return;
  if (sh.getMaxColumns() < entetes.length) sh.insertColumnsAfter(sh.getMaxColumns(), entetes.length - sh.getMaxColumns());
  var l = sh.getRange(1, 1, 1, entetes.length).getValues()[0];
  for (var i = 0; i < entetes.length; i++) if (!String(l[i] || '').trim()) sh.getRange(1, i + 1).setValue(entetes[i]);
}
/* Clients des camions. « groupe » réunit les points de livraison d'un même client (m1.26) ; « solde_depart » est ce que le client devait
 * au soir de « solde_depart_le », écrit par le bureau avec sa source : l'application ne l'invente jamais (vide = pas de solde de départ). */
function camClients_() {
  var sh = camFeuille_(CCL_ONGLET, CCL_ENTETES, [[CCL_DEFAUT, '(prix pour un client sans prix)', '', '', 'oui', maintenant_(), '', 'à remplir par le bureau']]);
  camEntetes_(sh, CCL_ENTETES);
  return camLignes_(sh, CCL_ENTETES.length).map(function (r, i) {
    var dep = r[9] === '' || r[9] === null || r[9] === undefined || String(r[9]).trim() === '' ? null : stkNum_(r[9]);
    return { ligne: i + 2, id: String(r[0]).trim(), nom: String(r[1] || '').trim(), prix: stkNum_(r[2]) || 0, station: String(r[3] || '').trim().toUpperCase(), actif: estVrai_(r[4]),
      groupe: String(r[8] || '').trim(), depart: dep, depart_le: dep === null ? '' : phJour_(r[10]), depart_source: dep === null ? '' : String(r[11] || '').trim() };
  }).filter(function (c) { return c.id; });
}
/* ============================ Soldes des clients des camions (m1.26, 27/11/2026) ============================
 * Demande du propriétaire (27/11/2026) : savoir ce que chaque client doit, en réunissant les stations d'un même client.
 * Solde = solde de départ (bureau) + crédits des livraisons − paiements reçus, dans la monnaie des rapports.
 * Les gourdes sont converties (5 gourdes = 1 dollar haïtien) ; les paiements dans une autre monnaie sont listés à part, jamais convertis.
 * Opérations annulées et camion de test : jamais comptés. Une livraison sans prix n'a pas de montant : elle est signalée, pas estimée. */
function camCle_(o) { return o.client_id || ('N:' + String(o.client || '').trim().toLowerCase()); }
function camArr_(x) { return Math.round(x * 100) / 100; }
function camEffet_(o) {   // effet d'une opération sur le solde, dans la monnaie des rapports ; null = à part (sans prix, autre monnaie)
  if (o.genre === 'livraison') return o.montant === '' ? null : (Number(o.credit) || 0);
  if (o.devise === CURRENCY) return -(Number(o.montant) || 0);
  if (o.devise === 'HTG' && CURRENCY === 'dollars haïtiens') return -camArr_((Number(o.montant) || 0) / 5);
  return null;
}
function camCompte_(o, c) {   // l'opération entre-t-elle dans le solde de ce client ?
  if (o.annule || o.camion === CAM_TEST || (o.genre !== 'livraison' && o.genre !== 'paiement')) return false;
  return !(c && c.depart !== null && c.depart_le && o.date <= c.depart_le);   // déjà comprise dans le solde de départ
}
function camSoldes_(clients, ops) {
  var par = {}, info = {};
  var fiche = function (k, nom) { return par[k] = par[k] || { cle: k, nom: nom, groupe: '', actif: true, depart: null, depart_le: '', depart_source: '', credit: 0, paye: 0, autres: {}, sans_prix: 0, dernier: '', solde: 0 }; };
  clients.forEach(function (c) {
    if (c.id === CCL_DEFAUT) return;
    info[c.id] = c; var f = fiche(c.id, c.nom); f.groupe = c.groupe; f.actif = c.actif; f.depart = c.depart; f.depart_le = c.depart_le; f.depart_source = c.depart_source;
  });
  ops.forEach(function (o) {
    var k = camCle_(o);
    if (!camCompte_(o, info[k])) return;
    var f = fiche(k, o.client), e = camEffet_(o);
    if (o.date > f.dernier) f.dernier = o.date;
    if (e === null) { if (o.genre === 'livraison') f.sans_prix++; else f.autres[o.devise] = camArr_((f.autres[o.devise] || 0) + (Number(o.montant) || 0)); }
    else if (e >= 0) f.credit += e; else f.paye += -e;
  });
  Object.keys(par).forEach(function (k) { var f = par[k]; f.credit = camArr_(f.credit); f.paye = camArr_(f.paye); f.solde = camArr_((f.depart || 0) + f.credit - f.paye); });
  return par;
}
/** Les soldes réunis par client (groupe), le plus gros d'abord. Un client sans groupe est seul dans le sien. */
function camGroupes_(par) {
  var g = {};
  Object.keys(par).forEach(function (k) {
    var f = par[k], cle = f.groupe ? 'G:' + f.groupe.toUpperCase() : k;
    var x = g[cle] = g[cle] || { cle: cle, nom: f.groupe || f.nom, groupe: f.groupe, solde: 0, autres: {}, sans_prix: 0, dernier: '', membres: [] };
    x.solde = camArr_(x.solde + f.solde); x.sans_prix += f.sans_prix; if (f.dernier > x.dernier) x.dernier = f.dernier;
    Object.keys(f.autres).forEach(function (d) { x.autres[d] = camArr_((x.autres[d] || 0) + f.autres[d]); });
    x.membres.push({ cle: f.cle, nom: f.nom, actif: f.actif, solde: f.solde, depart: f.depart, depart_le: f.depart_le, depart_source: f.depart_source, credit: f.credit, paye: f.paye, sans_prix: f.sans_prix, dernier: f.dernier });
  });
  return Object.keys(g).map(function (k) { g[k].membres.sort(function (a, b) { return a.nom.toLowerCase() < b.nom.toLowerCase() ? -1 : 1; }); return g[k]; })
    .sort(function (a, b) { return b.solde - a.solde || (a.nom.toLowerCase() < b.nom.toLowerCase() ? -1 : 1); });
}
/** Direction / comptabilité : relevé d'un client (ou de tout un groupe) — chaque livraison et chaque paiement, avec le solde après chacun. */
function camReleve_(p) {
  var clients = camClients_(), ops = camOps_(8000), par = camSoldes_(clients, ops), cle = String(p.cle || '');
  var gr = camGroupes_(par).filter(function (x) { return x.cle === cle; })[0];
  if (!gr) throw new UserError('Client introuvable.');
  var dans = {}, info = {}; gr.membres.forEach(function (m) { dans[m.cle] = m; }); clients.forEach(function (c) { info[c.id] = c; });
  var depart = 0; gr.membres.forEach(function (m) { depart += m.depart || 0; });
  var solde = camArr_(depart), lignes = ops.filter(function (o) { var k = camCle_(o); return dans[k] && camCompte_(o, info[k]); })
    .sort(function (a, b) { return a.pris_le < b.pris_le ? -1 : 1; }).map(function (o) {
      var e = camEffet_(o); if (e !== null) solde = camArr_(solde + e);
      return { id: o.id, date: o.date, heure: o.pris_le.slice(11, 16), camion: o.camion, client: o.client, genre: o.genre, gallons: o.gallons, prix: o.prix, montant: o.montant, paye: o.paye, credit: o.credit,
        devise: o.devise, bon_no: o.bon_no, effet: e, solde: solde, fichier: o.fichier_id };
    });
  return { cle: gr.cle, nom: gr.nom, groupe: gr.groupe, devise: CURRENCY, depart: camArr_(depart), membres: gr.membres, lignes: lignes, solde: gr.solde, autres: gr.autres, sans_prix: gr.sans_prix, calcule_le: maintenant_() };
}
function camOps_(max) {
  var sh = camFeuille_(COP_ONGLET, COP_ENTETES), n = sh.getLastRow();
  if (n < 2) return [];
  var from = Math.max(2, n - (max || 8000) + 1), vide = function (v) { return v === '' || v === null ? '' : Number(v); };
  return sh.getRange(from, 1, n - from + 1, COP_ENTETES.length).getValues().map(function (r, i) {
    return { ligne: from + i, id: String(r[0]), genre: String(r[1]), date: phJour_(r[2]), camion: String(r[3]).trim().toUpperCase(), client_id: String(r[4] || ''), client: String(r[5] || ''),
      gallons: vide(r[6]), prix: vide(r[7]), montant: vide(r[8]), paye: vide(r[9]), credit: vide(r[10]), devise: String(r[11] || 'HTG'), bon_no: String(r[12] || ''),
      fournisseur: String(r[13] || ''), detail: String(r[14] || ''), lat: vide(r[15]), lon: vide(r[16]), prec: vide(r[17]), pris_le: phTexte_(r[18]), recu_le: phTexte_(r[19]),
      fichier_id: String(r[20] || ''), cid: String(r[22] || ''), annule: estVrai_(r[23]), note: String(r[24] || '') };
  });
}
function camPourTel_(o) {   // ce que le téléphone du camion reçoit d'une opération
  return { cid: o.cid, genre: o.genre, date: o.date, heure: o.pris_le.slice(11, 16), camion: o.camion, client_id: o.client_id, client: o.client, gallons: o.gallons, prix: o.prix,
    montant: o.montant, paye: o.paye, credit: o.credit, devise: o.devise, bon_no: o.bon_no, fournisseur: o.fournisseur, detail: o.detail };
}
/** Écran de départ du téléphone d'un camion : clients et prix, 5 dernières opérations de chaque client, opérations du jour. */
function camInit_(app) {
  var cam = app.camion, auj = msgJour_(0), test = cam.code === CAM_TEST;
  var clients = camClients_(), defaut = clients.filter(function (c) { return c.id === CCL_DEFAUT; })[0];
  var ops = camOps_(8000).filter(function (o) { return !o.annule && (test || o.camion !== CAM_TEST); });
  var hist = {};
  ops.forEach(function (o) {
    if (o.genre !== 'livraison' && o.genre !== 'paiement') return;
    var k = o.client_id || ('N:' + o.client.toLowerCase());
    (hist[k] = hist[k] || []).push(camPourTel_(o));
  });
  Object.keys(hist).forEach(function (k) { hist[k] = hist[k].slice(-5).reverse(); });
  var par = camSoldes_(clients, ops.filter(function (o) { return o.camion !== CAM_TEST; })), grp = {};
  camGroupes_(par).forEach(function (g) { if (g.groupe) g.membres.forEach(function (m) { grp[m.cle] = { nom: g.nom, solde: g.solde, nb: g.membres.length }; }); });
  return { version: MOBILE_VERSION, mode: 'camion', devise: CURRENCY, camion: { code: cam.code, nom: cam.nom, test: test }, date: auj,
    clients: clients.filter(function (c) { return c.actif && c.id !== CCL_DEFAUT; }).map(function (c) {
      var x = { id: c.id, nom: c.nom, prix: c.prix, solde: par[c.id] ? par[c.id].solde : 0 }, g = grp[c.id];
      if (g && g.nb > 1) { x.groupe = g.nom; x.solde_groupe = g.solde; x.nb_groupe = g.nb; }
      return x;
    })
      .sort(function (a, b) { return a.nom.toLowerCase() < b.nom.toLowerCase() ? -1 : 1; }),
    prix_defaut: defaut ? defaut.prix : 0, historique: hist, fournisseurs: CAM_FOURNISSEURS, types_photo: CAM_PHOTOS, devises: camDevises_(),
    jour: ops.filter(function (o) { return o.camion === cam.code && o.date === auj; }).map(camPourTel_), rapports: camRapportsTel_(cam.code, ops) };
}
/** Une opération envoyée par le téléphone d'un camion. Le prix vient toujours du bureau ; le chauffeur ne tape que les gallons. */
function camEnvoyer_(app, p) {
  if (app.role !== 'camion' || !app.camion) throw new UserError('Réservé au téléphone d\'un camion.', 'FORBIDDEN');
  var cam = app.camion, genre = String(p.genre || '');
  if (COP_GENRES.indexOf(genre) < 0) throw new UserError('Opération inconnue.');
  var cid = String(p.cid || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 60);
  if (!cid) throw new UserError('Opération sans identifiant.');
  var photo = String(p.photo || ''), m = photo.match(/^data:image\/jpeg;base64,(.+)$/), octets = null;
  if (photo && !m) throw new UserError('Photo invalide.');
  if (m) { octets = Utilities.base64Decode(m[1]); if (octets.length > 3000000) throw new UserError('Photo trop lourde.'); }
  if (!octets && genre !== 'paiement') throw new UserError('Il faut la photo du bon ou du reçu.');
  var recu = new Date(), auj = Utilities.formatDate(recu, TZ, 'yyyy-MM-dd'), pris = new Date(Number(p.pris_ms) || 0), note = String(p.note || '').slice(0, 200);
  if (!(pris.getTime() > 0) || pris.getTime() > recu.getTime() + 10 * 60000) { note = (note ? note + ' ; ' : '') + 'Heure du téléphone refusée : heure de réception utilisée'; pris = recu; }
  else if (recu.getTime() - pris.getTime() > 31 * 864e5) throw new UserError('Opération trop ancienne (plus de 31 jours).');
  var date = Utilities.formatDate(pris, TZ, 'yyyy-MM-dd');
  var nombre = function (v, quoi, oblig) {
    var s = String(v === undefined || v === null ? '' : v).trim(); if (!s && !oblig) return '';
    var x = phNombre_(s); if (!x || !(Number(x) > 0)) throw new UserError(quoi + ' illisible. Écrivez seulement les chiffres.');
    return Number(x);
  };
  var devise = camDevises_().indexOf(String(p.devise)) >= 0 ? String(p.devise) : CURRENCY;
  var l = { client_id: '', client: '', gallons: '', prix: '', montant: '', paye: '', credit: '', devise: CURRENCY, bon_no: String(p.bon_no || '').trim().slice(0, 40), fournisseur: '', detail: '' };
  if (genre === 'livraison' || genre === 'paiement') {
    var clients = camClients_(), c = clients.filter(function (x) { return x.id === String(p.client_id || '') && x.id !== CCL_DEFAUT; })[0];
    l.client = c ? c.nom : String(p.client || '').trim().slice(0, 80);
    if (!l.client) throw new UserError('Choisissez le client.');
    l.client_id = c ? c.id : '';
    if (genre === 'livraison') {
      l.gallons = nombre(p.gallons, 'Nombre de gallons', true);
      if (l.gallons > 20000) throw new UserError('Plus de 20 000 gallons : vérifiez le chiffre.');
      var defaut = clients.filter(function (x) { return x.id === CCL_DEFAUT; })[0];
      var prix = c && c.prix > 0 ? c.prix : (defaut && defaut.prix > 0 ? defaut.prix : 0);
      var mode = String(p.mode || '');
      if (['comptant', 'credit', 'partie'].indexOf(mode) < 0) throw new UserError('Dites si le client a payé, ou si c\'est à crédit.');
      if (prix > 0) {
        l.prix = prix; l.montant = Math.round(l.gallons * prix * 100) / 100;
        l.paye = mode === 'comptant' ? l.montant : mode === 'credit' ? 0 : nombre(p.paye, 'Montant reçu', true);
        if (l.paye > l.montant) throw new UserError('Le montant reçu dépasse le montant de la livraison (' + l.montant + ' ' + CURRENCY + ').');
        l.credit = Math.round((l.montant - l.paye) * 100) / 100;
      } else {   // pas de prix au bureau : on ne l'invente pas, le montant reste vide
        if (mode === 'partie') l.paye = nombre(p.paye, 'Montant reçu', true); else if (mode === 'credit') l.paye = 0;
        note = (note ? note + ' ; ' : '') + 'Prix pas encore fixé par le bureau (' + (mode === 'comptant' ? 'payé comptant' : mode === 'credit' ? 'à crédit' : 'une partie payée') + ')';
      }
    } else {
      l.montant = nombre(p.montant, 'Montant reçu', true); l.paye = l.montant; l.devise = devise;
    }
  } else if (genre === 'photo') {
    var cp = camClients_().filter(function (x) { return x.id === String(p.client_id || '') && x.id !== CCL_DEFAUT; })[0];
    l.client = cp ? cp.nom : String(p.client || '').trim().slice(0, 80); l.client_id = cp ? cp.id : '';
    l.detail = CAM_PHOTOS.indexOf(String(p.detail)) >= 0 ? String(p.detail) : 'Autre';
  } else if (genre === 'chargement') {
    l.fournisseur = String(p.fournisseur || '').trim().slice(0, 60);
    if (!l.fournisseur) throw new UserError('Choisissez le fournisseur.');
    l.gallons = nombre(p.gallons, 'Nombre de gallons', true);
    if (l.gallons > 20000) throw new UserError('Plus de 20 000 gallons : vérifiez le chiffre.');
  } else {
    l.detail = String(p.detail || '').trim().slice(0, 120);
    if (!l.detail) throw new UserError('Écrivez ce qui a été acheté.');
    l.montant = nombre(p.montant, 'Montant', true); l.devise = devise;
  }
  var gps = function (v, max) { var x = Number(v); return v === '' || v === undefined || v === null || isNaN(x) || Math.abs(x) > max ? '' : Math.round(x * 1e6) / 1e6; };
  var lat = gps(p.lat, 90), lon = gps(p.lon, 180), prec = lat === '' || lon === '' ? '' : Math.round(Number(p.precision) || 0);
  if (lat === '' || lon === '') throw new UserError('Position GPS manquante. Allumez la localisation du téléphone, puis refaites l\'opération.', 'GPS');   // propriétaire, 26/11/2026 : toujours la position
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var sh = camFeuille_(COP_ONGLET, COP_ENTETES), n = sh.getLastRow(), col = COP_ENTETES.indexOf('cid') + 1;
    if (n > 1) {
      var from = Math.max(2, n - 2999), cids = sh.getRange(from, col, n - from + 1, 1).getValues();
      for (var i = 0; i < cids.length; i++) if (String(cids[i][0]) === cid) return { recu: true, deja: true, cid: cid };
    }
    var fid = '';
    if (octets) {
      var nomF = 'CAM' + cam.code + '_' + genre + '_' + date + '_' + Utilities.formatDate(pris, TZ, 'HHmmss') + '.jpg';
      var f = phDossier_({ code: 'CAM' + cam.code, nom: cam.nom }, date, 'CAM', 'RRD – Camions (app, privé)').createFile(Utilities.newBlob(octets, 'image/jpeg', nomF));
      f.setDescription('Camion (application RRD) — ' + cam.nom + ' — ' + genre + ' — ' + (l.client || l.fournisseur || l.detail) + ' — ' + cid);
      fid = f.getId();
    }
    var id = 'C' + Utilities.formatDate(recu, TZ, 'yyMMdd') + '-' + n;
    sh.appendRow([id, genre, date, cam.code, l.client_id, l.client, l.gallons, l.prix, l.montant, l.paye, l.credit, l.devise, l.bon_no, l.fournisseur, l.detail, lat, lon, prec,
      Utilities.formatDate(pris, TZ, 'yyyy-MM-dd HH:mm:ss'), Utilities.formatDate(recu, TZ, 'yyyy-MM-dd HH:mm:ss'), fid, app.id, cid, '', note]);
    return { recu: true, cid: cid, id: id, prix: l.prix, montant: l.montant, paye: l.paye, credit: l.credit };
  } finally { lock.releaseLock(); }
}
/** Direction / comptabilité : opérations des camions entre deux dates, totaux, clients et prix. */
function camListe_(p) {
  var auj = msgJour_(0), ok = function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d || '')); };
  var au = ok(p.au) ? String(p.au) : auj, du = ok(p.du) ? String(p.du) : au, filtre = String(p.camion || '').trim().toUpperCase();
  var cams = camions_(), noms = {}; cams.forEach(function (c) { noms[c.code] = c.nom; });
  var toutes = camOps_(8000), lesClients = camClients_(), soldes = camGroupes_(camSoldes_(lesClients, toutes));
  var ops = toutes.filter(function (o) { return o.date >= du && o.date <= au && (filtre ? o.camion === filtre : o.camion !== CAM_TEST); })
    .sort(function (a, b) { return a.pris_le < b.pris_le ? 1 : -1; });
  var t = { livre: 0, charge: 0, montant: 0, paye: 0, credit: 0, encaisse: 0, sans_prix: 0, achats: {} };
  ops.forEach(function (o) {
    if (o.annule) return;
    if (o.genre === 'livraison') { t.livre += Number(o.gallons) || 0; if (o.montant === '') t.sans_prix++; else { t.montant += o.montant; t.paye += Number(o.paye) || 0; t.credit += Number(o.credit) || 0; } }
    else if (o.genre === 'chargement') t.charge += Number(o.gallons) || 0;
    else if (o.genre === 'paiement') { if (o.devise === CURRENCY) t.encaisse += Number(o.montant) || 0; else t.achats['Paiements reçus en ' + o.devise] = (t.achats['Paiements reçus en ' + o.devise] || 0) + (Number(o.montant) || 0); }
    else if (o.genre === 'achat') t.achats[o.devise] = (t.achats[o.devise] || 0) + (Number(o.montant) || 0);
  });
  return { du: du, au: au, aujourdhui: auj, camion: filtre, camions: cams, totaux: t, devise: CURRENCY, rapports: camRapportsVue_(du, au, filtre), stations: stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return { code: String(s.code), nom: String(s.nom) }; }),
    clients: lesClients.map(function (c) { return { id: c.id, nom: c.nom, prix: c.prix, station: c.station, actif: c.actif, groupe: c.groupe, depart: c.depart, depart_le: c.depart_le, depart_source: c.depart_source }; }),
    soldes: soldes,
    ops: ops.map(function (o) { var x = camPourTel_(o); x.id = o.id; x.camion_nom = noms[o.camion] || o.camion; x.fichier = o.fichier_id; x.lat = o.lat; x.lon = o.lon; x.prec = o.prec; x.annule = o.annule; x.note = o.note; return x; }) };
}
/** Le bureau crée un client ou change son prix (aussi le prix par défaut, id = DEFAUT). Le changement ne touche pas les livraisons déjà reçues. */
function camClient_(ctx, p) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var clients = camClients_(), sh = ss_().getSheetByName(CCL_ONGLET), id = String(p.id || '').trim();
    var prix = p.prix === undefined || p.prix === null || String(p.prix).trim() === '' ? null : Number(phNombre_(String(p.prix)));
    if (prix !== null && !(prix > 0 && prix < 100000)) throw new UserError('Prix au gallon invalide.');
    var station = String(p.station || '').trim().toUpperCase();
    if (station && !stations_().some(function (s) { return String(s.code) === station; })) throw new UserError('Station inconnue.');
    if (id) {
      var c = clients.filter(function (x) { return x.id === id; })[0];
      if (!c) throw new UserError('Client introuvable.');
      var nom = id === CCL_DEFAUT ? c.nom : (String(p.nom || '').trim().slice(0, 80) || c.nom);
      sh.getRange(c.ligne, 2, 1, 6).setValues([[nom, prix === null ? (c.prix || '') : prix, p.station === undefined ? c.station : station, p.actif === undefined ? (c.actif ? 'oui' : 'non') : (p.actif ? 'oui' : 'non'), maintenant_(), ctx.nom]]);
      journal_(ctx, 'camion-client', id + ' ' + nom + ' prix=' + (prix === null ? c.prix : prix) + ' (avant : ' + c.prix + ')');
      if (p.groupe !== undefined && id !== CCL_DEFAUT) {
        var groupe = String(p.groupe || '').trim().replace(/\s+/g, ' ').slice(0, 40);
        var meme = clients.filter(function (x) { return x.groupe && x.groupe.toUpperCase() === groupe.toUpperCase(); })[0];
        if (meme) groupe = meme.groupe;   // même écriture que le groupe qui existe déjà
        sh.getRange(c.ligne, 9).setValue(groupe);
        journal_(ctx, 'camion-groupe', id + ' ' + nom + ' groupe=« ' + groupe + ' » (avant : « ' + c.groupe + ' »)');
      }
      if (p.depart !== undefined && id !== CCL_DEFAUT) {
        var avant = c.depart === null ? 'aucun' : c.depart + ' au ' + c.depart_le + ' (' + c.depart_source + ')';
        if (String(p.depart === null ? '' : p.depart).trim() === '') {
          sh.getRange(c.ligne, 10, 1, 3).setValues([['', '', '']]);
          journal_(ctx, 'camion-solde-depart', id + ' ' + nom + ' retiré (avant : ' + avant + ')');
        } else {
          var dep = Number(phNombre_(String(p.depart).trim().replace(/^-/, ''))) * (/^-/.test(String(p.depart).trim()) ? -1 : 1);
          if (isNaN(dep) || Math.abs(dep) > 100000000 || phNombre_(String(p.depart).trim().replace(/^-/, '')) === '') throw new UserError('Solde de départ illisible. Écrivez seulement les chiffres.');
          var le = String(p.depart_le || ''), src = String(p.depart_source || '').trim().slice(0, 200);
          if (!/^\d{4}-\d{2}-\d{2}$/.test(le) || le > msgJour_(0)) throw new UserError('Donnez la date du solde de départ (pas dans le futur).');
          if (src.length < 3) throw new UserError('Écrivez d\'où vient ce solde (cahier, feuille, relevé…).');
          sh.getRange(c.ligne, 10, 1, 3).setValues([[dep, le, src + ' — ' + ctx.nom + ', ' + maintenant_()]]);
          journal_(ctx, 'camion-solde-depart', id + ' ' + nom + ' ' + dep + ' au ' + le + ' — ' + src + ' (avant : ' + avant + ')');
        }
      }
    } else {
      var nomN = String(p.nom || '').trim().slice(0, 80);
      if (!nomN) throw new UserError('Écrivez le nom du client.');
      if (clients.some(function (x) { return x.nom.toLowerCase() === nomN.toLowerCase(); })) throw new UserError('Ce client existe déjà.');
      var num = 1; clients.forEach(function (x) { var k = Number(x.id.replace(/^K/, '')); if (k >= num) num = k + 1; });
      var grN = String(p.groupe || '').trim().replace(/\s+/g, ' ').slice(0, 40), memeN = clients.filter(function (x) { return x.groupe && x.groupe.toUpperCase() === grN.toUpperCase(); })[0];
      sh.appendRow(['K' + num, nomN, prix === null ? '' : prix, station, 'oui', maintenant_(), ctx.nom, 'créé par le bureau', memeN ? memeN.groupe : grN]);
      journal_(ctx, 'camion-client+', 'K' + num + ' ' + nomN + ' prix=' + (prix === null ? '' : prix));
    }
  } finally { lock.releaseLock(); }
  return camListe_(p.vue || {});
}
function camAnnuler_(ctx, p) {
  var motif = String(p.motif || '').trim().slice(0, 150);
  if (!motif) throw new UserError('Écrivez la raison.');
  var o = camOps_(8000).filter(function (x) { return x.id === String(p.id || ''); })[0];
  if (!o) throw new UserError('Opération introuvable.');
  if (o.annule) throw new UserError('Déjà annulée.');
  var sh = ss_().getSheetByName(COP_ONGLET);
  sh.getRange(o.ligne, COP_ENTETES.indexOf('annule') + 1, 1, 2).setValues([['oui', (o.note ? o.note + ' ; ' : '') + 'Annulée le ' + maintenant_() + ' par ' + ctx.nom + ' : ' + motif]]);
  journal_(ctx, 'camion-annule', o.id + ' ' + o.genre + ' ' + (o.client || o.fournisseur || o.detail) + ' : ' + motif);
  return camListe_(p.vue || {});
}

/* ============================ Rapport du jour des camions (m1.23, 26/11/2026) ============================
 * Demande du propriétaire (26/11/2026) : un rapport par camion et par jour. Le téléphone du camion montre le résumé de la journée
 * (calculé avec les opérations déjà reçues) ; le chauffeur écrit seulement les employés du camion et l'argent remis. Le bureau garde
 * une photo des totaux au moment du rapport dans « Rapports camion (app) ». Un seul rapport par camion et par jour ; « annuler »
 * (direction) marque la ligne et le chauffeur peut le refaire.
 */
var CRP_ONGLET = 'Rapports camion (app)';
var CRP_ENTETES = ['id', 'date', 'camion', 'employes', 'charge_gal', 'livre_gal', 'nb_livraisons', 'montant_livraisons', 'recu_livraisons', 'credit', 'paiements_recus', 'achats',
  'autres_monnaies', 'argent_attendu', 'argent_remis', 'difference', 'sans_prix', 'note', 'fait_le', 'recu_le', 'fichier_id', 'appareil', 'cid', 'annule',
  'depose', 'depose_pour', 'depose_par', 'remis_bureau', 'remis_bureau_a', 'en_main', 'en_main_par'];
/* Où est l'argent de la journée (m1.27, propriétaire, 27/11/2026) : déposé à la banque, remis au bureau, ou gardé en main — chaque montant avec
 * le nom de la personne qui a l'argent. « argent_remis » reste le total des trois. Un rapport d'avant m1.27 n'a que le total : le détail reste vide. */
/** Totaux d'un camion pour un jour, dans la monnaie des rapports. Les gourdes sont converties (5 gourdes = 1 dollar haïtien) ; les dollars US sont listés à part. */
function camTotauxJour_(ops, camion, date) {
  var t = { charge: 0, livre: 0, nb: 0, montant: 0, recu: 0, credit: 0, paiements: 0, achats: 0, autres: {}, sans_prix: 0 };
  var conv = function (o, signe) {   // montant dans la monnaie des rapports, ou null (à lister à part)
    if (o.devise === CURRENCY) return Number(o.montant) || 0;
    if (o.devise === 'HTG' && CURRENCY === 'dollars haïtiens') return Math.round((Number(o.montant) || 0) / 5 * 100) / 100;
    var k = (signe > 0 ? 'Paiements reçus en ' : 'Achats en ') + o.devise; t.autres[k] = (t.autres[k] || 0) + (Number(o.montant) || 0); return null;
  };
  ops.forEach(function (o) {
    if (o.annule || o.camion !== camion || o.date !== date) return;
    if (o.genre === 'chargement') t.charge += Number(o.gallons) || 0;
    else if (o.genre === 'livraison') { t.nb++; t.livre += Number(o.gallons) || 0; if (o.montant === '') t.sans_prix++; else { t.montant += o.montant; t.credit += Number(o.credit) || 0; } t.recu += Number(o.paye) || 0; }
    else if (o.genre === 'paiement') { var p = conv(o, 1); if (p !== null) t.paiements += p; }
    else if (o.genre === 'achat') { var a = conv(o, -1); if (a !== null) t.achats += a; }
  });
  ['charge', 'livre', 'montant', 'recu', 'credit', 'paiements', 'achats'].forEach(function (k) { t[k] = Math.round(t[k] * 100) / 100; });
  t.attendu = Math.round((t.recu + t.paiements - t.achats) * 100) / 100;
  return t;
}
function camRapports_(max) {
  var sh = camFeuille_(CRP_ONGLET, CRP_ENTETES), n = sh.getLastRow();
  camEntetes_(sh, CRP_ENTETES);
  if (n < 2) return [];
  var from = Math.max(2, n - (max || 3000) + 1), nbr = function (v) { return v === '' || v === null ? '' : Number(v); };
  return sh.getRange(from, 1, n - from + 1, CRP_ENTETES.length).getValues().map(function (r, i) {
    return { ligne: from + i, id: String(r[0]), date: phJour_(r[1]), camion: String(r[2]).trim().toUpperCase(), employes: String(r[3] || ''), charge: nbr(r[4]), livre: nbr(r[5]), nb: nbr(r[6]), montant: nbr(r[7]),
      recu: nbr(r[8]), credit: nbr(r[9]), paiements: nbr(r[10]), achats: nbr(r[11]), autres: String(r[12] || ''), attendu: nbr(r[13]), remis: nbr(r[14]), difference: nbr(r[15]), sans_prix: nbr(r[16]),
      note: String(r[17] || ''), fait_le: phTexte_(r[18]), fichier_id: String(r[20] || ''), cid: String(r[22] || ''), annule: estVrai_(r[23]),
      depose: nbr(r[24]), depose_pour: String(r[25] || ''), depose_par: String(r[26] || ''), bureau: nbr(r[27]), bureau_qui: String(r[28] || ''), main: nbr(r[29]), main_qui: String(r[30] || '') };
  });
}
/** Pour le téléphone du camion : rapports d'aujourd'hui et des 7 jours d'avant (faits ou non), pour envoyer ceux qui manquent (propriétaire, 26/11/2026). */
function camRapportsTel_(code, ops) {
  var raps = camRapports_(3000).filter(function (r) { return r.camion === code && !r.annule; });
  return [0, -1, -2, -3, -4, -5, -6, -7].map(function (n) { return msgJour_(n); }).map(function (d) {
    var r = raps.filter(function (x) { return x.date === d; })[0], t = camTotauxJour_(ops, code, d);
    return { date: d, fait: !!r, employes: r ? r.employes : '', remis: r ? r.remis : '', fait_le: r ? r.fait_le : '', ops: t.nb + (t.charge > 0 ? 1 : 0) + (t.paiements > 0 ? 1 : 0) + (t.achats > 0 ? 1 : 0), totaux: t };
  });
}
function camRapport_(app, p) {
  if (app.role !== 'camion' || !app.camion) throw new UserError('Réservé au téléphone d\'un camion.', 'FORBIDDEN');
  var cam = app.camion, cid = String(p.cid || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 60);
  if (!cid) throw new UserError('Rapport sans identifiant.');
  var date = String(p.date || ''), auj = msgJour_(0);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date > auj || date < msgJour_(-7)) throw new UserError('Date du rapport invalide.');
  var employes = String(p.employes || '').trim().slice(0, 150);
  if (!employes) throw new UserError('Écrivez les employés du camion.');
  var detail = ['depose', 'bureau', 'main'].some(function (k) { return p[k] !== undefined && p[k] !== null; }), remis, ou = { depose: '', depose_pour: '', depose_par: '', bureau: '', bureau_qui: '', main: '', main_qui: '' };
  if (detail) {   // téléphone m1.27 : trois endroits possibles, chacun avec le nom de la personne qui a l'argent
    var mt = function (k, quoi) { var v = String(p[k] === undefined || p[k] === null ? '' : p[k]).trim(); if (v === '' || v === '0') return 0; var x = phNombre_(v); if (x === '') throw new UserError(quoi + ' : montant illisible. Écrivez seulement les chiffres.'); return Number(x); };
    var nom = function (k) { return String(p[k] || '').trim().replace(/\s+/g, ' ').slice(0, 60); };
    ou.depose = mt('depose', 'Dépôt à la banque'); ou.bureau = mt('bureau', 'Argent remis au bureau'); ou.main = mt('main', 'Argent gardé en main');
    if (ou.depose > 0) {
      ou.depose_pour = BD_BENEF.indexOf(String(p.depose_pour)) >= 0 ? String(p.depose_pour) : 'Autre'; ou.depose_par = nom('depose_par');
      if (!ou.depose_par) throw new UserError('Écrivez le nom de la personne qui a fait le dépôt.');
      if (!String(p.photo || '')) throw new UserError('Il faut la photo du bordereau de dépôt.');
    }
    if (ou.bureau > 0) { ou.bureau_qui = nom('bureau_qui'); if (!ou.bureau_qui) throw new UserError('Écrivez le nom de la personne du bureau qui a reçu l\'argent.'); }
    if (ou.main > 0) { ou.main_qui = nom('main_qui'); if (!ou.main_qui) throw new UserError('Écrivez le nom de la personne qui garde l\'argent.'); }
    remis = Math.round((ou.depose + ou.bureau + ou.main) * 100) / 100;
  } else {
    var rs = String(p.remis === undefined || p.remis === null ? '' : p.remis).trim(); remis = rs === '0' ? '0' : phNombre_(rs);
    if (remis === '') throw new UserError('Argent remis illisible. Écrivez seulement les chiffres (0 si rien).');
    remis = Number(remis);
  }
  var photo = String(p.photo || ''), m = photo.match(/^data:image\/jpeg;base64,(.+)$/), octets = null;
  if (photo && !m) throw new UserError('Photo invalide.');
  if (m) { octets = Utilities.base64Decode(m[1]); if (octets.length > 3000000) throw new UserError('Photo trop lourde.'); }
  var recu = new Date(), fait = new Date(Number(p.pris_ms) || 0);
  if (!(fait.getTime() > 0) || fait.getTime() > recu.getTime() + 10 * 60000) fait = recu;
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var deja = camRapports_(3000).filter(function (r) { return r.camion === cam.code && r.date === date && !r.annule; })[0];
    if (deja) { if (deja.cid === cid) return { recu: true, deja: true, cid: cid }; throw new UserError('Le rapport du ' + date.split('-').reverse().join('/') + ' est déjà fait pour ce camion. Pour le corriger, prévenez la direction.', 'DEJA'); }
    var t = camTotauxJour_(camOps_(8000), cam.code, date), fid = '';
    if (octets) {
      var f = phDossier_({ code: 'CAM' + cam.code, nom: cam.nom }, date, 'CAM', 'RRD – Camions (app, privé)').createFile(Utilities.newBlob(octets, 'image/jpeg', 'CAM' + cam.code + '_rapport_' + date + '.jpg'));
      f.setDescription('Rapport du jour (application RRD) — ' + cam.nom + ' — ' + date + ' — ' + cid); fid = f.getId();
    }
    var sh = camFeuille_(CRP_ONGLET, CRP_ENTETES), id = 'RC' + date.replace(/-/g, '').slice(2) + '-' + cam.code + '-' + sh.getLastRow();
    camEntetes_(sh, CRP_ENTETES);
    sh.appendRow([id, date, cam.code, employes, t.charge, t.livre, t.nb, t.montant, t.recu, t.credit, t.paiements, t.achats, Object.keys(t.autres).map(function (k) { return k + ' : ' + t.autres[k]; }).join(' ; '),
      t.attendu, remis, Math.round((t.attendu - remis) * 100) / 100, t.sans_prix, String(p.note || '').slice(0, 300), Utilities.formatDate(fait, TZ, 'yyyy-MM-dd HH:mm:ss'), Utilities.formatDate(recu, TZ, 'yyyy-MM-dd HH:mm:ss'), fid, app.id, cid, '',
      ou.depose, ou.depose_pour, ou.depose_par, ou.bureau, ou.bureau_qui, ou.main, ou.main_qui]);
    if (detail) try {
      argEntrees_('CAM:' + cam.code, cam.nom, date, id, [{ lieu: 'depot', montant: ou.depose, tenu_par: ou.depose_par, pour: ou.depose_pour }, { lieu: 'bureau', montant: ou.bureau, tenu_par: ou.bureau_qui }, { lieu: 'main', montant: ou.main, tenu_par: ou.main_qui }], app.id, fid);
    } catch (e) { console.error(e && e.stack || e); }   // le rapport est déjà enregistré : le registre d'argent ne doit pas le faire échouer
    return { recu: true, cid: cid, id: id };
  } finally { lock.releaseLock(); }
}
function camRapportAnnuler_(ctx, p) {
  var motif = String(p.motif || '').trim().slice(0, 150);
  if (!motif) throw new UserError('Écrivez la raison.');
  var r = camRapports_(3000).filter(function (x) { return x.id === String(p.id || ''); })[0];
  if (!r) throw new UserError('Rapport introuvable.');
  if (r.annule) throw new UserError('Déjà annulé.');
  var sh = ss_().getSheetByName(CRP_ONGLET);
  sh.getRange(r.ligne, CRP_ENTETES.indexOf('annule') + 1).setValue('oui');
  sh.getRange(r.ligne, CRP_ENTETES.indexOf('note') + 1).setValue((r.note ? r.note + ' ; ' : '') + 'Annulé le ' + maintenant_() + ' par ' + ctx.nom + ' : ' + motif);
  journal_(ctx, 'camion-rapport-annule', r.id + ' : ' + motif);
  return camListe_(p.vue || {});
}
/** Rapports du jour pour le tableau de bord et le résumé du soir : un par camion et par jour où il y a eu une opération ou un rapport. */
function camRapportsVue_(du, au, filtre) {
  var ops = camOps_(8000), raps = camRapports_(3000).filter(function (r) { return !r.annule; }), noms = {}, cles = {}, out = [];
  camions_().forEach(function (c) { noms[c.code] = c.nom; });
  var garder = function (cam, d) { return d >= du && d <= au && (filtre ? cam === filtre : cam !== CAM_TEST); };
  ops.forEach(function (o) { if (!o.annule && garder(o.camion, o.date)) cles[o.date + '|' + o.camion] = 1; });
  raps.forEach(function (r) { if (garder(r.camion, r.date)) cles[r.date + '|' + r.camion] = 1; });
  Object.keys(cles).sort().reverse().forEach(function (k) {
    var d = k.split('|')[0], cam = k.split('|')[1], r = raps.filter(function (x) { return x.date === d && x.camion === cam; })[0], t = camTotauxJour_(ops, cam, d);
    out.push({ date: d, camion: cam, camion_nom: noms[cam] || cam, totaux: t, fait: !!r, id: r ? r.id : '', employes: r ? r.employes : '', remis: r ? r.remis : '', fait_le: r ? r.fait_le : '', note: r ? r.note : '', fichier: r ? r.fichier_id : '',
      difference: r ? Math.round((t.attendu - r.remis) * 100) / 100 : '', change: !!r && (r.attendu !== t.attendu || r.livre !== t.livre),
      ou: r && (r.depose !== '' || r.bureau !== '' || r.main !== '') ? { depose: r.depose || 0, depose_pour: r.depose_pour, depose_par: r.depose_par, bureau: r.bureau || 0, bureau_qui: r.bureau_qui, main: r.main || 0, main_qui: r.main_qui } : null });
  });
  return out;
}

/* ============================ Livraisons suggérées aux chauffeurs (m1.30 ; m1.32 : sans tournée, 27/11/2026) ============================
 * Demande du propriétaire (27/11/2026) : suggérer automatiquement aux chauffeurs les clients à livrer. Ce ne sont que des suggestions :
 * le bureau ne donne pas d'arrêts, le chauffeur n'est pas obligé de les suivre. Le bureau voit la même liste (onglet « À livrer »).
 * - L'historique vient du classeur « VRAC 2025 » (lecture seule, autorisée par le propriétaire le 27/11/2026) et des livraisons tapées
 *   dans l'application. Une même journée de livraison écrite aux deux endroits n'est comptée qu'une fois.
 * - Un nom écrit dans le classeur n'est relié à un client que s'il est écrit pareil, ou si l'onglet « Noms VRAC 2025 » le dit (avec sa
 *   source). Tout autre nom reste « à relier » : l'application ne devine jamais.
 * - Proposition = dernière livraison + rythme habituel (médiane des écarts entre les 9 derniers jours de livraison). Il faut au moins
 *   3 jours de livraison. C'est une estimation d'après le passé, jamais une commande : rien n'est créé ni envoyé tout seul.
 * - Nos stations : le stock estimé (onglet Stock) sert aussi, quand il est récent et sans alerte.
 * - Le téléphone du camion demande la liste à part (action camSuggestions), pour que l'ouverture reste rapide. */
var VRAC_ID = '1QVbTKGJvmli42yVcrurSZ22nazxweBBYvjLOalmYsMs';
var VRAC_JOURS = 200;        // on lit les mois qui touchent les 200 derniers jours
var VNM_ONGLET = 'Noms VRAC 2025', VNM_ENTETES = ['nom_ecrit', 'client_id', 'source', 'ajoute_le', 'ajoute_par'];
var VNM_PROP = 'liste du propriétaire (noms confirmés pour le rappel d\'encaissement, 11/2026)', VNM_NOTE = 'note du client dans « Clients camion » (26/11/2026)';
var VNM_DEFAUT = [   // seulement des noms déjà confirmés ; le reste est montré « à relier »
  ['Red Rebel Brochette', 'K1', VNM_PROP], ['Red Rebel 32', 'K2', VNM_PROP], ['RRDDED', 'K3', VNM_NOTE], ['Red Rebel 83', 'K5', VNM_PROP],
  ['Jules Daves', 'K8', VNM_NOTE], ['Jules 38', 'K9', VNM_NOTE], ['Esmatech 61', 'K11', VNM_PROP],
  ['Esmatech 65', 'K12', VNM_PROP], ['Esmatech 63', 'K12', VNM_PROP], ['Esmatech Diquini', 'K12', VNM_PROP],
  ['Esmatech 75', 'K13', VNM_PROP], ['Esmatech 79', 'K13', VNM_PROP], ['Esmatech Mahotière 75', 'K13', 'même nom, autre orthographe de Mahotier'],
  ['Esmatech 93', 'K14', VNM_PROP], ['Esmatech 91', 'K14', VNM_PROP], ['Esmatech 53', 'K38', VNM_NOTE],
  ['Supremo', 'K16', VNM_NOTE], ['Louco', 'K17', VNM_PROP], ['Louco Sapoti', 'K17', VNM_PROP], ['Louko Sapoti', 'K17', VNM_PROP],
  ['Jeanty', 'K19', VNM_NOTE], ['Jeanty Propane', 'K19', VNM_PROP], ['Black', 'K20', VNM_NOTE], ['Tipa Tipa', 'K21', VNM_NOTE],
  ['Ti pa ti pa', 'K21', 'même nom, écrit avec des espaces'], ['Saloto', 'K39', VNM_NOTE]];
var SUG_MIN = 3;             // jours de livraison qu'il faut pour parler d'un rythme
var SUG_AVANT = 2;           // « bientôt » : prévu dans 2 jours ou moins
var SUG_ARRET = 3;           // plus de 3 fois le rythme (et plus de 30 jours) sans livraison : le client ne commande plus
var SUG_STOCK = 2;           // nos stations : à livrer quand il reste 2 jours de gaz ou moins (estimation)
var SUG_HEURE_DEMAIN = 18;   // le chauffeur voit les suggestions du lendemain à partir de 18 h, ou dès son rapport du jour

function vracNorm_(s) {   // « Suprême  » = « supreme », « C. Jean » = « c jean »
  var t = String(s || '').toLowerCase().replace(/[̀-ͯ]/g, ''), de = 'àâäáãéèêëíìîïóòôöõúùûüçñ', vers = 'aaaaaeeeeiiiiooooouuuucn', o = '';
  for (var i = 0; i < t.length; i++) { var k = de.indexOf(t.charAt(i)); o += k < 0 ? t.charAt(i) : vers.charAt(k); }
  return o.replace(/[^a-z0-9]+/g, ' ').replace(/^ +| +$/g, '');
}
function vracNb_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var n = Number(String(v === null || v === undefined ? '' : v).replace(/[\s  ]/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}
function vracMois_(texte) {   // « Pour le mois de novembre2026 » → « 2026-11 » ; '' si illisible
  var m = /(janvier|fevrier|mars|avril|mai|juin|juillet|aout|septembre|octobre|novembre|decembre) ?(\d{4})/.exec(vracNorm_(texte));
  if (!m) return '';
  var k = ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre'].indexOf(m[1]) + 1;
  return m[2] + '-' + (k < 10 ? '0' : '') + k;
}
function sugMs_(d) { return Date.UTC(Number(d.slice(0, 4)), Number(d.slice(5, 7)) - 1, Number(d.slice(8, 10))); }
function sugJours_(a, b) { return Math.round((sugMs_(b) - sugMs_(a)) / 864e5); }   // nombre de jours de a à b
function sugPlus_(d, n) { return new Date(sugMs_(d) + n * 864e5).toISOString().slice(0, 10); }
function sugMediane_(l) { var t = l.slice().sort(function (a, b) { return a - b; }), n = t.length; return n % 2 ? t[(n - 1) / 2] : (t[n / 2 - 1] + t[n / 2]) / 2; }

/** Livraisons écrites dans le classeur VRAC 2025 : { noms: { nom normalisé: { ecrit, j: { date: gallons du jour } } }, onglets, lu_le, erreur } (gardé 3 h). */
function vracLire_(force) {
  var cache = CacheService.getScriptCache(), c = force ? null : cache.get('vrac_hist_v1');
  if (c) return JSON.parse(c);
  var out = { lu_le: maintenant_(), onglets: [], noms: {}, erreur: '' };
  try {
    var auj = msgJour_(0), limite = msgJour_(-VRAC_JOURS).slice(0, 7);
    SpreadsheetApp.openById(VRAC_ID).getSheets().forEach(function (sh) {
      if (sh.isSheetHidden()) return;
      var n = sh.getLastRow(); if (n < 8) return;
      var m = vracMois_(sh.getRange(3, 1).getDisplayValue());
      if (!m || m < limite || m > auj.slice(0, 7)) return;
      var fin = Number(sugPlus_(sugPlus_(m + '-28', 4).slice(0, 8) + '01', -1).slice(8, 10));   // dernier jour du mois
      var jour = 0, nb = 0, dernier = '';
      sh.getRange(7, 1, n - 6, 6).getValues().forEach(function (r) {
        var j = r[0] instanceof Date ? r[0].getDate() : vracNb_(r[0]);
        if (j >= 1 && j <= 31 && j === Math.floor(j)) jour = j;
        var cli = String(r[2] || '').replace(/\s+/g, ' ').replace(/^ | $/g, ''), gal = vracNb_(r[5]) || (vracNb_(r[3]) + vracNb_(r[4])), k = vracNorm_(cli);
        if (!k || !(gal > 0) || !jour || jour > fin || /^total/.test(k)) return;
        var d = m + '-' + (jour < 10 ? '0' : '') + jour;
        if (d > auj) return;
        var f = out.noms[k] = out.noms[k] || { ecrit: cli, j: {} };
        f.j[d] = Math.round(((f.j[d] || 0) + gal) * 10) / 10;
        nb++; if (d > dernier) dernier = d;
      });
      out.onglets.push({ nom: String(sh.getName()).replace(/^ +| +$/g, ''), mois: m, livraisons: nb, dernier: dernier });
    });
    out.onglets.sort(function (a, b) { return a.mois < b.mois ? -1 : a.mois > b.mois ? 1 : 0; });
  } catch (e) { out.erreur = String((e && e.message) || e); }
  if (!out.erreur) try { cache.put('vrac_hist_v1', JSON.stringify(out), 10800); } catch (e2) { }
  return out;
}
function vracFeuille_() {
  return camFeuille_(VNM_ONGLET, VNM_ENTETES, VNM_DEFAUT.map(function (a) { return [a[0], a[1], a[2], maintenant_(), 'Claude (demande du propriétaire, 27/11/2026)']; }));
}
function vracAlias_() {   // { nom normalisé: id du client }
  var o = {};
  camLignes_(vracFeuille_(), 2).forEach(function (r) { var k = vracNorm_(r[0]), id = String(r[1] || '').replace(/\s/g, ''); if (k && id) o[k] = id; });
  return o;
}
/** Jours de livraison de chaque client (classeur + application) et noms du classeur qui ne sont reliés à aucun client. */
function sugHistorique_(clients, ops, force) {
  var vr = vracLire_(force), alias = vracAlias_(), parNom = {}, existe = {}, par = {}, aRelier = [];
  clients.forEach(function (c) { if (c.id !== CCL_DEFAUT) { existe[c.id] = 1; parNom[vracNorm_(c.nom)] = c.id; } });
  Object.keys(vr.noms).forEach(function (k) {
    var f = vr.noms[k], id = existe[alias[k]] ? alias[k] : parNom[k], ds = Object.keys(f.j).sort();
    if (!id) { var g = 0; ds.forEach(function (d) { g += f.j[d]; }); aRelier.push({ ecrit: f.ecrit, n: ds.length, gallons: Math.round(g), dernier: ds[ds.length - 1] }); return; }
    var p = par[id] = par[id] || {};
    ds.forEach(function (d) { p[d] = Math.round(((p[d] || 0) + f.j[d]) * 10) / 10; });
  });
  var app = {};
  ops.forEach(function (o) {
    if (o.annule || o.genre !== 'livraison' || o.camion === CAM_TEST || !existe[o.client_id] || !(Number(o.gallons) > 0)) return;
    var a = app[o.client_id] = app[o.client_id] || {}; a[o.date] = Math.round(((a[o.date] || 0) + Number(o.gallons)) * 10) / 10;
  });
  Object.keys(app).forEach(function (id) { var p = par[id] = par[id] || {}; Object.keys(app[id]).forEach(function (d) { p[d] = Math.max(p[d] || 0, app[id][d]); }); });
  aRelier.sort(function (a, b) { return a.dernier < b.dernier ? 1 : a.dernier > b.dernier ? -1 : b.n - a.n; });
  return { par: par, a_relier: aRelier, source: { lu_le: vr.lu_le, onglets: vr.onglets, erreur: vr.erreur } };
}
/** Le rythme d'un client d'après ses jours de livraison { date: gallons }. statut : peu | arret | retard | bientot | plus_tard. */
function sugCalcul_(jours, auj) {
  var ds = Object.keys(jours || {}).sort(), rec = ds.slice(-9), n = rec.length, dernier = n ? rec[n - 1] : '';
  if (n < SUG_MIN) return { n: ds.length, dernier: dernier, statut: 'peu' };
  var ecarts = [], i; for (i = 1; i < n; i++) ecarts.push(sugJours_(rec[i - 1], rec[i]));
  var rythme = Math.max(1, sugMediane_(ecarts)), gal = sugMediane_(rec.map(function (d) { return jours[d]; }));
  var prochaine = sugPlus_(dernier, Math.round(rythme)), retard = sugJours_(prochaine, auj), depuis = sugJours_(dernier, auj);
  return { n: ds.length, utilises: n, du: rec[0], dernier: dernier, depuis: depuis, rythme: Math.round(rythme * 10) / 10, gallons: Math.round(gal), prochaine: prochaine, retard: retard,
    statut: depuis > Math.max(30, SUG_ARRET * rythme) ? 'arret' : retard >= 0 ? 'retard' : retard >= -SUG_AVANT ? 'bientot' : 'plus_tard' };
}
/** Stock estimé de nos stations, réduit à ce qu'il faut pour les suggestions (gardé 1 h : le calcul complet est lent). */
function sugStock_(force) {
  var cache = CacheService.getScriptCache(), c = force ? null : cache.get('sug_stock_v1');
  if (c) return JSON.parse(c);
  var o = {};
  try { stockEstime_().stations.forEach(function (s) { o[s.code] = { stock: s.stock, pct: s.pct, vente_jour: s.vente_jour || 0, dernier: s.dernier || '', alerte: s.alerte || '' }; }); } catch (e) { return o; }
  try { cache.put('sug_stock_v1', JSON.stringify(o), 3600); } catch (e2) { }
  return o;
}
/** Ce que lit le calcul (une seule fois, même pour deux jours). */
function sugDonnees_(force) {
  var clients = camClients_(), ops = camOps_(8000);
  return { clients: clients, h: sugHistorique_(clients, ops, force), soldes: camSoldes_(clients, ops), stock: sugStock_(force) };
}
/** Le calcul commun au bureau et aux chauffeurs : clients à livrer le jour « cible » (aujourd'hui par défaut) ou dans les 2 jours qui suivent,
 *  avec ce que le calcul a utilisé. Pour demain, le rythme et le stock sont comptés jusqu'à demain. */
function sugCalculer_(force, cible, donnees) {
  var auj = cible || msgJour_(0), dn = donnees || sugDonnees_(force), clients = dn.clients, h = dn.h, soldes = dn.soldes, stock = dn.stock;
  var liste = [], peu = [], arret = [], plusTard = 0, total = 0;
  clients.forEach(function (c) {
    if (c.id === CCL_DEFAUT || !c.actif) return;
    var s = sugCalcul_(h.par[c.id], auj), f = soldes[c.id];
    total += s.n;
    s.id = c.id; s.nom = c.nom; s.groupe = c.groupe; s.station = c.station; s.motif = 'rythme';
    s.solde = f && (f.depart !== null || f.credit || f.paye) ? f.solde : null;
    var st = c.station ? stock[c.station] : null;
    if (st) {
      if (st.alerte) s.stock = { fiable: false, raison: 'le calcul du stock a une alerte (onglet Stock)' };
      else if (!(st.vente_jour > 0)) s.stock = { fiable: false, raison: 'pas de ventes récentes dans les rapports' };
      else if (!st.dernier || st.dernier < msgJour_(-2)) s.stock = { fiable: false, raison: 'dernier rapport de la station : ' + (st.dernier ? st.dernier.split('-').reverse().join('/') : 'aucun') };
      else {   // jours de gaz qui restent le jour cible : stock au dernier rapport ÷ ventes moyennes, moins les jours écoulés depuis ce rapport
        var jr = Math.max(0, Math.floor(st.stock / st.vente_jour) - Math.max(0, sugJours_(st.dernier, auj)));
        s.stock = { fiable: true, gallons: st.stock, pct: st.pct, vente_jour: st.vente_jour, jours: jr, au: st.dernier, cible: auj };
        if (jr <= SUG_STOCK) { s.statut = 'retard'; s.motif = 'stock'; }
        else if (jr <= SUG_STOCK + 2 && s.statut !== 'retard') { s.statut = 'bientot'; s.motif = 'stock'; }
      }
    }
    if (s.statut === 'retard' || s.statut === 'bientot') liste.push(s);
    else if (s.statut === 'peu') peu.push({ id: c.id, nom: c.nom, n: s.n, dernier: s.dernier });
    else if (s.statut === 'arret') arret.push({ id: c.id, nom: c.nom, dernier: s.dernier, depuis: s.depuis, rythme: s.rythme });
    else plusTard++;
  });
  var rang = function (s) { return s.statut === 'retard' ? 0 : 1; };
  liste.sort(function (a, b) { return rang(a) - rang(b) || (b.retard === undefined ? -999 : b.retard) - (a.retard === undefined ? -999 : a.retard) || (a.nom.toLowerCase() < b.nom.toLowerCase() ? -1 : 1); });
  return { aujourdhui: auj, calcule_le: maintenant_(), liste: liste, plus_tard: plusTard, peu: peu, arret: arret, a_relier: h.a_relier, source: h.source, livraisons: total, clients: clients };
}
/** Tableau de bord, onglet « À livrer ». */
function sugListe_(p) {
  var r = sugCalculer_(!!(p && p.actualiser));
  r.clients = r.clients.filter(function (c) { return c.actif && c.id !== CCL_DEFAUT; }).map(function (c) { return { id: c.id, nom: c.nom }; })
    .sort(function (a, b) { return a.nom.toLowerCase() < b.nom.toLowerCase() ? -1 : 1; });
  r.devise = CURRENCY; r.reglages = { min: SUG_MIN, avant: SUG_AVANT, stock_jours: SUG_STOCK };
  return r;
}
/** Téléphone d'un camion : la même liste, sans les montants, comme suggestion ; pour aujourd'hui et pour demain.
 *  Le téléphone montre celle de demain à partir de 18 h (heure d'Haïti) ou dès que le rapport du jour est fait (propriétaire, 27/11/2026). */
function sugPourCamion_(app) {
  if (app.role !== 'camion' || !app.camion) throw new UserError('Réservé au téléphone d\'un camion.', 'FORBIDDEN');
  var dn = sugDonnees_(false), auj = msgJour_(0), dem = msgJour_(1);
  var court = function (r) {
    return { date: r.aujourdhui, liste: r.liste.map(function (s) {
      return { id: s.id, nom: s.nom, statut: s.statut, motif: s.motif, retard: s.retard, prochaine: s.prochaine, dernier: s.dernier, depuis: s.depuis, rythme: s.rythme, gallons: s.gallons,
        stock_jours: s.motif === 'stock' ? s.stock.jours : null };
    }) };
  };
  return { calcule_le: maintenant_(), lu_le: dn.h.source.lu_le, heure_demain: SUG_HEURE_DEMAIN, jours: [court(sugCalculer_(false, auj, dn)), court(sugCalculer_(false, dem, dn))] };
}
/** Le bureau relie un nom écrit dans le classeur VRAC 2025 à un client de l'application. */
function vracRelier_(ctx, p) {
  var k = vracNorm_(p.ecrit), id = String(p.client_id || '');
  var c = camClients_().filter(function (x) { return x.id === id && x.id !== CCL_DEFAUT; })[0], vr = vracLire_();
  if (!k || !vr.noms[k]) throw new UserError('Ce nom n\'est pas dans le classeur VRAC 2025. Actualisez la page.');
  if (!c) throw new UserError('Choisissez le client.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (vracAlias_()[k]) throw new UserError('Ce nom est déjà relié. Pour le changer, corrigez l\'onglet « ' + VNM_ONGLET + ' ».');
    vracFeuille_().appendRow([vr.noms[k].ecrit, c.id, 'Bureau, tableau de bord : ' + ctx.nom, maintenant_(), ctx.nom]);
    journal_(ctx, 'vrac-nom-relie', vr.noms[k].ecrit + ' → ' + c.id + ' ' + c.nom);
  } finally { lock.releaseLock(); }
  return sugListe_({});
}

/* ============================ Chat en direct (m1.33, 27/11/2026) ============================
 * Demande du propriétaire (27/11/2026) : un chat en direct dans l'application entre stations, camions, direction et comptabilité.
 * Les messages passent par la base Firebase du projet (« Realtime Database »), pas par ce serveur. Chaque téléphone a une identité
 * anonyme Firebase ; la base n'accepte que les identités inscrites dans sa liste « membres », et cette liste n'est écrite que par le
 * relais « RRD Notifications » (avec le compte Google du propriétaire, sans clé secrète). Ici, le serveur vérifie le code du téléphone
 * et note qui est qui dans l'onglet « Chat (membres) » ; le relais recopie cet onglet dans la base et retire les téléphones désactivés.
 * Qui parle sous quel nom : une station = la station ; un camion = le camion ; direction et comptabilité = la personne du code. */
var CHM_ONGLET = 'Chat (membres)';
var CHM_ENTETES = ['uid', 'code', 'nom', 'role', 'appareil', 'demande_le', 'actif', 'base', 'base_le'];
function chatMoi_(app) {
  if (app.role === 'transferts') throw new UserError('Le chat n\'est pas encore ouvert aux agents de transfert.', 'FORBIDDEN');
  if (app.role === 'camion') return { code: 'CAM-' + app.camion.code, nom: app.camion.nom, role: 'camion' };
  var act = String(app.ligne.code_activation || '').toUpperCase().replace(/[^A-Z0-9-]/g, '');
  if (app.role === 'management') return { code: act, nom: String(app.ligne.nom || '').trim() || 'Direction', role: 'direction' };
  if (app.role === 'accounting') return { code: act, nom: String(app.ligne.nom || '').trim() || 'Comptabilité', role: 'compta' };
  return { code: String(app.station.code).toUpperCase().replace(/[^A-Z0-9-]/g, ''), nom: String(app.station.nom), role: 'station' };
}
/** Le téléphone donne son identité Firebase ; on l'inscrit (ou on la garde) et on réveille le relais. */
function chatEntrer_(app, p) {
  var uid = String(p.uid || '').trim(), moi = chatMoi_(app), act = String(app.ligne.code_activation || '');
  if (!/^[A-Za-z0-9]{20,40}$/.test(uid)) throw new UserError('Identité du chat invalide.');
  if (!moi.code) throw new UserError('Code du téléphone illisible.');
  var lock = LockService.getScriptLock(), change = false; lock.waitLock(20000);
  try {
    var sh = camFeuille_(CHM_ONGLET, CHM_ENTETES), l = camLignes_(sh, CHM_ENTETES.length), deja = 0;
    l.forEach(function (r, i) {
      var u = String(r[0]), memeTel = String(r[4]) === act, actif = estVrai_(r[6]);
      if (u === uid) {
        deja = i + 2;
        if (!actif || String(r[1]) !== moi.code || String(r[2]) !== moi.nom || String(r[3]) !== moi.role || !memeTel) {
          sh.getRange(i + 2, 2, 1, 6).setValues([[moi.code, moi.nom, moi.role, act, maintenant_(), 'oui']]); change = true;
        }
      } else if (memeTel && actif) { sh.getRange(i + 2, 7).setValue('non'); change = true; }   // ancienne identité du même téléphone (application réinstallée)
    });
    if (!deja) { sh.appendRow([uid, moi.code, moi.nom, moi.role, act, maintenant_(), 'oui', '', '']); change = true; }
  } finally { lock.releaseLock(); }
  if (change) chatReveiller_();
  var r = camLignes_(camFeuille_(CHM_ONGLET, CHM_ENTETES), CHM_ENTETES.length).filter(function (x) { return String(x[0]) === uid; })[0];
  return { code: moi.code, nom: moi.nom, role: moi.role, inscrit: !!r && String(r[7]) === 'present' };
}
function chatReveiller_() {
  try {
    var url = prop_('PUSH_RELAIS_URL');
    if (url) UrlFetchApp.fetch(url, { method: 'post', payload: JSON.stringify({ action: 'chatMembres' }), contentType: 'text/plain', muteHttpExceptions: true });
  } catch (e) { }
}

/* Notifications du chat (m1.34, propriétaire 27/11/2026) : après un message, le téléphone le signale ici ; on note une ligne dans
 * « Chat (alertes) » (qui, à qui, quelle conversation — jamais le texte du message) et on réveille le relais, qui pousse la notification
 * aux téléphones visés, même application fermée. Message privé : l'autre personne. Général : tout le monde, seulement quand la
 * direction écrit (pour ne pas faire sonner tous les téléphones à chaque message). */
var CHA_ONGLET = 'Chat (alertes)', CHA_ENTETES = ['id', 'le', 'de', 'de_nom', 'cible', 'conv', 'appareil', 'pousse_le', 'pousse_a'];
function chatSignal_(app, p) {
  var moi = chatMoi_(app), conv = String(p.conv || ''), cible;
  if (conv === 'general') { if (moi.role !== 'direction') return { alerte: false }; cible = 'TOUS'; }
  else {
    var parts = conv.split('__');
    if (parts.length !== 2 || parts.indexOf(moi.code) < 0 || parts[0] === parts[1] || !/^[A-Z0-9-]{2,30}$/.test(parts[0] + parts[1])) throw new UserError('Conversation inconnue.');
    cible = parts[0] === moi.code ? parts[1] : parts[0];
  }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = camFeuille_(CHA_ONGLET, CHA_ENTETES), n = sh.getLastRow(), deja = false;
    if (n >= 2) sh.getRange(Math.max(2, n - 49), 1, Math.min(50, n - 1), CHA_ENTETES.length).getValues().forEach(function (r) {
      if (String(r[2]) === moi.code && String(r[5]) === conv && !String(r[7] || '').trim()) deja = true;   // une alerte pas encore partie suffit
    });
    if (!deja) sh.appendRow(['CA-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 90 + 10), maintenant_(), moi.code, moi.nom, cible, conv, String(app.ligne.code_activation || ''), '', '']);
  } finally { lock.releaseLock(); }
  try {
    var url = prop_('PUSH_RELAIS_URL');
    if (url) UrlFetchApp.fetch(url, { method: 'post', payload: JSON.stringify({ action: 'chatAlertes' }), contentType: 'text/plain', muteHttpExceptions: true });
  } catch (e) { }
  return { alerte: true };
}

/* ============================ Demandes « ouvrir une station » (site redrebel.us, m1.36, 27/11/2026) ============================
 * Demande de la direction (27/11/2026) : au lieu d'appeler, les personnes qui veulent ouvrir une station de remplissage remplissent
 * un formulaire sur le site. Chaque demande devient une ligne de l'onglet « Demandes station (site) » ; la direction les voit dans
 * l'onglet « 📝 Demandes » du tableau de bord. Le formulaire est public : textes coupés, pas de formule possible dans la feuille,
 * piège à robots, et au plus 30 demandes par heure en tout. */
var DMS_ONGLET = 'Demandes station (site)';
var DMS_ENTETES = ['id', 'recu_le', 'nom', 'telephone', 'email', 'ville', 'terrain', 'delai', 'message', 'langue', 'statut', 'suivi_par', 'suivi_le', 'note'];
var DMS_TERRAIN = { oui: 'oui, terrain ou local trouvé', cherche: 'je cherche encore', non: 'pas encore' };
var DMS_DELAI = { vite: 'dans les 3 mois', annee: 'dans l\'année', plus: 'plus tard / je me renseigne' };
var DMS_STATUTS = ['nouvelle', 'contactée', 'classée'];
function dmsTexte_(v, max) {
  var s = String(v === undefined || v === null ? '' : v).replace(/[\u0000-\u001f\u007f]+/g, ' ').replace(/\s+/g, ' ').replace(/^ | $/g, '').slice(0, max);
  return /^[=+\-@]/.test(s) ? '’' + s : s;   // jamais de formule dans la feuille
}
function demandeStation_(p) {
  if (String(p.site || '').trim()) return { recu: true };   // champ caché rempli : robot, on ne garde rien
  var nom = dmsTexte_(p.nom, 80), tel = dmsTexte_(p.telephone, 30), mail = dmsTexte_(p.email, 80), ville = dmsTexte_(p.ville, 80), msg = dmsTexte_(p.message, 1000);
  var terrain = DMS_TERRAIN[String(p.terrain || '')] || '', delai = DMS_DELAI[String(p.delai || '')] || '', langue = ['fr', 'ht', 'en'].indexOf(String(p.langue)) >= 0 ? String(p.langue) : 'fr';
  if (nom.length < 2) throw new UserError('nom', 'INVALIDE');
  if (String(tel).replace(/[^0-9]/g, '').length < 8) throw new UserError('telephone', 'INVALIDE');
  if (ville.length < 2) throw new UserError('ville', 'INVALIDE');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(mail.replace(/^’/, ''))) throw new UserError('email', 'INVALIDE');   // obligatoire (direction, 27/11/2026)
  if (msg.length < 10) throw new UserError('message', 'INVALIDE');
  var cache = CacheService.getScriptCache(), cle = 'dms_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHH'), n = Number(cache.get(cle) || 0);
  if (n >= 30) throw new UserError('trop', 'TROP');
  cache.put(cle, String(n + 1), 3700);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = camFeuille_(DMS_ONGLET, DMS_ENTETES), der = sh.getLastRow(), deja = false;
    if (der >= 2) sh.getRange(Math.max(2, der - 19), 1, Math.min(20, der - 1), 6).getValues().forEach(function (r) {   // même personne, même ville, à l'instant : double envoi
      var le = r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd HH:mm:ss') : String(r[1]);
      if (String(r[3]) === tel && String(r[5]) === ville && le.slice(0, 13) === maintenant_().slice(0, 13)) deja = true;
    });
    if (!deja) sh.appendRow(['DS-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 90 + 10), maintenant_(), nom, tel, mail, ville, terrain, delai, msg, langue, 'nouvelle', '', '', '']);
  } finally { lock.releaseLock(); }
  return { recu: true };
}
function demandesListe_() {
  var sh = camFeuille_(DMS_ONGLET, DMS_ENTETES);
  var l = camLignes_(sh, DMS_ENTETES.length).map(function (r, i) {
    var o = { ligne: i + 2 }; DMS_ENTETES.forEach(function (k, j) { o[k] = r[j] instanceof Date ? Utilities.formatDate(r[j], TZ, 'yyyy-MM-dd HH:mm:ss') : String(r[j] === null ? '' : r[j]); }); return o;
  }).filter(function (o) { return o.id; }).reverse();
  return { demandes: l, statuts: DMS_STATUTS, nouvelles: l.filter(function (o) { return o.statut === 'nouvelle'; }).length };
}
function demandeStatut_(ctx, p) {
  var st = String(p.statut || ''), note = dmsTexte_(p.note, 300);
  if (DMS_STATUTS.indexOf(st) < 0) throw new UserError('Statut inconnu.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = camFeuille_(DMS_ONGLET, DMS_ENTETES), d = demandesListe_().demandes.filter(function (x) { return x.id === String(p.id || ''); })[0];
    if (!d) throw new UserError('Demande introuvable. Actualisez la page.');
    sh.getRange(d.ligne, 11, 1, 4).setValues([[st, ctx.nom, maintenant_(), note ? (d.note ? d.note + ' ; ' : '') + note : d.note]]);
    journal_(ctx, 'demande-station', d.id + ' → ' + st);
  } finally { lock.releaseLock(); }
  return demandesListe_();
}


/* ============================ Inventaire des photos Telegram dans Drive (27/11/2026, lecture seule) ============================
 * Demande du propriétaire : faire de la place dans Drive (87 %) en retirant les photos Telegram déjà utilisées et rangées ailleurs.
 * Cette fonction NE SUPPRIME RIEN : elle compte, pèse et repère les doublons (même nom et même taille), puis écrit le résultat
 * dans l'onglet « Inventaire Drive (Telegram) » de la feuille privée. À lancer depuis l'éditeur ; elle reprend où elle s'était arrêtée. */
var INV_RACINES = { 'Telegram Desktop': '15aaDHT_86Fh7zFQL16E2nTJ6MdYbepFb', '11 Photos Telegram 2026': '1fAW8U-no5_BqEk-YOvm8b3pw_pK7tMQX',
  '09 Photos Telegram 2023': '1i_iqLHLu89Iwja99iLjEgC1pVDRqrG-N', '10 Exports WhatsApp & Telegram': '1pU6LF5CKXw6In9GtZ6ExGStRdmeOBZPs' };
function inventaireTelegram() {   // lecture seule ; s'arrête avant 5 minutes et reprend où il en était (même dans un très gros dossier)
  var debut = Date.now(), props = PropertiesService.getScriptProperties(), etat = JSON.parse(props.getProperty('INV_ETAT') || 'null');
  if (!etat) { etat = { pile: [], cur: null, fichiers: 0 }; Object.keys(INV_RACINES).forEach(function (n) { etat.pile.push({ id: INV_RACINES[n], chemin: n, racine: n }); }); }
  var sh = camFeuille_('Inventaire Drive (Telegram)', ['racine', 'dossier', 'fichier', 'id', 'type', 'taille_ko', 'cree_le', 'proprietaire']);
  var lot = [], temps = function () { return Date.now() - debut < 240000; };
  var vider = function () { if (lot.length) { sh.getRange(sh.getLastRow() + 1, 1, lot.length, 8).setValues(lot); etat.fichiers += lot.length; lot = []; } };
  while ((etat.cur || etat.pile.length) && temps()) {
    if (!etat.cur) {
      var d = etat.pile.pop(), dos = DriveApp.getFolderById(d.id), sous = dos.getFolders(), prop = '';
      while (sous.hasNext()) { var s = sous.next(); etat.pile.push({ id: s.getId(), chemin: d.chemin + ' / ' + s.getName(), racine: d.racine }); }
      try { prop = dos.getOwner() ? dos.getOwner().getEmail() : ''; } catch (e) { }
      etat.cur = { id: d.id, chemin: d.chemin, racine: d.racine, prop: prop, jeton: '' };
    }
    var c = etat.cur, it = c.jeton ? DriveApp.continueFileIterator(c.jeton) : DriveApp.getFolderById(c.id).getFiles();
    while (it.hasNext() && temps()) {
      var x = it.next();
      lot.push([c.racine, c.chemin, x.getName(), x.getId(), x.getMimeType(), Math.round(x.getSize() / 1024), Utilities.formatDate(x.getDateCreated(), TZ, 'yyyy-MM-dd'), c.prop]);
      if (lot.length >= 500) vider();
    }
    if (it.hasNext()) c.jeton = it.getContinuationToken(); else etat.cur = null;
  }
  vider();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'inventaireTelegram') ScriptApp.deleteTrigger(t); });
  var msg;
  if (etat.cur || etat.pile.length) {
    props.setProperty('INV_ETAT', JSON.stringify(etat));
    ScriptApp.newTrigger('inventaireTelegram').timeBased().after(60 * 1000).create();   // se relance tout seul dans 1 minute (lecture seule)
    msg = 'À SUIVRE — ' + etat.fichiers + ' fichiers notés, ' + etat.pile.length + ' dossiers restants (relance automatique) — ' + maintenant_();
  } else {
    props.deleteProperty('INV_ETAT');
    msg = 'FINI — ' + etat.fichiers + ' fichiers notés. Drive utilisé : ' + Math.round(DriveApp.getStorageUsed() / 1048576) + ' Mo sur ' + Math.round(DriveApp.getStorageLimit() / 1048576) + ' Mo — ' + maintenant_();
  }
  sh.getRange(1, 10).setValue(msg); Logger.log(msg);
}

/* ============================ À l'honneur : équipe et employé(e) du mois (m1.39, 27/11/2026) ============================
 * Propriétaire, 27/11/2026 : une page du site pour que les employés se sentent appréciés.
 *  - « Équipe du mois » = la 1re station du RRD Challenge du mois (points lus dans l'Activity Log), confirmée par la direction.
 *  - « Employé(e) du mois » = vote des collègues dans l'application (tout le réseau : stations et camions), du 20 à la fin du mois ;
 *    une voix par employé et par mois (la dernière compte) ; pas pour soi-même. La direction confirme.
 *  - Sur le site : prénom + nom de la station ; la photo seulement après le « oui » donné sur le téléphone de la station (ou du camion).
 *  - Rien n'est publié avant la confirmation de la direction ; jamais de montant ni de prix. */
var VOT_ONGLET = 'Votes employé du mois', VOT_ENTETES = ['id', 'recu_le', 'mois', 'votant', 'votant_station', 'pour', 'pour_station', 'appareil', 'cid', 'note'];
var HON_ONGLET = 'À l\'honneur (site)', HON_ENTETES = ['mois', 'genre', 'station', 'station_nom', 'noms', 'pourquoi', 'photo_id', 'accord', 'accord_le', 'accord_appareil', 'confirme_par', 'confirme_le', 'retire'];
var VOT_JOUR_DEBUT = 20;   // le vote ouvre le 20 du mois et ferme le dernier jour à minuit (heure d'Haïti)

function votNorm_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase(); }
function votOuvert_() { return Number(msgJour_(0).slice(8, 10)) >= VOT_JOUR_DEBUT; }
function horsTest_(code) { return code !== STATION_FORMATION && code !== 'CAM:' + CAM_TEST; }
function votNoms_() {
  var noms = {};
  stations_().forEach(function (s) { if (estVrai_(s.actif)) noms[String(s.code)] = String(s.nom); });
  camions_().forEach(function (c) { if (c.actif) noms['CAM:' + c.code] = c.nom; });
  return noms;
}
/** Les employés des 45 derniers jours, d'après les rapports des stations et des camions (sans doublon), par station / camion. */
function votCandidats_() {
  var lim = msgJour_(-45), noms = votNoms_(), par = {}, vus = {};
  var ajout = function (code, liste) {
    if (!noms[code] || !horsTest_(code)) return;
    (liste || []).forEach(function (n) {
      n = String(n || '').replace(/\s+/g, ' ').trim();
      if (n.length < 2 || n.length > 60) return;
      var k = code + '|' + votNorm_(n); if (vus[k]) return; vus[k] = 1;
      (par[code] = par[code] || []).push(n);
    });
  };
  rapportsEffectifs_().forEach(function (r) { if (!estAnnule(r) && iso(r.date_rapport) >= lim) ajout(String(r.station), decouperNoms(r.employes_service)); });
  camRapports_(1500).forEach(function (r) { if (!r.annule && r.date >= lim) ajout('CAM:' + r.camion, decouperNoms(r.employes)); });
  return Object.keys(par).sort(function (a, b) { return noms[a] < noms[b] ? -1 : 1; })
    .map(function (k) { return { code: k, nom: noms[k], employes: par[k].sort(function (a, b) { return votNorm_(a) < votNorm_(b) ? -1 : 1; }) }; });
}
function votLire_(mois) {
  var sh = camFeuille_(VOT_ONGLET, VOT_ENTETES), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, VOT_ENTETES.length).getValues().map(function (r) {
    return { id: Number(r[0]) || 0, recu_le: phTexte_(r[1]), mois: String(r[2]).slice(0, 7), votant: String(r[3]), votant_station: String(r[4]), pour: String(r[5]), pour_station: String(r[6]), cid: String(r[8]) };
  }).filter(function (v) { return !mois || v.mois === mois; });
}
/** Les voix qui comptent : la dernière de chaque votant (même nom, même station) pour le mois. */
function votCompte_(mois) {
  var der = {};
  votLire_(mois).forEach(function (v) { der[v.votant_station + '|' + votNorm_(v.votant)] = v; });
  var t = {}, noms = votNoms_();
  Object.keys(der).forEach(function (k) {
    var v = der[k], c = v.pour_station + '|' + votNorm_(v.pour);
    (t[c] = t[c] || { nom: v.pour, station: v.pour_station, station_nom: noms[v.pour_station] || v.pour_station, voix: 0 }).voix++;
  });
  return { votants: Object.keys(der).length, liste: Object.keys(t).map(function (k) { return t[k]; }).sort(function (a, b) { return b.voix - a.voix || (votNorm_(a.nom) < votNorm_(b.nom) ? -1 : 1); }) };
}
function honLire_() {
  var sh = camFeuille_(HON_ONGLET, HON_ENTETES), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HON_ENTETES.length).getValues().map(function (r, i) {
    return { ligne: i + 2, mois: String(r[0]).slice(0, 7), genre: String(r[1]), station: String(r[2]), station_nom: String(r[3]), noms: String(r[4]), pourquoi: String(r[5]),
      photo_id: String(r[6] || ''), accord: String(r[7] || ''), accord_le: phTexte_(r[8]), confirme_par: String(r[10] || ''), confirme_le: phTexte_(r[11]), retire: estVrai_(r[12]) };
  });
}
function honCode_(app) {   // le code de la station ou du camion de ce téléphone, ou ''
  if (app.role === 'camion' && app.camion) return 'CAM:' + app.camion.code;
  return (!app.role && app.station) ? String(app.station.code) : '';
}

/** Téléphone d'une station ou d'un camion : le vote du mois et, s'il y en a, la demande d'accord pour la photo sur le site. */
function votesTel_(app) {
  var code = honCode_(app);
  if (!code) throw new UserError('Réservé au téléphone d\'une station ou d\'un camion.', 'FORBIDDEN');
  var mois = msgJour_(0).slice(0, 7), groupes = horsTest_(code) ? votCandidats_() : [], deja = {};
  votLire_(mois).forEach(function (v) { if (v.votant_station === code) deja[votNorm_(v.votant)] = v.pour; });
  var miens = (groupes.filter(function (g) { return g.code === code; })[0] || { employes: [] }).employes;
  var accords = honLire_().filter(function (h) { return !h.retire && h.station === code && h.photo_id && !h.accord; })
    .map(function (h) { return { mois: h.mois, genre: h.genre, noms: h.noms, station_nom: h.station_nom }; });
  return { mois: mois, ouvert: votOuvert_() && horsTest_(code), code: code, miens: miens, groupes: groupes,
    deja: miens.filter(function (n) { return deja[votNorm_(n)]; }).map(function (n) { return { votant: n, pour: deja[votNorm_(n)] }; }), accords: accords };
}
function vote_(app, p) {
  var code = honCode_(app);
  if (!code) throw new UserError('Réservé au téléphone d\'une station ou d\'un camion.', 'FORBIDDEN');
  if (!horsTest_(code)) throw new UserError('Téléphone de test : le vote n\'est pas ouvert.');
  if (!votOuvert_()) throw new UserError('Le vote ouvre le ' + VOT_JOUR_DEBUT + ' du mois.');
  var cid = String(p.cid || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40), groupes = votCandidats_();
  var dans = function (st, nom) { var g = groupes.filter(function (x) { return x.code === st; })[0]; return g ? g.employes.filter(function (n) { return votNorm_(n) === votNorm_(nom); })[0] || '' : ''; };
  var votant = dans(code, p.votant), pourSt = String(p.pour_station || ''), pour = dans(pourSt, p.pour);
  if (!votant) throw new UserError('Choisissez votre nom dans la liste.');
  if (!pour) throw new UserError('Choisissez un(e) collègue dans la liste.');
  if (pourSt === code && votNorm_(pour) === votNorm_(votant)) throw new UserError('On ne vote pas pour soi-même.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var l = votLire_(), mois = msgJour_(0).slice(0, 7);
    if (cid && l.some(function (v) { return v.cid === cid; })) return { recu: true };
    var id = l.reduce(function (m, v) { return Math.max(m, v.id); }, 0) + 1;
    camFeuille_(VOT_ONGLET, VOT_ENTETES).appendRow([id, maintenant_(), mois, votant, code, pour, pourSt, app.id, cid, '']);
  } finally { lock.releaseLock(); }
  return { recu: true };
}
function accordPhoto_(app, p) {
  var code = honCode_(app);
  if (!code) throw new UserError('Réservé au téléphone d\'une station ou d\'un camion.', 'FORBIDDEN');
  var h = honLire_().filter(function (x) { return !x.retire && x.station === code && x.mois === String(p.mois) && x.genre === String(p.genre) && x.photo_id; })[0];
  if (!h) throw new UserError('Cette demande n\'existe plus.');
  if (h.accord) return { accord: h.accord };
  var oui = p.oui === true, sh = camFeuille_(HON_ONGLET, HON_ENTETES);
  if (oui) { try { DriveApp.getFileById(h.photo_id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { throw new UserError('Photo introuvable. Prévenez la direction.'); } }
  sh.getRange(h.ligne, 8, 1, 3).setValues([[oui ? 'oui' : 'non', maintenant_(), app.id]]);
  try { CacheService.getScriptCache().remove('honneur_public'); } catch (e) { }
  return { accord: oui ? 'oui' : 'non' };
}

/** Bureau : points du RRD Challenge du mois (Activity Log), voix, et ce qui est déjà confirmé. */
function honChallenge_(mois) {
  var log = SpreadsheetApp.openById(DSR_ACTIVITY_LOG_ID), sh = log.getSheetByName('RRD Challenge (points)'), tz = log.getSpreadsheetTimeZone(), t = {};
  if (!sh || sh.getLastRow() < 2) return [];
  var noms = votNoms_();
  sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues().forEach(function (r) {
    var d = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0] || '').slice(0, 10), code = stkCode_(r[1]);
    if (d.slice(0, 7) !== mois || !code || !noms[code] || !horsTest_(code)) return;
    var x = t[code] = t[code] || { code: code, nom: noms[code], points: 0, jours: 0 };
    x.points += Number(r[12]) || 0; x.jours++;
  });
  return Object.keys(t).map(function (k) { return t[k]; }).sort(function (a, b) { return b.points - a.points || (a.nom < b.nom ? -1 : 1); });
}
function honneurBureau_(p) {
  var mois = /^\d{4}-\d{2}$/.test(String(p.mois || '')) ? String(p.mois) : msgJour_(0).slice(0, 7), ch = [];
  var err = '';
  try { ch = honChallenge_(mois); } catch (e) { err = 'Classement du RRD Challenge illisible : ' + (e.message || e); }
  var c = votCompte_(mois), noms = votNoms_(), equipes = {};
  rapportsEffectifs_().forEach(function (r) {   // les noms de l'équipe pendant le mois, pour préremplir « Équipe du mois »
    if (estAnnule(r) || iso(r.date_rapport).slice(0, 7) !== mois) return;
    var code = String(r.station), e = equipes[code] = equipes[code] || [];
    decouperNoms(r.employes_service).forEach(function (n) { n = String(n).trim(); if (n && !e.some(function (x) { return votNorm_(x) === votNorm_(n); })) e.push(n); });
  });
  return { mois: mois, aujourdhui: msgJour_(0), ouvert: votOuvert_(), debut: VOT_JOUR_DEBUT, challenge: ch, challenge_erreur: err, votes: c.liste, votants: c.votants,
    equipes: equipes, stations: Object.keys(noms).filter(horsTest_).map(function (k) { return { code: k, nom: noms[k] }; }),
    honneur: honLire_().filter(function (h) { return h.mois === mois; }).map(function (h) { return { genre: h.genre, station: h.station, station_nom: h.station_nom, noms: h.noms, pourquoi: h.pourquoi, photo: !!h.photo_id, accord: h.accord, confirme_par: h.confirme_par, confirme_le: h.confirme_le, retire: h.retire }; }) };
}
function honneurConfirmer_(ctx, p) {
  var mois = String(p.mois || ''), genre = p.genre === 'equipe' ? 'equipe' : p.genre === 'employe' ? 'employe' : '';
  if (!/^\d{4}-\d{2}$/.test(mois) || !genre) throw new UserError('Mois ou genre manquant.');
  var noms = votNoms_(), station = String(p.station || '');
  if (!noms[station] || !horsTest_(station)) throw new UserError('Choisissez la station.');
  var liste = String(p.noms || '').replace(/[\r\n]+/g, ', ').replace(/\s+/g, ' ').trim().slice(0, 300);
  var pourquoi = String(p.pourquoi || '').replace(/\s+/g, ' ').trim().slice(0, 240);
  if (!liste) throw new UserError(genre === 'equipe' ? 'Écrivez les prénoms de l\'équipe.' : 'Choisissez l\'employé(e).');
  if (/^[=+\-@]/.test(liste)) liste = ' ' + liste;
  if (/^[=+\-@]/.test(pourquoi)) pourquoi = ' ' + pourquoi;
  var photoId = '';
  if (p.photo) {
    var m = /^data:image\/(jpeg|png);base64,(.+)$/.exec(String(p.photo));
    if (!m) throw new UserError('Photo illisible.');
    var dossier = phDossier_({ code: 'HON', nom: 'À l\'honneur' }, mois + '-01', 'HON', 'RRD – À l\'honneur (photos du site)');
    photoId = dossier.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), 'image/' + m[1], mois + '_' + genre + '_' + station + '.' + (m[1] === 'png' ? 'png' : 'jpg'))).getId();
  }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = camFeuille_(HON_ONGLET, HON_ENTETES), ex = honLire_().filter(function (h) { return h.mois === mois && h.genre === genre && !h.retire; })[0];
    var ligne = [mois, genre, station, noms[station], liste, pourquoi, photoId, '', '', '', ctx.nom, maintenant_(), ''];
    if (ex) {
      if (!photoId && ex.station === station) { ligne[6] = ex.photo_id; ligne[7] = ex.accord; ligne[8] = ex.accord_le; }   // même station, pas de nouvelle photo : l'accord donné reste
      sh.getRange(ex.ligne, 1, 1, HON_ENTETES.length).setValues([ligne]);
    } else sh.appendRow(ligne);
    journal_(ctx, 'honneur', mois + ' ' + genre + ' ' + station);
  } finally { lock.releaseLock(); }
  try { CacheService.getScriptCache().remove('honneur_public'); } catch (e) { }
  return honneurBureau_({ mois: mois });
}
function honneurRetirer_(ctx, p) {
  var h = honLire_().filter(function (x) { return x.mois === String(p.mois) && x.genre === String(p.genre) && !x.retire; })[0];
  if (!h) throw new UserError('Rien à retirer.');
  camFeuille_(HON_ONGLET, HON_ENTETES).getRange(h.ligne, 13).setValue('oui');
  if (h.photo_id) { try { DriveApp.getFileById(h.photo_id).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) { } }
  journal_(ctx, 'honneur-', h.mois + ' ' + h.genre);
  try { CacheService.getScriptCache().remove('honneur_public'); } catch (e) { }
  return honneurBureau_({ mois: h.mois });
}
/** Public (site redrebel.us) : la dernière équipe et le (la) dernier(ère) employé(e) du mois confirmés. Prénoms seulement. */
function honneurPublic_() {
  var cache = CacheService.getScriptCache(), c = cache.get('honneur_public');
  if (c) return JSON.parse(c);
  var prenom = function (n) { return String(n).trim().split(/\s+/)[0] || ''; };
  var garde = {};
  honLire_().filter(function (h) { return !h.retire; }).forEach(function (h) { if (!garde[h.genre] || h.mois >= garde[h.genre].mois) garde[h.genre] = h; });
  var o = { liste: ['equipe', 'employe'].filter(function (g) { return garde[g]; }).map(function (g) {
    var h = garde[g];
    return { genre: g, mois: h.mois, station: h.station_nom, prenoms: h.noms.split(/\s*,\s*/).map(prenom).filter(String), pourquoi: h.pourquoi.trim(),
      photo: h.photo_id && h.accord === 'oui' ? 'https://drive.google.com/thumbnail?id=' + h.photo_id + '&sz=w800' : '' };
  }) };
  try { cache.put('honneur_public', JSON.stringify(o), 600); } catch (e) { }
  return o;
}

/** Lecture seule : résumé de l'inventaire (dossiers, tailles, doublons nom + taille, fichiers liés depuis les feuilles) dans l'onglet
 * « Inventaire Drive (résumé) ». Ne supprime rien. Feuilles lues pour les liens : Activity Log et la feuille privée (jamais 1D0k4…). */
function analyseTelegram() {
  var sh = ss_().getSheetByName('Inventaire Drive (Telegram)'), n = sh.getLastRow();
  var vu = {}, rows = sh.getRange(2, 1, n - 1, 8).getValues().filter(function (r) { if (vu[r[3]]) return false; vu[r[3]] = 1; return true; });   // un fichier noté deux fois (1er essai interrompu) ne compte qu'une fois
  var refs = {}, re = /[-\w]{25,60}/g, lus = [];
  [[SpreadsheetApp.openById(DSR_ACTIVITY_LOG_ID), 'Activity Log'], [ss_(), 'feuille privée']].forEach(function (x) {
    x[0].getSheets().forEach(function (s) {
      var nom = s.getName(); if (/^Inventaire Drive/.test(nom) || s.getLastRow() < 1 || s.getLastColumn() < 1) return;
      var r = s.getDataRange(), v = r.getValues(), f = r.getFormulas();
      for (var i = 0; i < v.length; i++) for (var j = 0; j < v[i].length; j++) {
        var t = String(v[i][j]) + ' ' + f[i][j]; if (t.length < 26) continue;
        var m = t.match(re); if (m) m.forEach(function (id) { (refs[id] = refs[id] || {})[x[1] + ' › ' + nom] = 1; });
      }
      lus.push(x[1] + ' › ' + nom);
    });
  });
  var grp = function (r) { return String(r[1]).split(' / ').slice(0, 3).join(' / '); };
  var dos = {}, cle = {}, tot = { n: 0, ko: 0 };
  rows.forEach(function (r) {
    var g = grp(r), k = String(r[2]).toLowerCase() + '|' + r[5];
    var d = dos[g] = dos[g] || { n: 0, ko: 0, lies: 0, ko_lies: 0, dbl: 0, ko_dbl: 0, kado: /kado/i.test(r[1]), prop: {} };
    d.n++; d.ko += Number(r[5]) || 0; d.prop[r[7] || '?'] = 1; tot.n++; tot.ko += Number(r[5]) || 0;
    if (refs[r[3]]) { d.lies++; d.ko_lies += Number(r[5]) || 0; }
    if (Number(r[5]) > 0) (cle[k] = cle[k] || []).push(g);
  });
  rows.forEach(function (r) {   // doublon = même nom et même taille ailleurs (autre dossier ou même dossier)
    var k = String(r[2]).toLowerCase() + '|' + r[5]; if (!(Number(r[5]) > 0) || (cle[k] || []).length < 2) return;
    var d = dos[grp(r)]; d.dbl++; d.ko_dbl += Number(r[5]) || 0;
  });
  var paires = {};
  Object.keys(cle).forEach(function (k) { var l = cle[k]; if (l.length < 2) return; var u = l.filter(function (x, i) { return l.indexOf(x) === i; }).sort(); var p = u.join('  ⟷  '); paires[p] = (paires[p] || 0) + (l.length - 1) * Number(k.split('|')[1]); });
  var out = [['Inventaire Drive — résumé (lecture seule, rien supprimé) — ' + maintenant_(), '', '', '', '', '', '', ''],
    ['Total : ' + tot.n + ' fichiers, ' + Math.round(tot.ko / 1024) + ' Mo. Drive : ' + Math.round(DriveApp.getStorageUsed() / 1048576) + ' Mo utilisés sur ' + Math.round(DriveApp.getStorageLimit() / 1048576) + ' Mo.', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['dossier (3 niveaux)', 'fichiers', 'Mo', 'liés depuis une feuille', 'Mo liés', 'en double (nom + taille)', 'Mo en double', 'propriétaire(s) / Kado']];
  Object.keys(dos).sort(function (a, b) { return dos[b].ko - dos[a].ko; }).forEach(function (g) {
    var d = dos[g]; out.push([g, d.n, Math.round(d.ko / 1024), d.lies, Math.round(d.ko_lies / 1024), d.dbl, Math.round(d.ko_dbl / 1024), Object.keys(d.prop).join(', ') + (d.kado ? ' — KADO (à part)' : '')]);
  });
  out.push(['', '', '', '', '', '', '', '']); out.push(['copies en trop entre dossiers (Mo récupérables si on garde une seule copie)', '', 'Mo', '', '', '', '', '']);
  Object.keys(paires).sort(function (a, b) { return paires[b] - paires[a]; }).slice(0, 60).forEach(function (p) { out.push([p, '', Math.round(paires[p] / 1024), '', '', '', '', '']); });
  out.push(['', '', '', '', '', '', '', '']); out.push(['onglets lus pour les liens : ' + lus.length, lus.join(' ; ').slice(0, 45000), '', '', '', '', '', '']);
  var r = ss_().getSheetByName('Inventaire Drive (résumé)') || ss_().insertSheet('Inventaire Drive (résumé)');
  r.clear(); r.getRange(1, 1, out.length, 8).setValues(out);
  Logger.log('Résumé écrit : ' + out.length + ' lignes');
}

/** Ménage approuvé par le propriétaire le 27/11/2026 : met à la corbeille (récupérable 30 jours) trois dossiers en double.
 * Pour chaque dossier, vérifie d'abord avec l'inventaire que CHAQUE fichier existe ailleurs (même nom, même taille) et qu'aucune
 * feuille (Activity Log, feuille privée) ne le lie ; sinon le dossier n'est pas touché. Résultat écrit dans « Inventaire Drive (résumé) » J1. */
function menageTelegram() {
  var plan = [
    { racine: '1pU6LF5CKXw6In9GtZ6ExGStRdmeOBZPs', nom: 'ChatExport_2026-11-09 (1)', chemin: '10 Exports WhatsApp & Telegram / ChatExport_2026-11-09 (1)', ref: 'Telegram Desktop / ChatExport_2026-11-09 (1)' },
    { racine: '15aaDHT_86Fh7zFQL16E2nTJ6MdYbepFb', nom: 'ChatExport_2026-11-22', chemin: 'Telegram Desktop / ChatExport_2026-11-22', ref: 'Telegram Desktop / ChatExport_2026-11-25 (1)' },
    { racine: '', id: '1i_iqLHLu89Iwja99iLjEgC1pVDRqrG-N', chemin: '09 Photos Telegram 2023', ref: '' }   // ref vide : n'importe où ailleurs
  ];
  var sh = ss_().getSheetByName('Inventaire Drive (Telegram)'), vu = {};
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().filter(function (r) { if (vu[r[3]]) return false; vu[r[3]] = 1; return true; });
  var refs = {}, re = /[-\w]{25,60}/g;
  [SpreadsheetApp.openById(DSR_ACTIVITY_LOG_ID), ss_()].forEach(function (x) {
    x.getSheets().forEach(function (s) {
      if (/^Inventaire Drive/.test(s.getName()) || s.getLastRow() < 1 || s.getLastColumn() < 1) return;
      var r = s.getDataRange(), v = r.getValues(), f = r.getFormulas();
      for (var i = 0; i < v.length; i++) for (var j = 0; j < v[i].length; j++) { var t = String(v[i][j]) + ' ' + f[i][j]; if (t.length > 25) (t.match(re) || []).forEach(function (id) { refs[id] = 1; }); }
    });
  });
  var dans = function (d, p) { d = String(d); return d === p || d.indexOf(p + ' / ') === 0; };
  var cle = function (r) { return String(r[2]).toLowerCase() + '|' + r[5]; };
  var bilan = [];
  plan.forEach(function (x) {
    var a = rows.filter(function (r) { return dans(r[1], x.chemin); });
    var b = {}; rows.forEach(function (r) { if (!dans(r[1], x.chemin) && (!x.ref || dans(r[1], x.ref))) b[cle(r)] = 1; });
    var manque = a.filter(function (r) { return !b[cle(r)]; }), lies = a.filter(function (r) { return refs[r[3]]; });
    var mo = Math.round(a.reduce(function (s, r) { return s + (Number(r[5]) || 0); }, 0) / 1024);
    if (!a.length || manque.length || lies.length) { bilan.push('PAS TOUCHÉ — ' + x.chemin + ' : ' + a.length + ' fichiers, ' + manque.length + ' sans copie ailleurs' + (manque.length ? ' (ex. ' + manque.slice(0, 5).map(function (r) { return r[2]; }).join(', ') + ')' : '') + ', ' + lies.length + ' liés depuis une feuille'); return; }
    var dos = x.id ? DriveApp.getFolderById(x.id) : null;
    if (!dos) { var it = DriveApp.getFolderById(x.racine).getFoldersByName(x.nom), l = []; while (it.hasNext()) l.push(it.next()); if (l.length !== 1) { bilan.push('PAS TOUCHÉ — ' + x.chemin + ' : ' + l.length + ' dossiers de ce nom'); return; } dos = l[0]; }
    var prop = ''; try { prop = dos.getOwner().getEmail(); } catch (e) { }
    if (prop !== Session.getEffectiveUser().getEmail()) { bilan.push('PAS TOUCHÉ — ' + x.chemin + ' : propriétaire ' + prop); return; }
    dos.setTrashed(true);
    bilan.push('À LA CORBEILLE — ' + x.chemin + ' (id ' + dos.getId() + ') : ' + a.length + ' fichiers, ' + mo + ' Mo, tous avec une copie ailleurs, aucun lié');
  });
  var msg = 'Ménage du ' + maintenant_() + ' — ' + bilan.join(' || ');
  (ss_().getSheetByName('Inventaire Drive (résumé)') || ss_().insertSheet('Inventaire Drive (résumé)')).getRange(1, 10).setValue(msg);
  Logger.log(msg);
}
