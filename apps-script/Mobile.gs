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

var MOBILE_VERSION = 'm1.21';
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
  return ROLES_APPAREIL[station] || (station.indexOf('TR:') === 0 ? 'transferts' : '');
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
  if (app.role === 'transferts') return { version: MOBILE_VERSION, mode: 'transferts', utilisateur: { id: app.id, nom: app.nom, role: app.role } };
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
    version: MOBILE_VERSION, devise: CURRENCY, station: { code: st.code, nom: st.nom }, date: date, jours_max: JOURS_DIFFERE_MAX,
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
  if (!app.role || app.role === 'transferts') throw new UserError('Accès réservé à la direction.', 'FORBIDDEN');
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
    default: throw new UserError('Action inconnue.');
  }
}

function listeAppareils_() {
  var noms = {}; stations_().forEach(function (s) { noms[String(s.code)] = s.nom; });
  noms.DIRECTION = 'Direction'; noms.COMPTA = 'Comptabilité';
  var tr = agentsTransfertsSeuls_(); tr.forEach(function (t) { noms[t.code] = t.nom; });
  return {
    transferts: tr,
    stations: stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return { code: s.code, nom: s.nom }; }),
    appareils: lire_(APPAREILS).map(function (a) {
      return { code: String(a.code_activation), station: String(a.station), station_nom: noms[String(a.station).trim()] || String(a.station),
        actif: estVrai_(a.actif), utilise: !!String(a.jeton || '').trim(), enregistre_le: String(a.enregistre_le || ''),
        partage: !!ROLES_APPAREIL[String(a.station).trim().toUpperCase()] || String(a.station).trim().toUpperCase() === STATION_FORMATION, nb: appListe_(a.appareil_id).length,
        max: String(a.station).trim().toUpperCase() === STATION_FORMATION ? MAX_APPAREILS_FORMATION : MAX_APPAREILS_ROLE,
        dernier_contact: String(a.dernier_contact || ''), nom: String(a.nom || ''), push: !!String(a.push || '').trim() };
    })
  };
}

function ajouterAppareil_(ctx, p) {
  var station = String(p.station || '').trim().toUpperCase();
  var valides = stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return String(s.code); }).concat(['DIRECTION', 'COMPTA'])
    .concat(agentsTransfertsSeuls_().map(function (t) { return t.code; }));
  if (valides.indexOf(station) < 0) throw new UserError('Station inconnue.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var existants = lire_(APPAREILS).map(function (a) { return String(a.code_activation).trim().toUpperCase(); });
    var prefixe = station === 'DIRECTION' ? 'DIR' : station === 'COMPTA' ? 'CPT' : station.replace(/^TR:/, ''), code;
    do { code = prefixe + '-' + (1000 + Math.floor(Math.random() * 9000)); } while (existants.indexOf(code) >= 0);
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
  if (!id || !(phLire_(15000).some(function (r) { return r.fichier_id === id; }) || bdLire_(5000).some(function (r) { return r.fichier_id === id; }))) throw new UserError('Photo introuvable.');
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
  var jauges = {};   // dernier niveau du manomètre tapé dans l'application
  phLire_(3000).forEach(function (r) {
    if (PH_JAUGES.indexOf(r.type) < 0 || !regl[r.station] || !(Number(r.compteur) > 0)) return;
    var j = jauges[r.station];
    if (!j || r.pris_le > j.pris_le) jauges[r.station] = { pris_le: r.pris_le, pct: Number(r.compteur) };
  });
  var out = Object.keys(regl).map(function (code) {
    var g = regl[code], hh = hist[code] || {}, aa = app[code] || {}, jours = {};
    Object.keys(hh).forEach(function (d) { jours[d] = 1; }); Object.keys(aa).forEach(function (d) { jours[d] = 1; });
    var ds = Object.keys(jours).sort(), serie = ds.map(function (d) {
      var x = hh[d] || [0, 0, 0, 0], a = aa[d];
      return { d: d, vendu: a ? a[0] : x[0], recu: Math.max(a ? a[1] : 0, x[1], x[2], x[3]) };
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
    var j = jauges[code];
    return { code: code, nom: noms[code] || code, capacite: g.capacite, methode: methode, depuis: depuis, source: g.source, recu: Math.round(recu), vendu: Math.round(vendu),
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
/** Pour le téléphone d'une station : ses messages actifs, d'avant-hier à dans 14 jours. */
function messagesStation_(app) {
  if (app.role || !app.station) throw new UserError('Réservé au téléphone de la station.', 'FORBIDDEN');
  var code = String(app.station.code), du = msgJour_(-2), au = msgJour_(14);
  var miennes = {}; repLire_().forEach(function (r) { if (r.station === code) miennes[r.msg_id] = r.texte; });   // la dernière réponse de cette station à chaque message
  var l = msgLire_().filter(function (m) { return m.actif && (m.cible === 'TOUS' || m.cible === code) && m.date >= du && m.date <= au; })
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
function pushToken_(app, p) {
  if (app.role || !app.station) return { garde: false };
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
  var noms = { TOUS: 'Toutes les stations' }; stations_().forEach(function (s) { noms[String(s.code)] = String(s.nom); });
  var du = msgJour_(-30), parMsg = {};
  repLire_().forEach(function (r) { (parMsg[r.msg_id] = parMsg[r.msg_id] || []).push({ station: r.station, nom: noms[r.station] || r.station, ecrit_le: r.ecrit_le, texte: r.texte }); });
  return { aujourdhui: msgJour_(0), stations: stations_().filter(function (s) { return estVrai_(s.actif); }).map(function (s) { return { code: String(s.code), nom: String(s.nom) }; }),
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
  if (app.role || !app.station) throw new UserError('Réservé au téléphone de la station.', 'FORBIDDEN');
  var code = String(app.station.code), texte = String(p.texte || '').trim().slice(0, 500), cid = String(p.cid || '').slice(0, 40);
  if (texte.length < 3) throw new UserError('Écrivez votre réponse.');
  var m = msgLire_().filter(function (x) { return x.id === Number(p.msg_id); })[0];
  if (!m || !m.actif || (m.cible !== 'TOUS' && m.cible !== code)) throw new UserError('Ce message n\'existe plus.');
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
  if (cible !== 'TOUS' && !stations_().some(function (s) { return String(s.code) === cible && estVrai_(s.actif); })) throw new UserError('Station inconnue.');
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
