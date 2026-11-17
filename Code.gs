
const CONFIG = {
  // ---- À MODIFIER ----
  DESTINATAIRES: 'exqmple@email.com', // e-mails de l'équipe, séparés par des virgules
  SOCIETE: 'YOUR COMPANY NAME',         // nom affiché dans l'e-mail et les messages WhatsApp
  SIGNATURE: 'Maini Lotfi, Rust Software Engineer', // phrase affichée en bas de l'e-mail
  FEUILLE_LOCATIONS: 'FEUILLE',  // nom EXACT de l'onglet des locations (en bas de l'écran). Si introuvable : 1er onglet.
  HEURE_ENVOI: 8,                  // heure d'envoi (l'envoi a lieu dans l'heure qui suit)
  FUSEAU: 'Africa/Algiers',
  TEXTE_DATE: 'JMA',               // si une date est du TEXTE ambigu : 'MJA' = mois/jour/année, 'JMA' = jour/mois/année

  // ---- Colonnes de l'onglet des locations (lettres) : vérifie qu'elles correspondent à ton fichier ----
  LIGNE_ENTETE: 1,
  COL: {
    CLIENT: 'A',
    VEHICULE: 'B',
    PLAQUE: 'C',
    KM_LIVRAISON: 'D',     // km au départ
    KM_RESTITUTION: 'E',   // si rempli => voiture considérée comme rendue
    DATE_LIVRAISON: 'G',
    DATE_RESTITUTION: 'H',
    HEURE_LIVRAISON: 'I',
    CONTRAT: 'J',
    TOTAL_TTC: 'N',
    COMMERCIAL: 'O',
    TELEPHONE: 'P',
    CAUTION: 'Q',
    REMARQUE: 'R',
    PAYE: 'S',
    RESTE: 'T',
    DETTE: 'U',
    CONSOMMATION: 'F',     // calculé automatiquement : km restitution - km livraison
    PRIX_UNITAIRE: 'K',    // prix par jour (à saisir)
    VALEUR: 'L',           // calculé automatiquement : nombre de jours (restitution - livraison)
    TOTAL_HT: 'M'          // calculé automatiquement : prix unitaire x nombre de jours
  },

  // ---- Calculs automatiques dans la feuille des locations (créés par « Créer les onglets ») ----
  TVA: 0.19,                         // TOTAL TTC = TOTAL HT x (1 + TVA)
  CALCUL_AUTO: true,                 // true = RESTE / DETTE sont calculés par la feuille sur le TOTAL HT (RESTE = TOTAL HT - PAYE ; DETTE = RESTE d'une location terminée). false = ancien mode manuel
  LIGNES_FORMULES: 1500,             // nombre de lignes pré-remplies avec les formules
  COULEUR_ETAT: ['DATE_RESTITUTION'], // colonnes de l'état de location qui changent de couleur : vert = date pas encore atteinte · orange = date d'aujourd'hui · rouge = date dépassée (ajoute 'CLIENT' pour colorer aussi le nom)

  // ---- Vidange : choix proposés dans la liste de l'onglet « Voitures » (colonne « Intervalle vidange »), à choisir voiture par voiture
  INTERVALLES_VIDANGE: [5000, 6000, 7000, 8000, 9000, 10000, 12000, 15000, 20000],
  INTERVALLE_VIDANGE_DEFAUT: 10000,  // valeur mise à une nouvelle voiture (à changer ensuite si besoin)

  // ---- Réservations
  JOURS_RESERVATIONS: 7,             // réservations affichées : départ dans les X prochains jours (modifiable dans le menu « Choisir le contenu »)

  // ---- Règles ----
  SEUIL_VIDANGE_KM: 1000,            // alerte "bientôt" quand il reste ce nombre de km ou moins
  ALERTE_DOC_JOURS: [30, 15, 7, 3, 1, 0], // jours avant expiration pour envoyer un rappel
  RETARD_MAX_JOURS: 30,              // on n'affiche que les retards / expirations des 30 derniers jours
  JOURS_RETOURS_RECENTS: 1,          // « retours enregistrés » : restitution d'hier et d'aujourd'hui
  MAX_LIGNES_DETTES: 15,
  DETTE_MIN: 100,                    // les montants inférieurs sont ignorés (parasites du type "6J")
  DETTE_ESTIMEE_SI_NON_PAYE: true,   // si DETTE et RESTE sont vides mais que la colonne PAYE / REMARQUE dit "did not pay" => dette estimée = TOTAL TTC
  // Rubriques par défaut de l'e-mail. Le menu « Flotte > Choisir le contenu de l'e-mail » les remplace : ce que tu coches là est utilisé en priorité.
  RUBRIQUES: {
    PROGRAMME: true,             // retours et livraisons d'aujourd'hui et de demain
    RETARDS: true,               // retours en retard (avec lien WhatsApp)
    RETOURS_ENREGISTRES: true,   // retours d'hier / d'aujourd'hui avec km parcourus
    DETTES: true,                // dettes clients
    VIDANGES: true,              // vidanges proches ou dépassées (onglet « Voitures »)
    DOCUMENTS: true,             // assurance / contrôle technique qui expirent (onglet « Voitures »)
    LAVAGES: true,               // lavages d'hier / d'aujourd'hui (onglet « Lavages »)
    ENTRETIEN: true,             // réparations et achats « À faire » (onglet « Entretien »)
    RESERVATIONS: true,          // réservations à venir (onglet « Réservations »)
    BILAN: true                  // bilan financier : totaux du mois (locations, dettes, lavages, entretien)
  },
  JOURS_LAVAGES: 1,                  // lavages affichés : aujourd'hui + X jour(s) avant
  JOUR_ENTRETIEN: 1,                 // jour d'envoi des réparations en attente : 1 = lundi … 0 = dimanche, -1 = tous les jours
  ENVOYER_SI_VIDE: false,            // true = envoyer même s'il n'y a rien à signaler

  // ---- Notifications ----
  TELEGRAM_TOKEN: '',      // jeton du bot (créé avec @BotFather sur Telegram). Laisser vide = désactivé
  TELEGRAM_CHAT_ID: '',    // identifiant de ta discussion ou de ton groupe Telegram
  ALERTE_AVANT_HEURES: 2,  // alerte quand un retour / une livraison est prévu dans les X prochaines heures
  ALERTE_PAR_EMAIL: true,  // envoie aussi l'alerte par e-mail (notification push sur le téléphone)
  URL_DASHBOARD: ''        // lien du tableau de bord web (facultatif : détecté automatiquement après déploiement)
};

/*  MENU */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Flotte ')
    .addItem('1. Créer les onglets (1re fois)', 'initialiserOnglets')
    .addItem('Rafraîchir les couleurs des dates (seulement)', 'rafraichirCouleurs')
    .addItem('Diagnostic : où sont les erreurs ?', 'afficherDiagnostic')
    .addItem('Réparer les formules (erreurs, cases qui ne calculent pas)', 'reparerFormules')
    .addItem('2. Choisir le contenu de l’e-mail', 'choisirContenu')
    .addItem('3. Envoyer un e-mail de test', 'testerRappel')
    .addItem('4. Activer l’envoi automatique', 'activerEnvoiAutomatique')
    .addItem('Désactiver l’envoi automatique', 'desactiverEnvoiAutomatique')
    .addSeparator()
    .addItem('5. Activer les alertes horaires', 'activerAlertesHoraires')
    .addItem('Désactiver les alertes horaires', 'desactiverAlertesHoraires')
    .addItem('6. Tester Telegram', 'testerTelegram')
    .addToUi();
  toastOuverture_();
}

function activerEnvoiAutomatique() {
  desactiverEnvoiAutomatique();
  ScriptApp.newTrigger('rappelQuotidien').timeBased()
    .atHour(CONFIG.HEURE_ENVOI).everyDays(1).inTimezone(CONFIG.FUSEAU).create();
  SpreadsheetApp.getUi().alert('Envoi automatique activé : tous les jours vers ' + CONFIG.HEURE_ENVOI + 'h.');
}

function desactiverEnvoiAutomatique() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'rappelQuotidien')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function testerRappel() {
  const n = construireEtEnvoyer_(true);
  SpreadsheetApp.getUi().alert('E-mail envoyé à : ' + CONFIG.DESTINATAIRES + '\n' + n + ' rappel(s) trouvé(s).');
}

/** Appelée automatiquement chaque jour par le déclencheur. */
function rappelQuotidien() {
  construireEtEnvoyer_(false);
}

/*  OUTILS (dates, nombres)  */

const JOURS_FR_ = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
const MOIS_FR_ = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];

let TZ_FEUILLE_CACHE_ = null;
function fuseauFeuille_() {
  if (!TZ_FEUILLE_CACHE_) TZ_FEUILLE_CACHE_ = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || CONFIG.FUSEAU;
  return TZ_FEUILLE_CACHE_;
}

function colIndex_(lettre) {
  let n = 0;
  for (const ch of String(lettre).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Toutes les dates sont normalisées à minuit UTC : plus de décalage de fuseau ni d'heure d'été. */
function aujourdhui_() {
  const p = Utilities.formatDate(new Date(), CONFIG.FUSEAU, 'yyyy-MM-dd').split('-');
  return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
}

function decale_(d, jours) {
  return new Date(d.getTime() + jours * 86400000);
}

function memeJour_(a, b) {
  return a.getTime() === b.getTime();
}

function memeMois_(a, b) {
  return a.getUTCFullYear() === b.getUTCFullYear() && a.getUTCMonth() === b.getUTCMonth();
}

function joursEntre_(a, b) { // b - a, en jours
  return Math.round((b - a) / 86400000);
}

function enDate_(v) {
  if (v instanceof Date && !isNaN(v)) {
    // on lit le jour tel qu'il apparaît dans la feuille (fuseau de la feuille)
    const p = Utilities.formatDate(v, fuseauFeuille_(), 'yyyy-MM-dd').split('-');
    return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
  }
  if (typeof v === 'string') {
    const s = v.trim();
    let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) {
      const a = +m[1], b = +m[2], y = +m[3];
      let mois, jour;
      if (a > 12) { jour = a; mois = b; }                  // 25/03/2026 : forcément jour/mois
      else if (b > 12) { mois = a; jour = b; }             // 03/25/2026 : forcément mois/jour
      else if (CONFIG.TEXTE_DATE === 'MJA') { mois = a; jour = b; }
      else { jour = a; mois = b; }
      return new Date(Date.UTC(y, mois - 1, jour));
    }
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  }
  return null;
}

/** Lit un nombre, y compris « 150.000,00 DA », « 30,000 », « STILL 270000 DA ». */
function nombre_(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  let s = String(v == null ? '' : v).replace(/[\s\u00A0]/g, '').replace(/[^\d.,-]/g, '');
  if (!/\d/.test(s)) return 0;
  const neg = s.charAt(0) === '-';
  s = s.replace(/-/g, '');
  const pos = Math.max(s.lastIndexOf(','), s.lastIndexOf('.'));
  if (pos >= 0) {
    const fin = s.slice(pos + 1);
    if (fin.length === 3) s = s.replace(/[.,]/g, '');                      // séparateur de milliers
    else s = s.slice(0, pos).replace(/[.,]/g, '') + '.' + fin;             // séparateur décimal
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : (neg ? -n : n);
}

/** Montant positif ou 0 (texte sans chiffre, erreur #VALUE!, négatif => 0). */
function montant_(v) {
  if (typeof v === 'number') return v > 0 ? v : 0;
  const s = String(v == null ? '' : v);
  if (!/\d/.test(s)) return 0;
  const n = nombre_(s);
  return n > 0 ? n : 0;
}

/** Kilométrage : nombre, ou null si vide / "/" / erreur. */
function km_(v) {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  const s = String(v);
  if (!/\d/.test(s)) return null;
  return nombre_(s);
}

function kmParcourus_(x) {
  return (x.kmDep !== null && x.kmRet !== null && x.kmRet >= x.kmDep) ? x.kmRet - x.kmDep : null;
}

function fmt_(n) {
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '\u00A0');
}

function cle_(plaque) {
  return String(plaque || '').toUpperCase().replace(/\s+/g, '');
}

function dateFr_(d) {
  return Utilities.formatDate(d, 'GMT', 'dd/MM/yyyy');
}

function dateLongue_(d) {
  const j = d.getUTCDate();
  return JOURS_FR_[d.getUTCDay()] + ' ' + (j === 1 ? '1er' : j) + ' ' + MOIS_FR_[d.getUTCMonth()] + ' ' + d.getUTCFullYear();
}

function maj_(s) {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function pl_(n, sing, plur) {
  return n + ' ' + (n > 1 ? plur : sing);
}

function heure_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, fuseauFeuille_(), 'HH:mm').replace(':', 'h');
  }
  return String(v == null ? '' : v).trim().replace(/(\d)H(\d*)/i, '$1h$2');
}

function minutes_(h) {
  const m = String(h || '').match(/(\d{1,2})\D+(\d{2})?/);
  return m ? (+m[1]) * 60 + (+(m[2] || 0)) : 99999;
}

function parHeure_(a, b) {
  return minutes_(a.heure) - minutes_(b.heure);
}

function h_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Numéro au format WhatsApp (213…) ; accepte "0661,23,68,24", "07.73.03.29.95", "0551-85-08-77/0559124336"… */
function telWhatsapp_(s) {
  const parties = String(s || '').split(/[\/;]|,\s+|\s+(?:et|ou)\s+/i);
  for (const p of parties) {
    const t = p.replace(/[^\d+]/g, '');
    if (!t) continue;
    if (t.startsWith('+')) { if (t.length >= 10) return t.slice(1); continue; }
    if (t.startsWith('00')) { if (t.length >= 11) return t.slice(2); continue; }
    if (t.startsWith('0') && t.length === 10) return '213' + t.slice(1);
    if (t.length === 9 && /^[567]/.test(t)) return '213' + t;
  }
  return '';
}

function cautionTxt_(v) {
  const s = String(v == null ? '' : v).trim();
  return /^(none|aucune?|\/|-|0)?$/i.test(s) ? '' : s;
}

/** Traduit en français les phrases anglaises courantes saisies dans la feuille (affichage seulement). */
const TRADUCTIONS_ = [
  [/payed by bank check all the a?mm?ount/gi, 'Payé par chèque bancaire, montant total'],
  [/they pay by check bank/gi, 'Paiement par chèque bancaire'],
  [/will pay when he returns? the car/gi, 'Paiera au retour du véhicule'],
  [/did not pay yet/gi, 'Pas encore payé'],
  [/did not pay/gi, 'Non payé'],
  [/payed all the a?mm?ount/gi, 'Montant payé en totalité'],
  [/fully pa(?:y)?i?e?d/gi, 'Entièrement payé'],
  [/still (\d[\d.,\s]*)/gi, 'Reste $1'],
  [/\bpayed\b/gi, 'payé'],
  [/\bpassport\b/gi, 'passeport']
];

function fr_(s) {
  let t = String(s == null ? '' : s);
  TRADUCTIONS_.forEach(([re, fr]) => { t = t.replace(re, fr); });
  return t;
}

/* ====================== LECTURE DES DONNÉES ====================== */

/** Total HT d'une ligne : colonne TOTAL HT, sinon TOTAL TTC / (1 + TVA). */
function baseHT_(l, g) {
  const ht = montant_(l[g('TOTAL_HT')]);
  if (ht) return ht;
  const ttc = montant_(l[g('TOTAL_TTC')]);
  return ttc ? Math.round(ttc / (1 + CONFIG.TVA) * 100) / 100 : 0;
}

function lireLocations_(ss, aujourd, demain, r, info) {
  const f = ss.getSheetByName(CONFIG.FEUILLE_LOCATIONS) || ss.getSheets()[0];
  const n = f.getLastRow() - CONFIG.LIGNE_ENTETE;
  if (n < 1) return;
  const c = CONFIG.COL;
  const g = k => colIndex_(c[k]);
  const nbCols = Math.max(...Object.keys(c).map(g)) + 1;
  const lignes = f.getRange(CONFIG.LIGNE_ENTETE + 1, 1, n, nbCols).getValues();
  const debutRecents = decale_(aujourd, -CONFIG.JOURS_RETOURS_RECENTS);
  const limiteRetard = decale_(aujourd, -CONFIG.RETARD_MAX_JOURS);
  const reNonPaye = /not pay|non pay|impay|pas pay/i;
  const rePaye = /payed|paid|paied|fully|total|vers/i;

  lignes.forEach(l => {
    const client = String(l[g('CLIENT')]).trim();
    if (!client) return;
    const dl = enDate_(l[g('DATE_LIVRAISON')]);
    const dr = enDate_(l[g('DATE_RESTITUTION')]);
    const brutRet = String(l[g('KM_RESTITUTION')]).trim();
    const rendu = brutRet !== '' && brutRet.charAt(0) !== '#' && !/^[\/\-–.\s]+$/.test(brutRet);
    const plaque = String(l[g('PLAQUE')]).trim();
    const kmDep = km_(l[g('KM_LIVRAISON')]);
    const kmRet = rendu ? km_(l[g('KM_RESTITUTION')]) : null;
    const paye = String(l[g('PAYE')]).trim();
    const remarque = String(l[g('REMARQUE')]).trim();

    // dette : colonne DETTE, sinon RESTE, sinon estimation (TOTAL TTC) si "did not pay"
    let dette = montant_(l[g('DETTE')]);
    let estimee = false;
    if (dette < CONFIG.DETTE_MIN) {
      const reste = montant_(l[g('RESTE')]);
      // avec les calculs automatiques, RESTE n'est une dette que si la location est terminée (voiture rendue ou date de restitution passée)
      if (!CONFIG.CALCUL_AUTO || rendu || (dr && dr < aujourd)) dette = reste;
    }
    if (dette < CONFIG.DETTE_MIN && CONFIG.DETTE_ESTIMEE_SI_NON_PAYE && !CONFIG.CALCUL_AUTO) {
      const nonPaye = reNonPaye.test(paye) || (reNonPaye.test(remarque) && !rePaye.test(paye));
      const ht = baseHT_(l, g);
      if (nonPaye && ht >= CONFIG.DETTE_MIN) { dette = ht; estimee = true; }
    }
    if (dette < CONFIG.DETTE_MIN) dette = 0;

    const x = {
      client: client,
      vehicule: String(l[g('VEHICULE')]).trim(),
      plaque: plaque,
      heure: heure_(l[g('HEURE_LIVRAISON')]),
      contrat: String(l[g('CONTRAT')]).trim(),
      commercial: String(l[g('COMMERCIAL')]).trim(),
      tel: String(l[g('TELEPHONE')]).trim(),
      caution: cautionTxt_(l[g('CAUTION')]),
      remarque: remarque,
      kmDep: kmDep,
      kmRet: kmRet,
      dl: dl,
      dr: dr,
      dette: dette,
      estimee: estimee,
      retard: 0
    };

    // dernier kilométrage connu par voiture (sert aux alertes de vidange)
    if (plaque) {
      const mk = Math.max(kmDep || 0, kmRet || 0);
      const k = cle_(plaque);
      if (mk > (info.kmParPlaque[k] || 0)) info.kmParPlaque[k] = mk;
    }

    if (dl && memeJour_(dl, aujourd)) r.livAuj.push(x);
    else if (dl && memeJour_(dl, demain)) r.livDem.push(x);

    if (dr && !rendu) {
      if (memeJour_(dr, aujourd)) r.retAuj.push(x);
      else if (memeJour_(dr, demain)) r.retDem.push(x);
      else if (dr < aujourd) {
        x.retard = joursEntre_(dr, aujourd);
        if (x.retard <= CONFIG.RETARD_MAX_JOURS) r.retard.push(x);
      }
    }
    if (rendu && dr && dr >= debutRecents && dr <= aujourd) r.retFaits.push(x);
    if (!rendu && dl && dl <= aujourd && (!dr || dr >= limiteRetard)) info.enCours++;
    if (dl && memeMois_(dl, aujourd)) { info.caMois.nb++; info.caMois.total += baseHT_(l, g); info.caMois.ttc += montant_(l[g('TOTAL_TTC')]); }
    if (x.dette > 0) r.dettes.push(x);
  });

  r.dettes.sort((a, b) => b.dette - a.dette);
  r.retard.sort((a, b) => b.retard - a.retard);
  [r.retAuj, r.retDem, r.livAuj, r.livDem].forEach(liste => liste.sort(parHeure_));
}

function lireVoitures_(ss, aujourd, r, info) {
  const f = ss.getSheetByName('Voitures');
  if (!f || f.getLastRow() < 2) return;
  const v = f.getRange(2, 1, f.getLastRow() - 1, 10).getValues();
  v.forEach(l => {
    const plaque = String(l[0]).trim();
    if (!plaque) return;
    const modele = String(l[1]).trim();
    // km actuel = le plus grand entre la saisie manuelle et le dernier km vu dans les locations
    const kmActuel = Math.max(nombre_(l[2]), info.kmParPlaque[cle_(plaque)] || 0);
    if (kmActuel > 0 && l[3] !== '' && l[4] !== '') {
      const reste = nombre_(l[3]) + nombre_(l[4]) - kmActuel;
      if (reste <= CONFIG.SEUIL_VIDANGE_KM) r.vidanges.push({ plaque: plaque, modele: modele, reste: reste, km: kmActuel });
    }
    [[l[8], 'Assurance'], [l[9], 'Contrôle technique']].forEach(([val, nom]) => {
      const d = enDate_(val);
      if (!d) return;
      const j = joursEntre_(aujourd, d);
      if (CONFIG.ALERTE_DOC_JOURS.indexOf(j) >= 0 || (j < 0 && j >= -CONFIG.RETARD_MAX_JOURS)) {
        r.docs.push({ plaque: plaque, modele: modele, nom: nom, j: j, d: d });
      }
    });
  });
  r.vidanges.sort((a, b) => a.reste - b.reste);
}

function lireLavages_(ss, aujourd, r, info) {
  const f = ss.getSheetByName('Lavages');
  if (!f || f.getLastRow() < 2) return;
  info.lav = { mois: { nb: 0, total: 0 }, general: { nb: 0, total: 0 } };
  const debut = decale_(aujourd, -CONFIG.JOURS_LAVAGES);
  f.getRange(2, 1, f.getLastRow() - 1, 6).getValues().forEach(l => {
    const plaque = String(l[1]).trim(), prix = montant_(l[3]);
    if (!plaque && !prix) return;
    const brut = String(l[5]).trim();
    const statut = /^oui|^yes|^pay/i.test(brut) ? 'oui' : /^non|^no\b|^impay/i.test(brut) ? 'non' : '';
    const d = enDate_(l[0]);
    if (prix > 0) {
      info.lav.general.nb++; info.lav.general.total += prix;
      if (d && memeMois_(d, aujourd)) { info.lav.mois.nb++; info.lav.mois.total += prix; }
      if (statut === 'non') { info.lavNonPaye.nb++; info.lavNonPaye.total += prix; }
    }
    if (d && d >= debut && d <= aujourd) r.lavages.push({ date: d, plaque: plaque, type: String(l[2]).trim(), prix: prix, laveur: String(l[4]).trim(), statut: statut });
  });
  r.lavages.sort((a, b) => b.date - a.date);
}

/** Totaux d'entretien (dépenses = interventions non « À faire ») + liste des interventions « À faire » si lister = true. */
function lireEntretien_(ss, aujourd, r, info, lister) {
  const f = ss.getSheetByName('Entretien');
  if (!f || f.getLastRow() < 2) return;
  info.ent = { mois: { nb: 0, total: 0 }, general: { nb: 0, total: 0 }, attente: { nb: 0, total: 0 } };
  f.getRange(2, 1, f.getLastRow() - 1, 7).getValues().forEach(l => {
    const prix = montant_(l[4]), d = enDate_(l[0]), aFaire = String(l[6]).trim() === 'À faire';
    if (!(l[1] || l[3] || prix)) return;
    if (aFaire) {
      info.ent.attente.nb++; info.ent.attente.total += prix;
      if (lister) r.attente.push({ date: d, plaque: String(l[1]), type: String(l[2]), desc: String(l[3]), prix: prix });
    } else if (prix > 0) {
      info.ent.general.nb++; info.ent.general.total += prix;
      if (d && memeMois_(d, aujourd)) { info.ent.mois.nb++; info.ent.mois.total += prix; }
    }
  });
}

/* NOTIFICATIONS (Telegram, alertes horaires, Sheet)  */

function urlDashboard_() {
  if (CONFIG.URL_DASHBOARD) return CONFIG.URL_DASHBOARD;
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function envoyerTelegram_(texte) {
  if (!CONFIG.TELEGRAM_TOKEN || !CONFIG.TELEGRAM_CHAT_ID) return false;
  try {
    const rep = UrlFetchApp.fetch('https://api.telegram.org/bot' + CONFIG.TELEGRAM_TOKEN + '/sendMessage', {
      method: 'post', muteHttpExceptions: true,
      payload: { chat_id: CONFIG.TELEGRAM_CHAT_ID, text: texte, parse_mode: 'HTML', disable_web_page_preview: 'true' }
    });
    return rep.getResponseCode() === 200;
  } catch (e) { return false; }
}

function texteTelegram_(c) {
  const r = c.r, L = [];
  L.push('<b>' + h_(CONFIG.SOCIETE) + ' — ' + dateFr_(c.aujourd) + '</b>');
  L.push('Aujourd’hui : ' + pl_(r.retAuj.length, 'retour', 'retours') + ' · ' + pl_(r.livAuj.length, 'livraison', 'livraisons'));
  L.push('Demain : ' + pl_(r.retDem.length, 'retour', 'retours') + ' · ' + pl_(r.livDem.length, 'livraison', 'livraisons'));
  if (r.retard.length) {
    L.push(''); L.push('<b>' + pl_(r.retard.length, 'retour en retard', 'retours en retard') + '</b>');
    r.retard.slice(0, 8).forEach(x => L.push('• ' + h_(x.client) + ' — ' + h_(x.vehicule) + ' (' + x.retard + ' j)'));
  }
  const bloc = (titre, liste) => {
    if (!liste.length) return;
    L.push(''); L.push('<b>' + titre + '</b>');
    liste.slice(0, 10).forEach(x => L.push('• ' + (x.heure ? h_(x.heure) + ' — ' : '') + h_(x.client) + ' — ' + h_(x.vehicule) + (x.plaque ? ' (' + h_(x.plaque) + ')' : '')));
  };
  bloc('Retours aujourd’hui', r.retAuj);
  bloc('Livraisons aujourd’hui', r.livAuj);
  bloc('Retours demain', r.retDem);
  bloc('Livraisons demain', r.livDem);
  if (r.reservations.length) {
    L.push(''); L.push('<b>Réservations à venir</b>');
    r.reservations.slice(0, 8).forEach(x => L.push('• ' + (x.dans === 0 ? 'Aujourd’hui' : x.dans === 1 ? 'Demain' : dateFr_(x.du).slice(0, 5)) + (x.heure ? ' ' + h_(x.heure) : '') +
      ' — ' + h_(x.client) + ' — ' + h_(x.vehicule) + (x.lieu ? ' (' + h_(x.lieu) + ')' : '')));
  }
  const u = urlDashboard_();
  if (u) { L.push(''); L.push('<a href="' + u + '">Ouvrir le tableau de bord</a>'); }
  return L.join('\n');
}

function testerTelegram() {
  const ok = envoyerTelegram_('<b>' + h_(CONFIG.SOCIETE) + '</b> — les notifications Telegram fonctionnent.');
  SpreadsheetApp.getUi().alert(ok ? 'Message Telegram envoyé.' : 'Échec : vérifie TELEGRAM_TOKEN et TELEGRAM_CHAT_ID dans CONFIG (et envoie d’abord « /start » à ton bot).');
}

function activerAlertesHoraires() {
  desactiverAlertesHoraires();
  ScriptApp.newTrigger('alerteHoraire').timeBased().everyHours(1).inTimezone(CONFIG.FUSEAU).create();
  SpreadsheetApp.getUi().alert('Alertes activées : chaque heure (entre 7h et 21h), tu es prévenu des retours / livraisons prévus dans les ' +
    CONFIG.ALERTE_AVANT_HEURES + ' prochaines heures et des retards.');
}

function desactiverAlertesHoraires() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alerteHoraire').forEach(t => ScriptApp.deleteTrigger(t));
}

/** Déclenchée chaque heure : prévient une seule fois par événement. */
function alerteHoraire() {
  const c = collecte_(false);
  const hhmm = Utilities.formatDate(new Date(), CONFIG.FUSEAU, 'HH:mm').split(':');
  const maintenant = (+hhmm[0]) * 60 + (+hhmm[1]);
  if (maintenant < 7 * 60 || maintenant > 21 * 60) return;

  const props = PropertiesService.getScriptProperties();
  const cleJour = 'alertes_' + dateFr_(c.aujourd);
  props.getKeys().filter(k => k.indexOf('alertes_') === 0 && k !== cleJour).forEach(k => props.deleteProperty(k)); // nettoyage
  const deja = JSON.parse(props.getProperty(cleJour) || '{}');
  const lignes = [];
  const fenetre = CONFIG.ALERTE_AVANT_HEURES * 60;

  const verifier = (liste, type, texte) => liste.forEach(x => {
    const dans = minutes_(x.heure) - maintenant;
    if (dans < 0 || dans > fenetre) return;
    const k = type + '|' + cle_(x.plaque) + '|' + x.client;
    if (deja[k]) return;
    deja[k] = 1;
    lignes.push(texte + ' ' + (dans < 60 ? 'dans ' + Math.max(dans, 1) + ' min' : 'vers ' + x.heure) + ' — <b>' + h_(x.client) + '</b> — ' + h_(x.vehicule) + (x.plaque ? ' (' + h_(x.plaque) + ')' : ''));
  });
  verifier(c.r.retAuj, 'ret', 'Retour');
  verifier(c.r.livAuj, 'liv', 'Livraison');
  c.r.retard.forEach(x => {
    const k = 'retard|' + cle_(x.plaque) + '|' + x.client;
    if (deja[k]) return;
    deja[k] = 1;
    lignes.push('En retard (' + x.retard + ' j) — <b>' + h_(x.client) + '</b> — ' + h_(x.vehicule));
  });
  if (!lignes.length) return;
  props.setProperty(cleJour, JSON.stringify(deja));

  const u = urlDashboard_();
  envoyerTelegram_('<b>Alerte flotte ' + h_(CONFIG.SOCIETE) + '</b>\n' + lignes.join('\n') + (u ? '\n\n<a href="' + u + '">Tableau de bord</a>' : ''));
  if (CONFIG.ALERTE_PAR_EMAIL) {
    MailApp.sendEmail({
      to: CONFIG.DESTINATAIRES,
      subject: 'Alerte flotte — ' + pl_(lignes.length, 'événement', 'événements'),
      body: htmlVersTexte_(lignes.join('<br>')),
      htmlBody: '<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.7;color:#111827"><h3 style="margin:0 0 8px;color:#a4161a">Alerte flotte</h3>' + lignes.join('<br>') +
        (u ? '<p><a href="' + u + '" style="color:#a4161a;font-weight:bold">Ouvrir le tableau de bord</a></p>' : '') + '</div>',
      name: 'Flotte ' + CONFIG.SOCIETE
    });
  }
}

/** Petite notification dans le Sheet à l'ouverture. */
function toastOuverture_() {
  try {
    const c = collecte_(false);
    const r = c.r;
    const msg = pl_(r.retAuj.length, 'retour', 'retours') + ' · ' + pl_(r.livAuj.length, 'livraison', 'livraisons') + ' aujourd’hui' +
      (r.retard.length ? ' · ' + r.retard.length + ' en retard' : '');
    SpreadsheetApp.getActive().toast(msg, 'Flotte ' + CONFIG.SOCIETE, 10);
  } catch (e) { /* silencieux */ }
}

/*  RÉGLAGES : CONTENU DE L'E-MAIL (liste de choix) */

const RUBRIQUES_DEF_ = [
  { cle: 'PROGRAMME', titre: 'Retours et livraisons', desc: 'Programme d’aujourd’hui et de demain : heure, km au départ, caution, relance WhatsApp' },
  { cle: 'RETARDS', titre: 'Retours en retard', desc: 'Véhicules non restitués à la date prévue, avec lien de relance WhatsApp' },
  { cle: 'RETOURS_ENREGISTRES', titre: 'Retours enregistrés', desc: 'Km au départ, km au retour et km parcourus (hier et aujourd’hui)' },
  { cle: 'DETTES', titre: 'Dettes clients', desc: 'Montants restant dus par client (calculés sur le total HT) et total des dettes' },
  { cle: 'VIDANGES', titre: 'Vidanges', desc: 'Véhicules dont la vidange est proche ou dépassée (onglet « Voitures »)' },
  { cle: 'DOCUMENTS', titre: 'Assurance et contrôle technique', desc: 'Documents qui arrivent à expiration ou sont expirés (onglet « Voitures »)' },
  { cle: 'LAVAGES', titre: 'Lavages', desc: 'Lavages d’hier et d’aujourd’hui, montants et reste à payer aux laveurs (onglet « Lavages »)' },
  { cle: 'ENTRETIEN', titre: 'Réparations et achats en attente', desc: 'Interventions au statut « À faire » (onglet « Entretien »)' },
  { cle: 'RESERVATIONS', titre: 'Réservations à venir', desc: 'Réservations futures (onglet « Réservations ») : période, véhicule, client, lieu de livraison, commercial' },
  { cle: 'BILAN', titre: 'Bilan financier', desc: 'Totaux du mois : locations, dettes, lavages, entretien et réparations' }
];

/** Choix enregistrés depuis le menu ; à défaut, valeurs de CONFIG.RUBRIQUES. */
function reglages_() {
  let saved = {};
  try { saved = JSON.parse(PropertiesService.getScriptProperties().getProperty('reglages') || '{}'); } catch (err) { saved = {}; }
  const R = Object.assign({}, CONFIG.RUBRIQUES, saved.rubriques || {});
  const chaqueJour = typeof saved.entretienChaqueJour === 'boolean' ? saved.entretienChaqueJour : CONFIG.JOUR_ENTRETIEN < 0;
  const jr = parseInt(saved.joursReservations, 10);
  return { R: R, entretienChaqueJour: chaqueJour, joursReservations: jr > 0 ? jr : CONFIG.JOURS_RESERVATIONS };
}

function choisirContenu() {
  const s = reglages_();
  const css = 'body{margin:0;padding:22px 24px;font-family:Arial,Helvetica,sans-serif;color:#111827;font-size:14px}' +
    'h2{margin:0 0 6px;font-size:18px}.intro{margin:0 0 14px;color:#6b7280;line-height:1.5}' +
    '.tools{margin:0 0 8px;font-size:13px}.tools a{color:#a4161a;margin-right:16px;text-decoration:none;font-weight:bold}' +
    '.row{display:flex;gap:12px;align-items:flex-start;padding:11px 12px;border:1px solid #e2e5ea;margin-bottom:-1px;cursor:pointer}' +
    '.row:hover{background:#f9fafb}.row input{margin-top:3px;width:17px;height:17px;accent-color:#a4161a;flex:none}' +
    '.row b{display:block;font-size:14px}.row small{display:block;color:#6b7280;font-size:12px;line-height:1.4;margin-top:2px}' +
    '.opt{margin-top:14px;background:#f6f7f9}' +
    '.bar{margin-top:18px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}' +
    'button{padding:10px 18px;border:1px solid #cfd4da;background:#fff;font-size:13px;font-weight:bold;cursor:pointer;border-radius:2px}' +
    'button.p{background:#a4161a;border-color:#a4161a;color:#fff}#msg{font-size:13px;color:#4b5563}' +
    'select{padding:7px 8px;border:1px solid #cfd4da;font-size:13px;border-radius:2px;background:#fff;flex:none}';
  const js = 'function tout(v){document.querySelectorAll("input[data-k]").forEach(function(c){c.checked=v})}' +
    'function lire(){var R={};document.querySelectorAll("input[data-k]").forEach(function(c){R[c.getAttribute("data-k")]=c.checked});' +
    'return {rubriques:R,entretienChaqueJour:document.getElementById("entretienJ").checked,joursReservations:parseInt(document.getElementById("joursR").value,10)}}' +
    'function enr(test){var m=document.getElementById("msg");m.textContent=test?"Envoi du test…":"Enregistrement…";' +
    'var run=google.script.run.withSuccessHandler(function(){m.textContent=test?"E-mail de test envoyé.":"Choix enregistrés.";' +
    'if(!test)setTimeout(function(){google.script.host.close()},700)}).withFailureHandler(function(e){m.textContent="Erreur : "+e.message});' +
    'if(test){run.envoyerTestAvecChoix(lire())}else{run.enregistrerReglages(lire())}}';
  const optsJ = [[1, 'Aujourd’hui et demain'], [2, '2 prochains jours'], [3, '3 prochains jours'], [7, '7 prochains jours'], [14, '14 prochains jours'], [30, '30 prochains jours']];
  if (!optsJ.some(o => o[0] === s.joursReservations)) optsJ.push([s.joursReservations, s.joursReservations + ' prochains jours']);
  const selectJ = '<select id="joursR">' + optsJ.map(o => '<option value="' + o[0] + '"' + (o[0] === s.joursReservations ? ' selected' : '') + '>' + o[1] + '</option>').join('') + '</select>';
  const lignes = RUBRIQUES_DEF_.map(d =>
    '<label class="row"><input type="checkbox" data-k="' + d.cle + '"' + (s.R[d.cle] ? ' checked' : '') + '><span><b>' + h_(d.titre) + '</b><small>' + h_(d.desc) + '</small></span></label>').join('');
  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' + css + '</style></head><body>' +
    '<h2>Contenu de l’e-mail automatique</h2>' +
    '<p class="intro">Coche les rubriques à inclure. Ce choix est enregistré et utilisé pour l’envoi automatique et pour les tests.</p>' +
    '<div class="tools"><a href="#" onclick="tout(true);return false">Tout cocher</a><a href="#" onclick="tout(false);return false">Tout décocher</a></div>' +
    lignes +
    '<label class="row opt"><input type="checkbox" id="entretienJ"' + (s.entretienChaqueJour ? ' checked' : '') + '><span><b>Réparations en attente : envoyer tous les jours</b>' +
    '<small>Si cette case est décochée, elles sont envoyées une fois par semaine (le lundi).</small></span></label>' +
    '<label class="row opt"><span style="flex:1"><b>Réservations : période affichée</b><small>Réservations dont le départ tombe dans cette période, à partir d’aujourd’hui</small></span>' + selectJ + '</label>' +
    '<div class="bar"><button class="p" onclick="enr(false)">Enregistrer</button><button onclick="enr(true)">Enregistrer et envoyer un test</button><span id="msg"></span></div>' +
    '<script>' + js + '<\/script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(790), 'Contenu de l’e-mail automatique');
}

function enregistrerReglages(choix) {
  const R = {};
  RUBRIQUES_DEF_.forEach(d => { R[d.cle] = !!(choix && choix.rubriques && choix.rubriques[d.cle]); });
  const jr = parseInt(choix && choix.joursReservations, 10);
  PropertiesService.getScriptProperties().setProperty('reglages', JSON.stringify({ rubriques: R, entretienChaqueJour: !!(choix && choix.entretienChaqueJour),
    joursReservations: jr > 0 ? Math.min(jr, 90) : CONFIG.JOURS_RESERVATIONS }));
  return true;
}

function envoyerTestAvecChoix(choix) {
  enregistrerReglages(choix);
  return construireEtEnvoyer_(true);
}

/*  COLLECTE + BILAN  */

function collecte_(forcer) {
  const ss = SpreadsheetApp.getActive();
  const aujourd = aujourdhui_();
  const demain = decale_(aujourd, 1);
  const r = { livAuj: [], livDem: [], retAuj: [], retDem: [], retard: [], retFaits: [], dettes: [], vidanges: [], docs: [], attente: [], lavages: [], reservations: [] };
  const info = { kmParPlaque: {}, enCours: 0, caMois: { nb: 0, total: 0, ttc: 0 }, lav: null, lavNonPaye: { nb: 0, total: 0 }, ent: null };
  lireLocations_(ss, aujourd, demain, r, info);
  lireVoitures_(ss, aujourd, r, info);
  lireLavages_(ss, aujourd, r, info);
  lireReservations_(ss, aujourd, r, reglages_().joursReservations);
  const lister = forcer || reglages_().entretienChaqueJour || aujourd.getUTCDay() === CONFIG.JOUR_ENTRETIEN;
  lireEntretien_(ss, aujourd, r, info, lister);
  return { aujourd: aujourd, demain: demain, r: r, info: info };
}

/** Lignes du bilan financier : { g: groupe, l: libellé, d: détail, v: montant, c: couleur }. */
function bilanLignes_(c) {
  const i = c.info, L = [];
  const mois = MOIS_FR_[c.aujourd.getUTCMonth()];
  const totalDettes = c.r.dettes.reduce((s, x) => s + x.dette, 0);
  L.push({ g: 'Locations', l: 'Locations livrées en ' + mois, d: pl_(i.caMois.nb, 'location', 'locations') + ' · total HT (TTC : ' + fmt_(i.caMois.ttc) + ' DA)', v: i.caMois.total });
  if (c.r.reservations.length) L.push({ g: 'Locations', l: 'Réservations à venir', d: pl_(c.r.reservations.length, 'réservation', 'réservations') + ' · total HT', v: c.r.reservations.reduce((t, x) => t + x.total, 0) });
  L.push({ g: 'Clients', l: 'Dettes clients (HT)', d: pl_(c.r.dettes.length, 'dossier', 'dossiers') + ' avec un solde à encaisser', v: totalDettes, c: totalDettes ? COULEUR_.rouge : null });
  if (i.lav) {
    L.push({ g: 'Lavages', l: 'Lavages de ' + mois, d: pl_(i.lav.mois.nb, 'lavage', 'lavages'), v: i.lav.mois.total });
    L.push({ g: 'Lavages', l: 'Lavages, total général', d: pl_(i.lav.general.nb, 'lavage', 'lavages'), v: i.lav.general.total });
    L.push({ g: 'Lavages', l: 'Reste à payer aux laveurs', d: pl_(i.lavNonPaye.nb, 'lavage non payé', 'lavages non payés'), v: i.lavNonPaye.total, c: i.lavNonPaye.total ? COULEUR_.orange : null });
  }
  if (i.ent) {
    L.push({ g: 'Entretien et réparations', l: 'Dépenses de ' + mois, d: pl_(i.ent.mois.nb, 'intervention', 'interventions'), v: i.ent.mois.total });
    L.push({ g: 'Entretien et réparations', l: 'Dépenses, total général', d: pl_(i.ent.general.nb, 'intervention', 'interventions'), v: i.ent.general.total });
    L.push({ g: 'Entretien et réparations', l: 'En attente (à faire)', d: pl_(i.ent.attente.nb, 'intervention', 'interventions') + ' · montant estimé', v: i.ent.attente.total, c: i.ent.attente.total ? COULEUR_.orange : null });
  }
  return L;
}

/*  DESIGN DE L'E-MAIL   */

const COULEUR_ = { marine: '#a4161a', rouge: '#b91c1c', vert: '#0f766e', orange: '#92400e', violet: '#475569', gris: '#6b7280' };
const BRISE_ = 'word-break:break-word;overflow-wrap:anywhere;';
const POLICE_ = "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif";

const CSS_EMAIL_ =
  'body{margin:0;padding:0;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}' +
  'table{border-collapse:collapse}a{word-break:break-word}' +
  '@media only screen and (min-width:621px){.px{padding-left:28px!important;padding-right:28px!important}}' +
  '@media only screen and (max-width:620px){' +
  '.outer{padding:0!important}.px{padding-left:18px!important;padding-right:18px!important}' +
  '.h1{font-size:21px!important;line-height:27px!important}' +
  '}' +
  '@media (prefers-color-scheme:dark){' +
  '.dk-page{background:#0b0f17!important}.dk-card{background:#151b27!important;border-color:#2a3242!important}' +
  '.dk-card2{background:#1b2230!important;border-color:#2a3242!important}' +
  '.dk-text{color:#f3f4f6!important}.dk-sub{color:#d1d5db!important}.dk-muted{color:#9ca3af!important}' +
  '.dk-line{border-color:#2a3242!important}.dk-soft{background:#1f2937!important;color:#e5e7eb!important}' +
  '}';

function champ_(lib, val) {
  return '<div style="display:inline-block;vertical-align:top;margin:12px 22px 0 0;max-width:100%;' + BRISE_ + '">' +
    '<div class="dk-muted" style="font-size:11px;line-height:14px;letter-spacing:.6px;text-transform:uppercase;color:#6b7280">' + lib + '</div>' +
    '<div class="dk-text" style="font-size:14px;line-height:20px;font-weight:bold;color:#111827">' + val + '</div></div>';
}

function pastille_(txt, fond, couleur) {
  return '<span style="display:inline-block;background:' + fond + ';color:' + couleur + ';border-radius:2px;padding:3px 8px;font-size:11px;line-height:15px;font-weight:bold;letter-spacing:.7px;text-transform:uppercase;white-space:nowrap">' + txt + '</span>';
}

function bouton_(txt, url, plein) {
  const style = plein
    ? 'background:#128c7e;color:#ffffff;border:1px solid #128c7e;'
    : 'background:#ffffff;color:#374151;border:1px solid #cfd4da;';
  return '<a href="' + url + '" style="display:inline-block;' + style + 'border-radius:2px;padding:9px 16px;margin:12px 8px 0 0;font-size:13px;line-height:18px;font-weight:bold;text-decoration:none">' + txt + '</a>';
}

function lienWhatsapp_(x, mode) {
  const num = telWhatsapp_(x.tel);
  if (!num || !x.dr) return '';
  const texte = mode === 'retard'
    ? 'Bonjour ' + x.client + ', sauf erreur de notre part, la restitution de votre véhicule ' + x.vehicule +
      ' était prévue le ' + dateFr_(x.dr) + '. Merci de nous contacter dès que possible. ' + CONFIG.SOCIETE + '.'
    : 'Bonjour ' + x.client + ', petit rappel : la restitution de votre véhicule ' + x.vehicule +
      ' est prévue le ' + dateFr_(x.dr) + '. Merci, ' + CONFIG.SOCIETE + '.';
  return 'https://wa.me/' + num + '?text=' + encodeURIComponent(texte);
}

function telHtml_(tel) {
  const num = telWhatsapp_(tel);
  return num ? '<a href="tel:+' + num + '" style="color:inherit;text-decoration:none">' + h_(tel) + '</a>' : h_(tel);
}

/** Carte  */
function carte_(o) {
  const droite = o.droite ? '<td valign="top" align="right" style="padding-left:12px;white-space:nowrap;font-size:16px;line-height:22px;font-weight:bold">' + o.droite + '</td>' : '';
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" class="dk-card2" style="margin:0 0 10px;background:#ffffff;border:1px solid #e2e5ea;border-left:3px solid ' + o.couleur + ';border-radius:2px;border-collapse:separate"><tr><td style="padding:14px 14px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td valign="top" style="' + BRISE_ + '">' +
    ((o.pastilles && o.pastilles.length) ? '<div style="margin:0 0 7px">' + o.pastilles.join(' ') + '</div>' : '') +
    '<div class="dk-text" style="font-size:16px;line-height:22px;font-weight:bold;color:#111827">' + o.titre + '</div>' +
    (o.sous ? '<div class="dk-sub" style="font-size:14px;line-height:22px;color:#4b5563">' + o.sous + '</div>' : '') +
    '</td>' + droite + '</tr></table>' +
    ((o.details && o.details.length) ? '<div style="font-size:0;line-height:0">' + o.details.join('') + '</div>' : '') +
    (o.barre !== undefined && o.barre !== null ? '<div class="dk-soft" style="height:4px;background:#eceef1;margin-top:14px"><div style="width:' + o.barre + '%;height:4px;background:#a4161a"></div></div>' : '') +
    (o.note ? '<div class="dk-soft" style="margin-top:12px;padding:9px 12px;background:#f6f7f9;border-left:3px solid #cfd4da;font-size:13px;line-height:19px;color:#4b5563;' + BRISE_ + '"><b>Remarque :</b> ' + h_(o.note.length > 160 ? o.note.slice(0, 157) + '…' : o.note) + '</div>' : '') +
    ((o.actions && o.actions.length) ? '<div>' + o.actions.join('') + '</div>' : '') +
    '</td></tr></table>';
}

/** mode : 'retard' | 'retour' | 'livraison' | 'fait' */
function ligneLocation_(x, mode, kmMax) {
  const couleur = { retard: COULEUR_.rouge, retour: COULEUR_.marine, livraison: COULEUR_.vert, fait: COULEUR_.gris }[mode];
  const pastilles = [];
  if (mode === 'retard') pastilles.push(pastille_(x.retard + ' j de retard', '#b91c1c', '#ffffff'));
  else if (mode === 'fait') pastilles.push(pastille_('Rendu le ' + (x.dr ? dateFr_(x.dr).slice(0, 5) : '—'), '#4b5563', '#ffffff'));
  else pastilles.push(pastille_(x.heure ? h_(x.heure) : 'Heure à confirmer', couleur, '#ffffff'));
  if (mode === 'retour') pastilles.push(pastille_('Retour', '#fbe9e9', '#a4161a'));
  if (mode === 'livraison') pastilles.push(pastille_('Livraison', '#dff3ef', '#0f766e'));

  const details = [];
  if (x.contrat) details.push(champ_('Contrat', h_(x.contrat)));
  if ((mode === 'retour' || mode === 'retard') && x.dr) details.push(champ_('Retour prévu', dateFr_(x.dr)));
  let barre = null;
  if (mode === 'fait') {
    if (x.kmDep !== null && x.kmRet !== null) {
      const p = kmParcourus_(x);
      details.push(champ_('Km départ', fmt_(x.kmDep)));
      details.push(champ_('Km retour', fmt_(x.kmRet)));
      if (p !== null) { details.push(champ_('Parcourus', fmt_(p) + '&nbsp;km')); barre = Math.max(4, Math.round(100 * p / (kmMax || p || 1))); }
      else details.push(champ_('Kilométrage', '<span style="color:#b91c1c">Incohérent (retour inférieur au départ)</span>'));
    } else if (x.kmRet !== null) details.push(champ_('Km retour', fmt_(x.kmRet)));
    else if (x.kmDep !== null) details.push(champ_('Km départ', fmt_(x.kmDep)));
  } else if (x.kmDep !== null) details.push(champ_('Km au départ', fmt_(x.kmDep)));
  if (x.caution && mode !== 'fait') details.push(champ_(mode === 'livraison' ? 'Caution à percevoir' : 'Caution à restituer', h_(fr_(x.caution))));
  if (x.commercial) details.push(champ_('Commercial', h_(x.commercial)));
  if (x.tel && /\d/.test(x.tel)) details.push(champ_('Téléphone', telHtml_(x.tel)));

  const actions = [];
  if (mode === 'retour' || mode === 'retard') {
    const wa = lienWhatsapp_(x, mode), num = telWhatsapp_(x.tel);
    if (wa) actions.push(bouton_('Relancer par WhatsApp', wa, true));
    if (num) actions.push(bouton_('Appeler', 'tel:+' + num, false));
  }
  return carte_({
    couleur: couleur, pastilles: pastilles, titre: h_(x.client),
    sous: h_(x.vehicule) + (x.plaque ? ' &nbsp;<span class="dk-soft" style="display:inline-block;background:#f6f7f9;border:1px solid #e2e5ea;border-radius:2px;padding:0 7px;font-size:12px;line-height:18px;letter-spacing:.6px;color:#374151;white-space:nowrap">' + h_(x.plaque) + '</span>' : ''),
    details: details, barre: barre, note: x.remarque ? fr_(x.remarque) : '', actions: actions
  });
}

function titreJour_(titre, sous) {
  return '<div style="margin:36px 0 0;padding:12px 16px;background:#1f2937;border-radius:2px;color:#ffffff">' +
    '<div style="font-size:15px;line-height:20px;font-weight:bold;letter-spacing:1px;text-transform:uppercase">' + titre + '</div>' +
    '<div style="font-size:13px;line-height:18px;color:#cbd5e1;margin-top:2px">' + sous + '</div></div>';
}

function titreSection_(titre, couleur, nb, note) {
  return '<div class="dk-line dk-text" style="margin:28px 0 ' + (note ? '6' : '12') + 'px;padding:0 0 8px 10px;border-left:3px solid ' + couleur + ';border-bottom:1px solid #cfd4da;font-size:12px;line-height:16px;font-weight:bold;letter-spacing:1.2px;text-transform:uppercase;color:#111827">' + titre +
    (nb !== null && nb !== undefined ? ' <span class="dk-muted" style="font-weight:normal;letter-spacing:0;color:#6b7280">(' + nb + ')</span>' : '') + '</div>' +
    (note ? '<div class="dk-muted" style="margin:0 0 12px;font-size:12px;line-height:18px;color:#6b7280">' + note + '</div>' : '');
}

function bloc_(titre, couleur, liste, fabrique, max, note) {
  if (!liste.length) return '';
  const affiche = max ? liste.slice(0, max) : liste;
  let html = titreSection_(titre, couleur, liste.length, note);
  html += affiche.map(x => fabrique(x)).join('');
  if (affiche.length < liste.length) html += '<div class="dk-muted" style="margin:0 0 8px;font-size:13px;color:#6b7280">… et ' + (liste.length - affiche.length) + ' autre(s) non affiché(s)</div>';
  return html;
}

/** Indicateur chiffré : largeur en %, retour à la ligne naturel (sans media query). */
function tuile_(valeur, libelle, couleur, pct) {
  return '<div style="display:inline-block;width:' + pct + '%;min-width:100px;box-sizing:border-box;vertical-align:top;padding:4px">' +
    '<div class="dk-card2" style="background:#ffffff;border:1px solid #e2e5ea;border-top:3px solid ' + couleur + ';border-radius:2px;padding:14px 6px;text-align:center">' +
    '<div class="dk-text" style="font-size:28px;line-height:32px;font-weight:bold;color:#111827">' + valeur + '</div>' +
    '<div class="dk-muted" style="margin-top:5px;font-size:11px;line-height:14px;letter-spacing:.4px;text-transform:uppercase;color:#6b7280;' + BRISE_ + '">' + libelle + '</div></div></div>';
}

function bilanHtml_(c) {
  const L = bilanLignes_(c);
  if (!L.length) return '';
  let html = titreSection_('Bilan financier', COULEUR_.violet, null, MOIS_FR_[c.aujourd.getUTCMonth()] + ' ' + c.aujourd.getUTCFullYear() + ' · montants en dinars (DA)');
  html += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" class="dk-card2" style="background:#ffffff;border:1px solid #e2e5ea;border-collapse:separate;border-radius:2px">';
  let groupe = '';
  L.forEach(x => {
    if (x.g !== groupe) {
      groupe = x.g;
      html += '<tr><td colspan="2" class="dk-soft dk-line" style="padding:9px 16px;background:#f6f7f9;border-bottom:1px solid #e2e5ea;font-size:11px;line-height:14px;letter-spacing:1px;text-transform:uppercase;font-weight:bold;color:#4b5563">' + h_(x.g) + '</td></tr>';
    }
    html += '<tr><td class="dk-line" valign="top" style="padding:11px 16px;border-bottom:1px solid #eef0f2;font-size:14px;line-height:20px;' + BRISE_ + '">' +
      '<span class="dk-text" style="color:#111827">' + h_(x.l) + '</span>' +
      (x.d ? '<div class="dk-muted" style="font-size:12px;line-height:17px;color:#6b7280">' + h_(x.d) + '</div>' : '') + '</td>' +
      '<td class="dk-line" valign="top" align="right" style="padding:11px 16px;border-bottom:1px solid #eef0f2;white-space:nowrap;font-size:15px;line-height:20px;font-weight:bold;color:' + (x.c || '#111827') + '">' + fmt_(x.v) + '&nbsp;DA</td></tr>';
  });
  return html + '</table>';
}

function htmlVersTexte_(html) {
  return html
    .replace(/<head[\s\S]*?<\/head>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<div style="display:none[^>]*>[\s\S]*?<\/div>/gi, '')
    .replace(/<\/(tr|h\d|p|div|li)>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/td>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
    .replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

/* CONSTRUCTION DE L'E-MAIL  */

function construireEtEnvoyer_(forcer) {
  if (/exemple\d*@/i.test(CONFIG.DESTINATAIRES)) {
    throw new Error('Remplace d’abord les adresses « exemple@gmail.com » par celles de ton équipe (CONFIG.DESTINATAIRES).');
  }
  const c = collecte_(forcer);
  const aujourd = c.aujourd, demain = c.demain, info = c.info;
  const R = reglages_().R;
  const e = Object.assign({}, c.r);
  if (!R.PROGRAMME) { e.retAuj = []; e.retDem = []; e.livAuj = []; e.livDem = []; }
  if (!R.RETARDS) e.retard = [];
  if (!R.RETOURS_ENREGISTRES) e.retFaits = [];
  if (!R.DETTES) e.dettes = [];
  if (!R.VIDANGES) e.vidanges = [];
  if (!R.DOCUMENTS) e.docs = [];
  if (!R.LAVAGES) e.lavages = [];
  if (!R.ENTRETIEN) e.attente = [];
  if (!R.RESERVATIONS) e.reservations = [];

  const total = Object.keys(e).reduce((n, k) => n + e[k].length, 0) + (R.BILAN ? 1 : 0);
  if (!total && !forcer && !CONFIG.ENVOYER_SI_VIDE) return 0;

  const nRetAuj = e.retAuj.length, nLivAuj = e.livAuj.length, nRetDem = e.retDem.length, nLivDem = e.livDem.length, nRetard = e.retard.length;
  const totalDettes = e.dettes.reduce((s, x) => s + x.dette, 0);
  const kmParc = e.retFaits.map(x => kmParcourus_(x) || 0);
  const kmTotal = kmParc.reduce((s, v) => s + v, 0), kmMax = Math.max.apply(null, kmParc.concat([1]));
  const url = urlDashboard_();

  const resume = R.PROGRAMME
    ? 'Aujourd’hui : ' + pl_(nRetAuj, 'retour', 'retours') + ', ' + pl_(nLivAuj, 'livraison', 'livraisons') + ' · Demain : ' + pl_(nRetDem, 'retour', 'retours') + ', ' + pl_(nLivDem, 'livraison', 'livraisons')
    : 'Récapitulatif de la flotte du ' + dateFr_(aujourd);

  // ---- Document + en-tête
  let html = '<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<meta name="color-scheme" content="light dark"><meta name="supported-color-schemes" content="light dark">' +
    '<title>Récapitulatif quotidien de la flotte</title><style>' + CSS_EMAIL_ + '</style></head>' +
    '<body class="dk-page" style="margin:0;padding:0;background:#f1f2f4">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:transparent;font-size:1px;line-height:1px">' + h_(resume) + '&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" class="dk-page" style="background:#f1f2f4"><tr><td align="center" class="outer" style="padding:16px 6px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" class="dk-card" style="width:100%;max-width:680px;background:#ffffff;border:1px solid #e2e5ea;border-collapse:separate;font-family:' + POLICE_ + ';color:#111827">' +
    '<tr><td class="px" style="background:#a4161a;padding:28px 20px 26px;color:#ffffff">' +
    '<div style="font-size:12px;line-height:16px;letter-spacing:2.5px;text-transform:uppercase;font-weight:bold;color:#f6cfcf">' + h_(CONFIG.SOCIETE) + '</div>' +
    '<div class="h1" style="font-size:25px;line-height:31px;font-weight:bold;margin-top:10px;color:#ffffff">Récapitulatif quotidien de la flotte</div>' +
    '<div style="font-size:14px;line-height:20px;margin-top:6px;color:#f6cfcf">' + h_(maj_(dateLongue_(aujourd))) + '</div>' +
    (nRetard ? '<div style="margin-top:16px"><span style="display:inline-block;background:#ffffff;color:#a4161a;border-radius:2px;padding:6px 12px;font-size:12px;line-height:16px;font-weight:bold;letter-spacing:.6px;text-transform:uppercase">' + pl_(nRetard, 'retour en retard à relancer', 'retours en retard à relancer') + '</span></div>' : '') +
    '</td></tr>' +
    '<tr><td class="px dk-text" style="padding:26px 20px 30px;font-size:15px;line-height:24px;color:#111827">' +
    '<p style="margin:0 0 6px">Bonjour à toute l’équipe,</p>' +
    '<p style="margin:0">Veuillez trouver ci-dessous le récapitulatif de la flotte' + (R.PROGRAMME ? ' pour <b>aujourd’hui</b> et <b>demain</b>' : '') + '.</p>';

  // ---- Indicateurs
  const tuiles = [];
  if (R.PROGRAMME) tuiles.push(['Retours aujourd’hui', nRetAuj, COULEUR_.marine], ['Livraisons aujourd’hui', nLivAuj, COULEUR_.vert], ['Retours demain', nRetDem, COULEUR_.marine], ['Livraisons demain', nLivDem, COULEUR_.vert]);
  if (R.RETARDS) tuiles.push(['Retours en retard', nRetard, nRetard ? COULEUR_.rouge : '#9ca3af'], ['Véhicules loués', info.enCours, COULEUR_.violet]);
  if (tuiles.length) {
    const pct = tuiles.length <= 4 ? 100 / tuiles.length : 33.33;
    html += '<div style="margin:20px -4px 4px;font-size:0;line-height:0">' + tuiles.map(t => tuile_(t[1], t[0], t[2], pct)).join('') + '</div>';
  }

  // ---- Sections
  html += bloc_('Retours en retard', COULEUR_.rouge, e.retard, x => ligneLocation_(x, 'retard'));
  const jour_ = (titre, d, retours, livraisons) => {
    const entete = titreJour_(titre, h_(dateLongue_(d)));
    if (!retours.length && !livraisons.length) {
      return entete + '<p class="dk-muted" style="margin:14px 0 0;font-size:14px;font-style:italic;color:#6b7280">Aucun retour ni aucune livraison prévu.</p>';
    }
    return entete + bloc_('Retours', COULEUR_.marine, retours, x => ligneLocation_(x, 'retour')) +
      bloc_('Livraisons', COULEUR_.vert, livraisons, x => ligneLocation_(x, 'livraison'));
  };
  if (R.PROGRAMME) { html += jour_('Aujourd’hui', aujourd, e.retAuj, e.livAuj); html += jour_('Demain', demain, e.retDem, e.livDem); }
  html += bloc_('Réservations à venir', COULEUR_.violet, e.reservations, x => ligneReservation_(x), 0,
    'Départs des ' + (reglages_().joursReservations === 1 ? 'aujourd’hui et demain' : reglages_().joursReservations + ' prochains jours'));

  html += bloc_('Retours enregistrés', COULEUR_.gris, e.retFaits, x => ligneLocation_(x, 'fait', kmMax), 0,
    'Restitution du ' + dateFr_(decale_(aujourd, -CONFIG.JOURS_RETOURS_RECENTS)) + ' au ' + dateFr_(aujourd) + (kmTotal ? ' · total parcouru : <b>' + fmt_(kmTotal) + ' km</b>' : ''));

  html += bloc_('Dettes clients', COULEUR_.orange, e.dettes, x => carte_({
    couleur: COULEUR_.orange, titre: h_(x.client),
    sous: h_(x.vehicule) + (x.contrat ? ' · ' + h_(x.contrat) : ''),
    details: [x.tel && /\d/.test(x.tel) ? champ_('Téléphone', telHtml_(x.tel)) : '', x.commercial ? champ_('Commercial', h_(x.commercial)) : ''].filter(Boolean),
    note: x.remarque ? fr_(x.remarque) : '',
    droite: '<span style="color:' + COULEUR_.orange + '">' + (x.estimee ? '≈&nbsp;' : '') + fmt_(x.dette) + '&nbsp;DA</span>'
  }), CONFIG.MAX_LIGNES_DETTES,
    'Total : <b>' + fmt_(totalDettes) + ' DA</b>' + (e.dettes.some(x => x.estimee) ? ' · ≈ = estimation d’après le total HT (paiement non confirmé)' : ''));

  html += bloc_('Vidanges à prévoir', COULEUR_.orange, e.vidanges, x => carte_({
    couleur: x.reste <= 0 ? COULEUR_.rouge : COULEUR_.orange,
    pastilles: [x.reste <= 0 ? pastille_('Dépassée de ' + fmt_(-x.reste) + ' km', '#b91c1c', '#ffffff') : pastille_('Dans ' + fmt_(x.reste) + ' km', '#fdf0e1', '#92400e')],
    titre: h_(x.modele), sous: h_(x.plaque), details: [champ_('Km actuel', fmt_(x.km) + '&nbsp;km')]
  }));

  html += bloc_('Assurance et contrôle technique', COULEUR_.orange, e.docs, x => carte_({
    couleur: x.j <= 0 ? COULEUR_.rouge : COULEUR_.orange,
    pastilles: [x.j < 0 ? pastille_('Expiré depuis ' + (-x.j) + ' j', '#b91c1c', '#ffffff') : x.j === 0 ? pastille_('Expire aujourd’hui', '#b91c1c', '#ffffff') : pastille_('Expire dans ' + x.j + ' j', '#fdf0e1', '#92400e')],
    titre: h_(x.nom), sous: h_(x.modele) + ' · ' + h_(x.plaque), details: [champ_('Expiration', dateFr_(x.d))]
  }));

  const lavTotal = e.lavages.reduce((t, x) => t + x.prix, 0), lavNP = info.lavNonPaye;
  html += bloc_('Lavages', COULEUR_.vert, e.lavages, x => carte_({
    couleur: COULEUR_.vert,
    pastilles: [pastille_(dateFr_(x.date).slice(0, 5), '#4b5563', '#ffffff')].concat(
      x.statut === 'oui' ? [pastille_('Payé au laveur', '#dff3ef', '#0f766e')] : x.statut === 'non' ? [pastille_('À payer au laveur', '#fbe9e9', '#a4161a')] : []),
    titre: h_(x.plaque || x.type || 'Lavage'),
    sous: h_([x.type, x.laveur].filter(Boolean).join(' · ')),
    droite: x.prix ? fmt_(x.prix) + '&nbsp;DA' : ''
  }), 0,
    'Total : <b>' + fmt_(lavTotal) + ' DA</b>' + (lavNP.nb ? ' · reste à payer aux laveurs : <b>' + fmt_(lavNP.total) + ' DA</b> (' + pl_(lavNP.nb, 'lavage', 'lavages') + ')' : ''));

  html += bloc_('Réparations et achats en attente', COULEUR_.violet, e.attente, x => carte_({
    couleur: COULEUR_.violet, titre: h_(x.type || 'Intervention'), sous: h_(x.plaque),
    details: [x.desc ? champ_('Détail', h_(x.desc)) : '', x.date ? champ_('Demandé le', dateFr_(x.date)) : ''].filter(Boolean),
    droite: x.prix ? fmt_(x.prix) + '&nbsp;DA' : ''
  }));

  if (R.BILAN) html += bilanHtml_(c);

  // ---- Lien tableau de bord + signature + pied de page
  if (url) {
    html += '<div style="margin:30px 0 0"><a href="' + url + '" style="display:inline-block;background:#1f2937;color:#ffffff;border-radius:2px;padding:12px 24px;font-size:13px;font-weight:bold;letter-spacing:.4px;text-decoration:none">Ouvrir le tableau de bord</a></div>';
  }
  html += '<p style="margin:32px 0 0;font-size:15px;line-height:24px">Bonne journée à toute l’équipe.</p>' +
    '<p style="margin:16px 0 0;font-size:15px;line-height:24px">Cordialement,<br><b>Maini Lotfi</b></p>' +
    '</td></tr>' +
    '<tr><td class="px dk-soft dk-muted" style="padding:16px 20px;background:#f6f7f9;border-top:1px solid #e2e5ea;font-size:12px;line-height:18px;color:#6b7280;' + BRISE_ + '">E-mail automatique envoyé par ' + h_(CONFIG.SIGNATURE) + '.</td></tr>' +
    '</table></td></tr></table></body></html>';

  const sujet = ' Récapitulatif du ' + dateFr_(aujourd) +
    (R.PROGRAMME ? ' · ' + pl_(nRetAuj, 'retour', 'retours') + ', ' + pl_(nLivAuj, 'livraison', 'livraisons') : '') + (nRetard ? ' · ' + nRetard + ' en retard' : '');

  MailApp.sendEmail({ to: CONFIG.DESTINATAIRES, subject: sujet, htmlBody: html, body: htmlVersTexte_(html), name: 'Flotte ' + CONFIG.SOCIETE });
  if (total) envoyerTelegram_(texteTelegram_(c));
  return total;
}

/* TABLEAU DE BORD WEB */

function doGet() {
  return HtmlService.createHtmlOutput(htmlDashboard_(collecte_(true)))
    .setTitle('Flotte ' + CONFIG.SOCIETE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function itemDash_(o, i) {
  return '<div class="item" style="--i:' + i + ';--c:' + o.couleur + '">' +
    (o.gauche ? '<div class="time">' + o.gauche + '</div>' : '') +
    '<div class="body"><div class="nm">' + o.titre + '</div>' + (o.sous ? '<div class="sub">' + o.sous + '</div>' : '') +
    (o.meta && o.meta.length ? '<div class="meta">' + o.meta.map(m => '<span>' + m + '</span>').join('') + '</div>' : '') +
    (o.barre != null ? '<div class="bar"><i style="width:' + o.barre + '%"></i></div>' : '') +
    (o.note ? '<div class="note"><b>Remarque :</b> ' + h_(o.note) + '</div>' : '') +
    (o.actions && o.actions.length ? '<div class="act">' + o.actions.join('') + '</div>' : '') + '</div>' +
    (o.droite ? '<div class="right" style="color:' + (o.couleurDroite || '#92400e') + '">' + o.droite + '</div>' : '') + '</div>';
}

function locDash_(x, mode, i) {
  const couleur = { retard: '#b91c1c', retour: '#a4161a', livraison: '#0f766e', fait: '#6b7280' }[mode];
  let g = x.heure ? h_(x.heure) : '—';
  if (mode === 'retard') g = x.retard + ' j'; else if (mode === 'fait') g = x.dr ? dateFr_(x.dr).slice(0, 5) : '—';
  const meta = [];
  if (x.contrat) meta.push('Contrat <b>' + h_(x.contrat) + '</b>');
  if ((mode === 'retour' || mode === 'retard') && x.dr) meta.push('Retour prévu <b>' + dateFr_(x.dr) + '</b>');
  let barre = null;
  if (mode === 'fait' && x.kmDep !== null && x.kmRet !== null) {
    const p = kmParcourus_(x);
    meta.push('Km départ <b>' + fmt_(x.kmDep) + '</b> · retour <b>' + fmt_(x.kmRet) + '</b>');
    if (p !== null) { meta.push('Parcourus <b>' + fmt_(p) + ' km</b>'); barre = Math.min(100, Math.max(4, p / 5)); }
  } else if (x.kmDep !== null && mode !== 'fait') meta.push('Km au départ <b>' + fmt_(x.kmDep) + '</b>');
  if (x.caution && mode !== 'fait') meta.push('Caution <b>' + h_(fr_(x.caution)) + '</b>');
  if (x.commercial) meta.push('Commercial <b>' + h_(x.commercial) + '</b>');
  if (x.tel && /\d/.test(x.tel)) meta.push('Tél. <b>' + h_(x.tel) + '</b>');
  const actions = [];
  if (mode === 'retour' || mode === 'retard') {
    const wa = lienWhatsapp_(x, mode), num = telWhatsapp_(x.tel);
    if (wa) actions.push('<a target="_top" href="' + wa + '">Relancer par WhatsApp</a>');
    if (num) actions.push('<a class="o" target="_top" href="tel:+' + num + '">Appeler</a>');
  }
  return itemDash_({ gauche: g, couleur: couleur, titre: h_(x.client),
    sous: h_(x.vehicule) + (x.plaque ? ' · ' + h_(x.plaque) : ''), meta: meta, barre: barre, note: x.remarque ? fr_(x.remarque) : '', actions: actions }, i);
}

function secDash_(cle, titre, couleur, liste, fab) {
  if (!liste.length) return '';
  return '<section data-k="' + cle + '"><h2 style="border-color:' + couleur + '">' + titre + '<span class="cnt">' + liste.length + '</span></h2>' +
    liste.map((x, i) => fab(x, i)).join('') + '</section>';
}

function htmlDashboard_(c) {
  const r = c.r, info = c.info;
  const totalDettes = r.dettes.reduce((s, x) => s + x.dette, 0);
  const hhmm = Utilities.formatDate(new Date(), CONFIG.FUSEAU, 'HH:mm').replace(':', 'h');
  const kpi = (n, lib, i, col) => '<div class="kpi" style="--i:' + i + ';--c:' + col + '"><b data-n="' + Math.round(n) + '">0</b><span>' + lib + '</span></div>';

  const css = `
:root{--bg:#f1f2f4;--card:#ffffff;--line:#e2e5ea;--txt:#111827;--sub:#4b5563;--mut:#6b7280;--brand:#a4161a}
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:var(--txt);background:var(--bg)}
.top{background:var(--brand);color:#fff}.in{max-width:980px;margin:0 auto;padding:26px 16px}
.co{font-size:12px;letter-spacing:2.5px;text-transform:uppercase;font-weight:700;color:#f6cfcf}
h1{margin:8px 0 4px;font-size:clamp(22px,5vw,30px);line-height:1.2}
.date{color:#f6cfcf;font-size:14px}
.wrap{max-width:980px;margin:0 auto;padding:8px 16px 60px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin:22px 0 8px}
.kpi{background:var(--card);border:1px solid var(--line);border-top:3px solid var(--c);border-radius:2px;padding:16px;animation:fade .5s both;animation-delay:calc(var(--i)*.05s)}
.kpi b{display:block;font-size:30px;line-height:1.1}
.kpi span{font-size:11px;letter-spacing:1px;text-transform:uppercase;color:var(--mut)}
.chips{display:flex;gap:8px;flex-wrap:wrap;margin:18px 0 0}
.chip{cursor:pointer;border:1px solid #cfd4da;background:#fff;color:var(--txt);border-radius:2px;padding:9px 18px;font-size:13px;font-weight:600}
.chip:hover{background:#f6f7f9}.chip.on{background:#1f2937;border-color:#1f2937;color:#fff}
section{margin-top:28px}
section h2{display:flex;align-items:center;gap:10px;margin:0 0 12px;padding:0 0 8px 10px;border-left:3px solid var(--brand);border-bottom:1px solid #cfd4da;font-size:12px;letter-spacing:1.2px;text-transform:uppercase}
.cnt{color:var(--mut);font-weight:400;letter-spacing:0}.cnt:before{content:"("}.cnt:after{content:")"}
.item{display:flex;gap:14px;background:var(--card);border:1px solid var(--line);border-left:3px solid var(--c);border-radius:2px;padding:14px 16px;margin-bottom:10px;animation:fade .5s both;animation-delay:calc(var(--i)*.04s)}
.time{min-width:62px;height:fit-content;text-align:center;font-weight:700;font-size:15px;border-radius:2px;padding:8px 6px;background:var(--c);color:#fff}
.body{flex:1;min-width:0;overflow-wrap:anywhere}
.nm{font-weight:700;font-size:16px}.sub{color:var(--sub);font-size:14px;margin-top:1px}
.meta{margin-top:8px;font-size:13px;color:var(--mut);display:flex;flex-wrap:wrap;gap:4px 18px}.meta b{color:var(--txt)}
.note{margin-top:9px;padding:8px 12px;background:#f6f7f9;border-left:3px solid #cfd4da;font-size:13px;color:var(--sub)}
.bar{height:4px;background:#eceef1;margin-top:12px}.bar i{display:block;height:100%;background:var(--brand)}
.act a{display:inline-block;margin:10px 8px 0 0;padding:8px 14px;border-radius:2px;font-size:13px;font-weight:700;text-decoration:none;color:#fff;background:#128c7e}
.act a.o{background:#fff;color:#374151;border:1px solid #cfd4da}
.right{font-weight:700;font-size:16px;white-space:nowrap;align-self:flex-start}
.empty{color:var(--mut);font-style:italic;margin-top:28px}
.hide{display:none!important}
@keyframes fade{from{opacity:0}to{opacity:1}}
@media(max-width:560px){.item{flex-wrap:wrap}.right{width:100%}}
@media(prefers-reduced-motion:reduce){*{animation:none!important}}`;

  const js = `
document.querySelectorAll('[data-n]').forEach(function(el){var n=+el.dataset.n,v=0,s=Math.max(1,Math.ceil(n/25));
 if(!n){el.textContent=0;return}var t=setInterval(function(){v=Math.min(n,v+s);el.textContent=v.toLocaleString('fr-FR');if(v>=n)clearInterval(t)},30)});
document.querySelectorAll('.chip').forEach(function(b){b.onclick=function(){
 document.querySelectorAll('.chip').forEach(function(x){x.classList.remove('on')});b.classList.add('on');var k=b.dataset.k;
 document.querySelectorAll('section').forEach(function(s){s.classList.toggle('hide',k!=='*'&&(' '+s.dataset.k+' ').indexOf(' '+k+' ')<0)})}});`;

  let body = '<div class="top"><div class="in"><div class="co">' + h_(CONFIG.SOCIETE) + '</div><h1>Tableau de bord de la flotte</h1>' +
    '<div class="date">' + h_(maj_(dateLongue_(c.aujourd))) + ' · mis à jour à ' + hhmm + '</div></div></div><div class="wrap">' +
    '<div class="kpis">' +
    kpi(r.retAuj.length, 'Retours aujourd’hui', 0, '#a4161a') + kpi(r.livAuj.length, 'Livraisons aujourd’hui', 1, '#0f766e') +
    kpi(r.retDem.length, 'Retours demain', 2, '#a4161a') + kpi(r.livDem.length, 'Livraisons demain', 3, '#0f766e') +
    kpi(r.retard.length, 'En retard', 4, '#b91c1c') + kpi(info.enCours, 'Véhicules loués', 5, '#475569') +
    (totalDettes ? kpi(totalDettes, 'Dettes (DA)', 6, '#92400e') : '') + '</div>' +
    '<div class="chips"><button class="chip on" data-k="*">Tout</button><button class="chip" data-k="r">Retours</button><button class="chip" data-k="l">Livraisons</button><button class="chip" data-k="a">Alertes</button><button class="chip" data-k="b">Bilan</button></div>';

  body += secDash_('r a', 'Retours en retard', '#b91c1c', r.retard, (x, i) => locDash_(x, 'retard', i));
  body += secDash_('r', 'Retours aujourd’hui', '#a4161a', r.retAuj, (x, i) => locDash_(x, 'retour', i));
  body += secDash_('l', 'Livraisons aujourd’hui', '#0f766e', r.livAuj, (x, i) => locDash_(x, 'livraison', i));
  body += secDash_('r', 'Retours demain', '#a4161a', r.retDem, (x, i) => locDash_(x, 'retour', i));
  body += secDash_('l', 'Livraisons demain', '#0f766e', r.livDem, (x, i) => locDash_(x, 'livraison', i));
  body += secDash_('l', 'Réservations à venir', '#475569', r.reservations, (x, i) => itemDash_({
    gauche: x.dans === 0 ? 'Auj.' : x.dans === 1 ? 'Demain' : dateFr_(x.du).slice(0, 5), couleur: x.dans <= 1 ? '#a4161a' : '#475569',
    titre: h_(x.client), sous: h_(x.vehicule) + (x.heure ? ' · ' + h_(x.heure) : ''),
    meta: ['Période <b>' + dateFr_(x.du) + (x.au ? ' au ' + dateFr_(x.au) : '') + '</b>', x.lieu ? 'Livraison <b>' + h_(x.lieu) + '</b>' : '', x.commercial ? 'Commercial <b>' + h_(x.commercial) + '</b>' : '',
      x.tel && /\d/.test(x.tel) ? 'Tél. <b>' + h_(x.tel) + '</b>' : '', x.total ? 'Total HT <b>' + fmt_(x.total) + ' DA</b>' : ''].filter(Boolean),
    note: x.remarque ? fr_(x.remarque) : '' }, i));
  body += secDash_('r', 'Retours enregistrés', '#6b7280', r.retFaits, (x, i) => locDash_(x, 'fait', i));
  body += secDash_('a', 'Dettes clients', '#92400e', r.dettes.slice(0, CONFIG.MAX_LIGNES_DETTES), (x, i) => itemDash_({
    couleur: '#92400e', titre: h_(x.client), sous: h_(x.vehicule) + (x.contrat ? ' · ' + h_(x.contrat) : ''),
    meta: [x.tel && /\d/.test(x.tel) ? 'Tél. <b>' + h_(x.tel) + '</b>' : '', x.commercial ? 'Commercial <b>' + h_(x.commercial) + '</b>' : ''].filter(Boolean),
    note: x.remarque ? fr_(x.remarque) : '', droite: (x.estimee ? '≈ ' : '') + fmt_(x.dette) + ' DA' }, i));
  body += secDash_('a', 'Vidanges à prévoir', '#92400e', r.vidanges, (x, i) => itemDash_({
    couleur: x.reste <= 0 ? '#b91c1c' : '#92400e', titre: h_(x.modele), sous: h_(x.plaque), meta: ['Km actuel <b>' + fmt_(x.km) + ' km</b>'],
    droite: x.reste <= 0 ? 'Dépassée de ' + fmt_(-x.reste) + ' km' : 'Dans ' + fmt_(x.reste) + ' km', couleurDroite: x.reste <= 0 ? '#b91c1c' : '#92400e' }, i));
  body += secDash_('a', 'Assurance et contrôle technique', '#92400e', r.docs, (x, i) => itemDash_({
    couleur: x.j <= 0 ? '#b91c1c' : '#92400e', titre: h_(x.nom), sous: h_(x.modele) + ' · ' + h_(x.plaque), meta: ['Expiration <b>' + dateFr_(x.d) + '</b>'],
    droite: x.j < 0 ? 'Expiré depuis ' + (-x.j) + ' j' : x.j === 0 ? 'Expire aujourd’hui' : 'Dans ' + x.j + ' j', couleurDroite: x.j <= 0 ? '#b91c1c' : '#92400e' }, i));
  body += secDash_('a', 'Lavages', '#0f766e', r.lavages, (x, i) => itemDash_({
    couleur: '#0f766e', titre: h_(x.plaque || x.type || 'Lavage'), sous: h_([x.type, x.laveur].filter(Boolean).join(' · ')),
    meta: [x.statut === 'oui' ? 'Payé au laveur' : x.statut === 'non' ? 'À payer au laveur' : ''].filter(Boolean),
    droite: x.prix ? fmt_(x.prix) + ' DA' : '', couleurDroite: '#0f766e' }, i));
  body += secDash_('a', 'Réparations et achats en attente', '#475569', r.attente, (x, i) => itemDash_({
    couleur: '#475569', titre: h_(x.type || 'Intervention'), sous: h_(x.plaque),
    meta: [x.desc ? 'Détail <b>' + h_(x.desc) + '</b>' : '', x.date ? 'Demandé le <b>' + dateFr_(x.date) + '</b>' : ''].filter(Boolean), droite: x.prix ? fmt_(x.prix) + ' DA' : '', couleurDroite: '#475569' }, i));
  body += secDash_('b', 'Bilan financier', '#475569', bilanLignes_(c), (x, i) => itemDash_({
    couleur: x.c || '#475569', titre: h_(x.l), sous: h_(x.g + ' · ' + x.d), droite: fmt_(x.v) + ' DA', couleurDroite: x.c || '#111827' }, i));

  if (!r.retAuj.length && !r.livAuj.length && !r.retDem.length && !r.livDem.length && !r.retard.length) {
    body += '<p class="empty">Aucun retour ni aucune livraison prévu aujourd’hui ou demain.</p>';
  }
  body += '</div>';
  return '<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><style>' + css + '</style></head><body>' + body + '<script>' + js + '<\/script></body></html>';
}

/* ÉTAT DE LOCATION : titres, calculs automatiques, couleurs  */

const MARQUE_REGLE_ = '#IAM';   // marqueur caché dans nos règles de couleur : permet de les remplacer sans toucher aux tiennes

// titres dans l'ordre des colonnes (comme ton ancienne feuille). Seules les cases de titre VIDES sont remplies.
const TITRES_ETAT_ = [
  ['CLIENT', 'NOM'], ['VEHICULE', 'TYPE DE VEHICULE'], ['PLAQUE', 'MAT'],
  ['KM_LIVRAISON', 'KM LIVRAISON'], ['KM_RESTITUTION', 'KM RESTITUTION'], ['CONSOMMATION', 'CONSOMMATION KM'],
  ['DATE_LIVRAISON', 'DATE DE LIVRAISON'], ['DATE_RESTITUTION', 'DATE DE RESTITUTION'], ['HEURE_LIVRAISON', 'HEURE DE LIVRAISON'],
  ['CONTRAT', 'N° DE CONTRAT'], ['PRIX_UNITAIRE', 'PRIX UNITAIRE'], ['VALEUR', 'VALEUR (NB JOURS)'],
  ['TOTAL_HT', 'TOTAL HT'], ['TOTAL_TTC', 'TOTAL TTC'], ['COMMERCIAL', 'COMMERCIAL'], ['TELEPHONE', 'NUM° TLF'],
  ['CAUTION', 'CAUTION'], ['REMARQUE', 'REMARQUE'], ['PAYE', 'PAYE'], ['RESTE', 'RESTE'], ['DETTE', 'DETTE']
];

const NOTES_ETAT_ = {
  CONSOMMATION: 'Calcul automatique : KM RESTITUTION - KM LIVRAISON.',
  DATE_RESTITUTION: 'Couleur automatique : VERT = date pas encore atteinte · ORANGE = date d’aujourd’hui · ROUGE = date dépassée.',
  VALEUR: 'Calcul automatique : DATE DE RESTITUTION - DATE DE LIVRAISON (nombre de jours).',
  TOTAL_HT: 'Calcul automatique : PRIX UNITAIRE x nombre de jours.',
  TOTAL_TTC: 'Calcul automatique : TOTAL HT + TVA.',
  PAYE: 'À saisir : montant payé en DA (il est déduit du TOTAL HT). Un texte comme « PAYED ALL THE AMMOUNT » est compris comme « payé en totalité ».',
  RESTE: 'Calcul automatique : TOTAL HT - PAYE.',
  DETTE: 'Calcul automatique : RESTE, une fois la voiture rendue ou la date de restitution dépassée.'
};

/** Formules de la ligne r (références relatives : elles se recopient vers le bas). Le reste à payer se calcule sur le TOTAL HT. */
function formulesEtat_() {
  const c = CONFIG.COL;
  const R = (k, r) => '$' + c[k] + r;
  const tva = 1 + CONFIG.TVA;
  const brut = {
    CONSOMMATION: r => '=IF(AND(ISNUMBER(' + R('KM_LIVRAISON', r) + '),ISNUMBER(' + R('KM_RESTITUTION', r) + ')),' + R('KM_RESTITUTION', r) + '-' + R('KM_LIVRAISON', r) + ',"")',
    VALEUR: r => '=IF(AND(ISNUMBER(' + R('DATE_LIVRAISON', r) + '),ISNUMBER(' + R('DATE_RESTITUTION', r) + ')),MAX(0,' + R('DATE_RESTITUTION', r) + '-' + R('DATE_LIVRAISON', r) + '),"")',
    TOTAL_HT: r => '=IF(AND(ISNUMBER(' + R('PRIX_UNITAIRE', r) + '),ISNUMBER(' + R('VALEUR', r) + ')),' + R('PRIX_UNITAIRE', r) + '*' + R('VALEUR', r) + ',"")',
    TOTAL_TTC: r => '=IF(ISNUMBER(' + R('TOTAL_HT', r) + '),ROUND(' + R('TOTAL_HT', r) + '*' + tva + ',2),"")',
    // RESTE = TOTAL HT - montant payé. PAYE accepte un nombre ou un texte (« 99.000 DA », « PAYED ALL THE AMMOUNT », « DID NOT PAY YET », « STILL 270000 »…)
    RESTE: r => {
      const M = R('TOTAL_HT', r), N = R('TOTAL_TTC', r), S = R('PAYE', r);
      return '=LET(base,IF(ISNUMBER(' + M + '),' + M + ',IF(ISNUMBER(' + N + '),ROUND(' + N + '/' + tva + ',2),"")),' +
        'IF(NOT(ISNUMBER(base)),"",LET(p,' + S + '&"",' +
        'ch,IFERROR(VALUE(REGEXREPLACE(REGEXREPLACE(p,"[.,]\\d{1,2}\\b",""),"[^0-9]","")),0),' +
        'pay,IF(ISNUMBER(' + S + '),' + S + ',IF(AND(ch>0,REGEXMATCH(LOWER(p),"still|reste")),base-ch,IF(ch>0,ch,' +
        'IF(REGEXMATCH(LOWER(p),"not pay|non pay|impay|pas pay|yet"),0,IF(REGEXMATCH(LOWER(p),"pay|pai|vers|total|full|complet"),base,0))))),' +
        'MAX(0,base-pay))))';
    },
    // DETTE = RESTE, mais seulement quand la location est terminée (voiture rendue, ou date de restitution dépassée)
    DETTE: r => '=IF(NOT(ISNUMBER(' + R('RESTE', r) + ')),"",IF(AND(' + R('RESTE', r) + '>0,OR(ISNUMBER(' + R('KM_RESTITUTION', r) + '),AND(ISNUMBER(' + R('DATE_RESTITUTION', r) + '),' +
      R('DATE_RESTITUTION', r) + '<TODAY()))),' + R('RESTE', r) + ',0))'
  };
  // chaque formule est protégée par IFERROR : jamais de message d'erreur dans la feuille
  const sure = {};
  Object.keys(brut).forEach(k => { sure[k] = r => '=IFERROR(' + brut[k](r).slice(1) + ',"")'; });
  return sure;
}

const ORDRE_CALCULS_ETAT_ = ['CONSOMMATION', 'VALEUR', 'TOTAL_HT', 'TOTAL_TTC', 'RESTE', 'DETTE'];
const FORMATS_CALCULS_ETAT_ = { CONSOMMATION: '#,##0', VALEUR: '0', TOTAL_HT: '#,##0.00', TOTAL_TTC: '#,##0.00', RESTE: '#,##0.00', DETTE: '#,##0.00' };

/** Pose la formule sur chaque série de lignes ciblées (une seule requête par série). Retourne le nombre de cases remplies. */
function remplirSerieEtat_(f, col, premiere, cibles, modele, format) {
  let i = 0, nb = 0;
  while (i < cibles.length) {
    if (!cibles[i]) { i++; continue; }
    let j = i;
    while (j + 1 < cibles.length && cibles[j + 1]) j++;
    const rng = f.getRange(premiere + i, col + 1, j - i + 1, 1);
    if (format) rng.setNumberFormat(format);   // d'abord un vrai format (pas « texte »), ensuite la formule
    rng.setFormula(modele(premiere + i));
    nb += j - i + 1;
    i = j + 1;
  }
  return nb;
}

/**
 * Couleur automatique de la case « date de restitution » (colonnes de CONFIG.COULEUR_ETAT) :
 *   ORANGE = date de restitution = aujourd'hui
 *   ROUGE  = date de restitution dépassée
 *   VERT   = date de restitution pas encore atteinte
 * + DETTE en rouge quand elle est positive. Nos anciennes règles sont remplacées, les tiennes ne sont pas touchées.
 */
function appliquerCouleursEtat_(f, debut, dernier) {
  const c = CONFIG.COL, nb = dernier - debut + 1;
  const H = '$' + c.DATE_RESTITUTION + debut, U = '$' + c.DETTE + debut;
  const formule = s => '=AND(' + s + ',LEN("' + MARQUE_REGLE_ + '")>0)';
  const aujour = formule('ISNUMBER(' + H + '),' + H + '=TODAY()');
  const retard = formule('ISNUMBER(' + H + '),' + H + '<TODAY()');
  const avenir = formule('ISNUMBER(' + H + '),' + H + '>TODAY()');
  const regle = (fo, fond, texte, plage) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(fo).setBackground(fond).setFontColor(texte).setBold(true).setRanges([plage]).build();

  const nouvelles = [];
  (CONFIG.COULEUR_ETAT || []).forEach(k => {
    if (!c[k]) return;
    const plage = f.getRange(debut, colIndex_(c[k]) + 1, nb, 1);
    plage.setBackground(null); // les anciennes couleurs posées à la main sont remplacées par les couleurs automatiques
    nouvelles.push(
      regle(aujour, '#f97316', '#ffffff', plage),   // orange = retour aujourd'hui
      regle(retard, '#ff0000', '#ffffff', plage),   // rouge  = date dépassée
      regle(avenir, '#16a34a', '#ffffff', plage)    // vert   = date pas encore atteinte
    );
  });
  nouvelles.push(regle(formule('ISNUMBER(' + U + '),' + U + '>0'), '#ff0000', '#ffffff', f.getRange(debut, colIndex_(c.DETTE) + 1, nb, 1)));

  const gardees = f.getConditionalFormatRules().filter(r => {
    const b = r.getBooleanCondition();
    return !(b && b.getCriteriaValues().some(v => String(v).indexOf(MARQUE_REGLE_) >= 0));
  });
  f.setConditionalFormatRules(nouvelles.concat(gardees));
}

/** Crée / complète l'onglet des locations. Retourne { f, nbTitres, resume }. Rien n'est effacé : seules les cases vides sont remplies. */
function preparerEtatLocation_(ss, opts) {
  opts = opts || {};   // opts.forcer = écrase aussi les cases déjà remplies des colonnes calculées ; opts.resteTout = calcule RESTE / DETTE même si PAYE est vide
  const nomsOnglets = ['Voitures', 'Lavages', 'Entretien', 'Totaux', NOM_RESERVATIONS_];
  let f = ss.getSheetByName(CONFIG.FEUILLE_LOCATIONS);
  if (!f) {
    const premier = ss.getSheets()[0];
    if (nomsOnglets.indexOf(premier.getName()) < 0) {
      f = premier;
      if (f.getLastRow() === 0 && f.getLastColumn() === 0) f.setName(CONFIG.FEUILLE_LOCATIONS);
    } else {
      f = ss.insertSheet(CONFIG.FEUILLE_LOCATIONS, 0);
    }
  }
  const c = CONFIG.COL, idx = k => colIndex_(c[k]);
  const e = CONFIG.LIGNE_ENTETE, debut = e + 1;
  const nbCols = Math.max.apply(null, Object.keys(c).map(idx)) + 1;
  const dernier = Math.max(CONFIG.LIGNES_FORMULES, f.getLastRow() + 300);
  if (f.getMaxRows() < dernier) f.insertRowsAfter(f.getMaxRows(), dernier - f.getMaxRows());
  if (f.getMaxColumns() < nbCols) f.insertColumnsAfter(f.getMaxColumns(), nbCols - f.getMaxColumns());

  // 1. titres
  const entetes = f.getRange(e, 1, 1, nbCols).getValues()[0];
  let nbTitres = 0;
  TITRES_ETAT_.forEach(([k, titre]) => {
    const col = idx(k);
    if (String(entetes[col]).trim() !== '') return;
    const cell = f.getRange(e, col + 1);
    const rouge = ['DATE_RESTITUTION', 'TELEPHONE', 'DETTE'].indexOf(k) >= 0;
    cell.setValue(titre).setBackground(rouge ? '#dc2626' : '#111827').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    if (NOTES_ETAT_[k]) cell.setNote(NOTES_ETAT_[k]);
    nbTitres++;
  });
  if (nbTitres) {
    if (f.getFrozenRows() < e) f.setFrozenRows(e);
    f.setRowHeight(e, 38);
  }
  if (nbTitres === TITRES_ETAT_.length) { // feuille neuve : largeurs confortables, format des dates
    TITRES_ETAT_.forEach(([k]) => f.setColumnWidth(idx(k) + 1, { CLIENT: 230, VEHICULE: 160, REMARQUE: 240, PAYE: 130, CAUTION: 140, TELEPHONE: 140 }[k] || 120));
    [c.DATE_LIVRAISON, c.DATE_RESTITUTION].forEach(L => f.getRange(debut, colIndex_(L) + 1, dernier - e, 1).setNumberFormat('dd/MM/yyyy'));
    f.setFrozenColumns(1);
  }

  // 2. calculs automatiques (cases vides uniquement)
  const nRows = dernier - e;
  const zone = f.getRange(debut, 1, nRows, nbCols);
  const vals = zone.getValues(), forms = zone.getFormulas();
  const casse = v => /^\s*=/.test(String(v)) || /^#(VALUE|REF|N\/A|NAME|DIV\/0|ERROR|NUM|NULL)/i.test(String(v));
  const vide = (i, k) => forms[i][idx(k)] === '' && (vals[i][idx(k)] === '' || casse(vals[i][idx(k)]));
  const rempli = (i, k) => String(vals[i][idx(k)]).trim() !== '';
  const formules = formulesEtat_();
  const resume = [];
  ORDRE_CALCULS_ETAT_.forEach(k => {
    const cibles = vals.map((_, i) => {
      if (!opts.forcer && !vide(i, k)) return false;
      // RESTE / DETTE : lignes futures, ou lignes déjà saisies où PAYE est rempli (sinon tes anciennes lignes passeraient toutes en « dette »)
      if (k === 'RESTE' || k === 'DETTE') return !!opts.resteTout || !rempli(i, 'CLIENT') || rempli(i, 'PAYE');
      return true;
    });
    // cases « cassées » non ciblées (formule enregistrée comme texte, erreur saisie) : on les vide
    viderSerieEtat_(f, idx(k), debut, vals.map((_, i) => !cibles[i] && forms[i][idx(k)] === '' && casse(vals[i][idx(k)])));
    const nb = remplirSerieEtat_(f, idx(k), debut, cibles, r => formules[k](r), FORMATS_CALCULS_ETAT_[k]);
    resume.push(k + ' : ' + nb);
  });

  // 3. couleurs automatiques
  appliquerCouleursEtat_(f, debut, dernier);

  // 4. diagnostic : dates / prix saisis en TEXTE (les formules ne peuvent pas les calculer)
  const diag = { dates: [], prix: [] };
  vals.forEach((l, i) => {
    if (!rempli(i, 'CLIENT')) return;
    const texte = k => typeof l[idx(k)] === 'string' && l[idx(k)].trim() !== '';
    if (texte('DATE_LIVRAISON') || texte('DATE_RESTITUTION')) diag.dates.push(debut + i);
    if (texte('PRIX_UNITAIRE')) diag.prix.push(debut + i);
  });
  return { f: f, nbTitres: nbTitres, resume: resume, diag: diag };
}

function texteDiag_(d) {
  const liste = a => a.slice(0, 8).join(', ') + (a.length > 8 ? '… (' + a.length + ' lignes au total)' : '');
  let t = '';
  if (d.dates.length) t += '\nAttention : date en texte (ex. « OUVERTE »), non calculable, aux lignes : ' + liste(d.dates) + '.';
  if (d.prix.length) t += '\nAttention : prix unitaire en texte, non calculable, aux lignes : ' + liste(d.prix) + '.';
  return t;
}

/*  RÉSERVATIONS À VENIR  */

const NOM_RESERVATIONS_ = 'Réservations';
const ENTETES_RESERVATIONS_ = ['DU (DÉBUT)', 'AU (FIN)', 'HEURE DE LIVRAISON', 'NB JOURS', 'VÉHICULE', 'NOM DU CLIENT', 'NUM° TLF',
  'PRIX PAR JOUR (DA)', 'TOTAL HT (DA)', 'LIEU DE LIVRAISON', 'COMMERCIAL EN CONTACT', 'REMARQUE', 'STATUT'];

/** Réservations dont le départ tombe entre aujourd'hui et aujourd'hui + jours (les « Annulée » sont ignorées). */
function lireReservations_(ss, aujourd, r, jours) {
  const f = ss.getSheetByName(NOM_RESERVATIONS_);
  if (!f || f.getLastRow() < 2) return;
  const limite = decale_(aujourd, jours);
  f.getRange(2, 1, f.getLastRow() - 1, 13).getValues().forEach(l => {
    const client = String(l[5]).trim(), du = enDate_(l[0]), statut = String(l[12]).trim();
    if (!client || !du || /annul/i.test(statut)) return;
    if (du < aujourd || du > limite) return;
    const au = enDate_(l[1]);
    const prix = montant_(l[7]);
    const nbj = au ? Math.max(0, joursEntre_(du, au)) : 0;
    let total = montant_(l[8]);
    if (!total && prix && nbj) total = prix * nbj;
    r.reservations.push({ du: du, au: au, heure: heure_(l[2]), jours: nbj, vehicule: String(l[4]).trim(), client: client, tel: String(l[6]).trim(),
      prix: prix, total: total, lieu: String(l[9]).trim(), commercial: String(l[10]).trim(), remarque: String(l[11]).trim(), statut: statut,
      dans: joursEntre_(aujourd, du) });
  });
  r.reservations.sort((a, b) => (a.du - b.du) || (minutes_(a.heure) - minutes_(b.heure)));
}

function lienWhatsappResa_(x) {
  const num = telWhatsapp_(x.tel);
  if (!num) return '';
  const texte = 'Bonjour ' + x.client + ', nous vous rappelons votre réservation ' + (x.vehicule ? 'du véhicule ' + x.vehicule + ' ' : '') +
    'du ' + dateFr_(x.du) + (x.au ? ' au ' + dateFr_(x.au) : '') + (x.heure ? ' à ' + x.heure : '') +
    (x.lieu ? ', livraison : ' + x.lieu : '') + '. Merci, ' + CONFIG.SOCIETE + '.';
  return 'https://wa.me/' + num + '?text=' + encodeURIComponent(texte);
}

function ligneReservation_(x) {
  const quand = x.dans === 0 ? pastille_('Aujourd’hui', '#b91c1c', '#ffffff')
    : x.dans === 1 ? pastille_('Demain', '#a4161a', '#ffffff')
    : pastille_('Dans ' + x.dans + ' jours', '#475569', '#ffffff');
  const pastilles = [quand];
  if (x.heure) pastilles.push(pastille_(h_(x.heure), '#fbe9e9', '#a4161a'));
  if (/confirm/i.test(x.statut)) pastilles.push(pastille_('Confirmée', '#dff3ef', '#0f766e'));
  else if (/attente/i.test(x.statut)) pastilles.push(pastille_('En attente', '#fdf0e1', '#92400e'));

  const details = [champ_('Période', dateFr_(x.du) + (x.au ? ' au ' + dateFr_(x.au) : ''))];
  if (x.jours) details.push(champ_('Durée', pl_(x.jours, 'jour', 'jours')));
  if (x.lieu) details.push(champ_('Lieu de livraison', h_(x.lieu)));
  if (x.commercial) details.push(champ_('Commercial', h_(x.commercial)));
  if (x.tel && /\d/.test(x.tel)) details.push(champ_('Téléphone', telHtml_(x.tel)));
  if (x.prix) details.push(champ_('Prix par jour', fmt_(x.prix) + '&nbsp;DA'));
  if (x.total) details.push(champ_('Total HT', fmt_(x.total) + '&nbsp;DA'));

  const actions = [], wa = lienWhatsappResa_(x), num = telWhatsapp_(x.tel);
  if (wa) actions.push(bouton_('Rappeler par WhatsApp', wa, true));
  if (num) actions.push(bouton_('Appeler', 'tel:+' + num, false));
  return carte_({
    couleur: x.dans <= 1 ? COULEUR_.marine : COULEUR_.violet, pastilles: pastilles,
    titre: h_(x.client), sous: h_(x.vehicule), details: details, note: x.remarque ? fr_(x.remarque) : '', actions: actions
  });
}

/* VOITURES : synchronisation des plaques  */

/** Ajoute à l'onglet Voitures les plaques de l'état de location qui n'y sont pas encore (intervalle de vidange par défaut, à choisir ensuite par voiture). */
function synchroniserVoitures_(f, loc) {
  if (!loc || loc.getLastRow() <= CONFIG.LIGNE_ENTETE) return 0;
  const connues = {};
  const colA = f.getRange(2, 1, Math.max(f.getMaxRows() - 1, 1), 1).getValues();
  let dernierA = -1;
  colA.forEach((l, i) => { const k = cle_(l[0]); if (k) { connues[k] = true; dernierA = i; } });
  const ip = colIndex_(CONFIG.COL.PLAQUE), im = colIndex_(CONFIG.COL.VEHICULE);
  const n = loc.getLastRow() - CONFIG.LIGNE_ENTETE;
  const out = [];
  loc.getRange(CONFIG.LIGNE_ENTETE + 1, 1, n, Math.max(ip, im) + 1).getValues().forEach(l => {
    const p = String(l[ip]).trim(), k = cle_(p);
    if (k && !connues[k]) { connues[k] = true; out.push([p, String(l[im]).trim()]); }
  });
  if (!out.length) return 0;
  const premiere = 2 + dernierA + 1;
  if (f.getMaxRows() < premiere + out.length - 1) f.insertRowsAfter(f.getMaxRows(), premiere + out.length - 1 - f.getMaxRows());
  f.getRange(premiere, 1, out.length, 2).setValues(out);
  f.getRange(premiere, 5, out.length, 1).setValue(CONFIG.INTERVALLE_VIDANGE_DEFAUT);
  return out.length;
}

/* FORMULES DES AUTRES ONGLETS  */

/** Supprime les cases « cassées » (formule enregistrée comme TEXTE, ou erreur saisie) sur les lignes ciblées. */
function viderSerieEtat_(f, col, premiere, cibles) {
  let i = 0;
  while (i < cibles.length) {
    if (!cibles[i]) { i++; continue; }
    let j = i;
    while (j + 1 < cibles.length && cibles[j + 1]) j++;
    f.getRange(premiere + i, col + 1, j - i + 1, 1).clearContent();
    i = j + 1;
  }
}

/**
 * Voitures : Km actuel (lu dans l'état de location), prochaine vidange, km restants, statut.
 */
function ecrireFormulesVoitures_(f, nomLoc) {
  const L = "'" + nomLoc.replace(/'/g, "''") + "'!";
  const c = CONFIG.COL;
  const col = k => L + c[k] + ':' + c[k];
  const n = 300;

  // Km actuel : on garde ce que tu as saisi, on ne remplit que les cases vides (ou cassées) avec le plus grand km vu dans les locations
  const zoneC = f.getRange(2, 3, n - 1, 1);
  const vC = zoneC.getValues(), fC = zoneC.getFormulas();
  const casse = v => /^\s*=/.test(String(v)) || /^#(VALUE|REF|N\/A|NAME|DIV\/0|ERROR|NUM|NULL)/i.test(String(v));
  const aRemplir = vC.map((l, i) => fC[i][0] === '' && (l[0] === '' || casse(l[0])));
  viderSerieEtat_(f, 2, 2, vC.map((l, i) => fC[i][0] === '' && casse(l[0])));
  remplirSerieEtat_(f, 2, 2, aRemplir, r => '=IF(A' + r + '="","",IFERROR(LET(m,MAX(MAXIFS(' + col('KM_LIVRAISON') + ',' + col('PLAQUE') + ',A' + r + '),MAXIFS(' + col('KM_RESTITUTION') + ',' + col('PLAQUE') + ',A' + r + ')),IF(m>0,m,"")),""))', '#,##0');

  f.getRange('D2:E' + n).setNumberFormat('#,##0');
  f.getRange('F2:G' + n).setNumberFormat('#,##0');
  f.getRange('H2:H' + n).setNumberFormat('0');                 // jamais « @ » (texte) sur une case qui contient une formule
  f.getRange('I2:J' + n).setNumberFormat('dd/MM/yyyy');
  f.getRange('F2:F' + n).setFormula('=IFERROR(IF(OR(A2="",D2="",E2=""),"",D2+E2),"")');
  f.getRange('G2:G' + n).setFormula('=IFERROR(IF(OR(F2="",C2=""),"",F2-C2),"")');
  f.getRange('H2:H' + n).setFormula('=IFERROR(IF(G2="","",IF(G2<=0,"Dépassée",IF(G2<=' + CONFIG.SEUIL_VIDANGE_KM + ',"Bientôt","OK"))),"")');
}

function ecrireFormulesReservations_(f) {
  const n = 1000;
  f.getRange('A2:B' + n).setNumberFormat('dd/MM/yyyy');
  f.getRange('D2:D' + n).setNumberFormat('0');
  f.getRange('H2:I' + n).setNumberFormat('#,##0.00');
  f.getRange('D2:D' + n).setFormula('=IFERROR(IF(AND(ISNUMBER(A2),ISNUMBER(B2)),MAX(0,B2-A2),""),"")');
  f.getRange('I2:I' + n).setFormula('=IFERROR(IF(AND(ISNUMBER(H2),ISNUMBER(D2)),H2*D2,""),"")');
}

function ecrireTotaux_(f, nomLoc) {
  const L = "'" + nomLoc.replace(/'/g, "''") + "'!";
  const col = k => L + CONFIG.COL[k] + '2:' + CONFIG.COL[k];
  const moisDebut = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)';
  const duMois = (somme, dates) => '=IFERROR(SUMIFS(' + somme + ',' + dates + ',">="&' + moisDebut + ',' + dates + ',"<"&EDATE(' + moisDebut + ',1)),0)';
  const total = plage => '=IFERROR(SUM(' + plage + '),0)';
  f.getRange('A2:A11').setValues([['Locations — total HT du mois (livraisons)'], ['Locations — total TTC du mois'], ['Locations — total HT général'],
    ['Reste à encaisser (HT) — toutes locations'], ['Dettes clients (HT) — locations terminées'],
    ['Lavages — total du mois'], ['Lavages — total général'], ['Entretien / réparations — total du mois'], ['Entretien / réparations — total général'],
    ['Réservations à venir — total HT']]);
  f.getRange('B2:B11').setNumberFormat('#,##0');
  f.getRange('B2').setFormula(duMois(col('TOTAL_HT'), col('DATE_LIVRAISON')));
  f.getRange('B3').setFormula(duMois(col('TOTAL_TTC'), col('DATE_LIVRAISON')));
  f.getRange('B4').setFormula(total(col('TOTAL_HT')));
  f.getRange('B5').setFormula(total(col('RESTE')));
  f.getRange('B6').setFormula(total(col('DETTE')));
  f.getRange('B7').setFormula(duMois('Lavages!D2:D', 'Lavages!A2:A'));
  f.getRange('B8').setFormula(total('Lavages!D2:D'));
  f.getRange('B9').setFormula(duMois('Entretien!E2:E', 'Entretien!A2:A'));
  f.getRange('B10').setFormula(total('Entretien!E2:E'));
  f.getRange('B11').setFormula('=IFERROR(SUMIFS(\'' + NOM_RESERVATIONS_ + '\'!I2:I,\'' + NOM_RESERVATIONS_ + '\'!A2:A,">="&TODAY()),0)');
  f.getRange('A13').setValue('Lavages par laveur').setFontWeight('bold');
  f.getRange('A14').setFormula('=IFERROR(QUERY(Lavages!A2:F,"select E, count(B), sum(D) where E is not null group by E label E \'Laveur\', count(B) \'Nb\', sum(D) \'Montant (DA)\'",0),"")');
  f.getRange('D13').setValue('Lavages par voiture').setFontWeight('bold');
  f.getRange('D14').setFormula('=IFERROR(QUERY(Lavages!A2:F,"select B, count(B), sum(D) where B is not null group by B label B \'Plaque\', count(B) \'Nb\', sum(D) \'Montant (DA)\'",0),"")');
  f.getRange('G13').setValue('Dépenses d’entretien par voiture').setFontWeight('bold');
  f.getRange('G14').setFormula('=IFERROR(QUERY(Entretien!A2:G,"select B, count(B), sum(E) where B is not null group by B label B \'Plaque\', count(B) \'Nb\', sum(E) \'Montant (DA)\'",0),"")');
  f.setColumnWidth(1, 320);
}

/* ====================== DIAGNOSTIC ET RÉPARATION ====================== */

function adresse_(ligne, col0) {
  let s = '', n = col0 + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s + ligne;
}

/** Cherche, dans tous les onglets : cellules en erreur, formules écrites en TEXTE, formules dans une case au format « texte ». */
function diagnostiquerFeuilles_(ss) {
  const res = [];
  ss.getSheets().forEach(f => {
    const nl = f.getLastRow(), nc = f.getLastColumn();
    if (nl < 1 || nc < 1) return;
    const rng = f.getRange(1, 1, nl, nc);
    const vals = rng.getValues(), fors = rng.getFormulas(), fmts = rng.getNumberFormats();
    const d = { nom: f.getName(), erreurs: [], texteFormule: [], formatTexte: [] };
    vals.forEach((row, i) => row.forEach((v, j) => {
      const fo = fors[i][j], s = typeof v === 'string' ? v : '';
      if (/^#(VALUE|REF|N\/A|NAME|DIV\/0|ERROR|NUM|NULL)/i.test(s)) d.erreurs.push(adresse_(i + 1, j));
      else if (!fo && /^\s*=/.test(s)) d.texteFormule.push(adresse_(i + 1, j));
      if (fo && fmts[i][j] === '@') d.formatTexte.push(adresse_(i + 1, j));
    }));
    res.push(d);
  });
  return res;
}

function texteDiagnostic_(res) {
  const ex = a => a.slice(0, 4).join(', ') + (a.length > 4 ? '…' : '');
  return res.map(d => {
    const p = [];
    if (d.erreurs.length) p.push(d.erreurs.length + ' cellule(s) en erreur (' + ex(d.erreurs) + ')');
    if (d.texteFormule.length) p.push(d.texteFormule.length + ' formule(s) écrite(s) en TEXTE (' + ex(d.texteFormule) + ')');
    if (d.formatTexte.length) p.push(d.formatTexte.length + ' formule(s) dans une case au format texte (' + ex(d.formatTexte) + ')');
    return '• ' + d.nom + ' : ' + (p.length ? p.join(' ; ') : 'aucun problème détecté');
  }).join('\n');
}

function afficherDiagnostic() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  let locale = '';
  try { locale = ss.getSpreadsheetLocale(); } catch (e) { /* inconnu */ }
  ui.alert('Diagnostic des formules', 'Paramètres régionaux du fichier : ' + (locale || 'inconnus') + '\n\n' + texteDiagnostic_(diagnostiquerFeuilles_(ss)) +
    '\n\nSi des problèmes sont listés : menu Flotte > Réparer les formules.', ui.ButtonSet.OK);
}

/**
 * Transforme en VRAIES dates / VRAIS nombres les valeurs saisies en texte (« 25/03/2026 », « 142 913 », « 3000,00 »).
 * Les dates sont écrites en numéro de série : aucun décalage de fuseau possible. Les formules sont conservées.
 */
function convertirTextes_(f, premiere, dernier, colsDates, colsNombres) {
  const res = { dates: 0, nombres: 0 };
  const n = dernier - premiere + 1;
  if (!f || n < 1) return res;
  const traiter = (col0, estDate, format) => {
    const rng = f.getRange(premiere, col0 + 1, n, 1);
    const vals = rng.getValues(), fors = rng.getFormulas();
    let modif = 0;
    const out = vals.map((l, i) => {
      const v = l[0];
      if (fors[i][0]) return [fors[i][0]];
      if (typeof v === 'string') {
        const s = v.trim();
        if (estDate && /^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{4}$/.test(s)) {
          const d = enDate_(s);
          if (d) { modif++; return [d.getTime() / 86400000 + 25569]; }
        }
        if (!estDate && /^-?[\d\s\u00A0.,]+$/.test(s) && /\d/.test(s)) { modif++; return [nombre_(s)]; }
      }
      return [v];
    });
    if (modif) {
      rng.setNumberFormat(format);   // d'abord un vrai format (pas « texte »), puis les valeurs
      rng.setValues(out);
      if (estDate) res.dates += modif; else res.nombres += modif;
    }
  };
  (colsDates || []).forEach(c0 => traiter(c0, true, 'dd/MM/yyyy'));
  (colsNombres || []).forEach(([c0, format]) => traiter(c0, false, format));
  return res;
}

function convertirTousTextes_(ss) {
  const tot = { dates: 0, nombres: 0 };
  const ajoute = r => { tot.dates += r.dates; tot.nombres += r.nombres; };
  const c = CONFIG.COL, idx = k => colIndex_(c[k]);
  const loc = ss.getSheetByName(CONFIG.FEUILLE_LOCATIONS) || ss.getSheets()[0];
  ajoute(convertirTextes_(loc, CONFIG.LIGNE_ENTETE + 1, Math.max(CONFIG.LIGNES_FORMULES, loc.getLastRow() + 300),
    [idx('DATE_LIVRAISON'), idx('DATE_RESTITUTION')], [[idx('KM_LIVRAISON'), '#,##0'], [idx('KM_RESTITUTION'), '#,##0'], [idx('PRIX_UNITAIRE'), '#,##0.00']]));
  [['Voitures', [8, 9], [[3, '#,##0'], [4, '#,##0']]], ['Lavages', [0], [[3, '#,##0.00']]], ['Entretien', [0], [[4, '#,##0.00'], [5, '#,##0']]],
    [NOM_RESERVATIONS_, [0, 1], [[7, '#,##0.00']]]].forEach(([nom, dates, nombres]) => {
    const f = ss.getSheetByName(nom);
    if (f && f.getLastRow() >= 2) ajoute(convertirTextes_(f, 2, f.getLastRow(), dates, nombres));
  });
  return tot;
}

/** Menu « Réparer les formules » : formats, textes convertis, formules réécrites dans tous les onglets. */
function reparerFormules() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  const avant = diagnostiquerFeuilles_(ss);
  const rep = ui.alert('Réparer les formules',
    'État actuel :\n' + texteDiagnostic_(avant) + '\n\n' +
    'Le script va :\n' +
    '1. retirer le format « texte » des colonnes calculées ;\n' +
    '2. convertir en vraies dates / vrais nombres ceux qui sont écrits en texte (état de location, Voitures, Lavages, Entretien, Réservations) ;\n' +
    '3. réécrire les formules (Voitures, Réservations, Totaux, et dans l’état de location : CONSOMMATION KM, VALEUR, TOTAL HT, TOTAL TTC, RESTE, DETTE).\n' +
    'Tes autres données ne sont pas modifiées.\n\n' +
    'RESTE et DETTE = TOTAL HT - PAYE.\n' +
    'OUI : calculer toutes les lignes (PAYE vide = rien payé : les anciennes locations terminées deviennent des dettes).\n' +
    'NON : calculer RESTE et DETTE seulement quand PAYE est rempli.\n' +
    'ANNULER : ne rien faire.', ui.ButtonSet.YES_NO_CANCEL);
  if (rep !== ui.Button.YES && rep !== ui.Button.NO) return;

  const conv = convertirTousTextes_(ss);
  const etat = preparerEtatLocation_(ss, { forcer: true, resteTout: rep === ui.Button.YES });
  const nomLoc = etat.f.getName();
  const v = ss.getSheetByName('Voitures'); if (v) ecrireFormulesVoitures_(v, nomLoc);
  const rs = ss.getSheetByName(NOM_RESERVATIONS_); if (rs) ecrireFormulesReservations_(rs);
  const t = ss.getSheetByName('Totaux');
  if (t && /^locations/i.test(String(t.getRange('A2').getValue()))) ecrireTotaux_(t, nomLoc);

  const apres = diagnostiquerFeuilles_(ss);
  ui.alert('Formules réparées',
    'Dates converties : ' + conv.dates + ' · nombres convertis : ' + conv.nombres + '\n' +
    'Formules de l’état de location : ' + etat.resume.join(', ') + texteDiag_(etat.diag) + '\n\n' +
    'Après réparation :\n' + texteDiagnostic_(apres) +
    '\n\nSi une colonne reste vide, vérifie que les dates et le prix unitaire sont de vrais nombres / vraies dates, et que les colonnes correspondent à CONFIG.COL en haut du code.', ui.ButtonSet.OK);
}

function rafraichirCouleurs() {
  const ss = SpreadsheetApp.getActive();
  const f = ss.getSheetByName(CONFIG.FEUILLE_LOCATIONS) || ss.getSheets()[0];
  appliquerCouleursEtat_(f, CONFIG.LIGNE_ENTETE + 1, Math.max(CONFIG.LIGNES_FORMULES, f.getLastRow() + 300));
  SpreadsheetApp.getUi().alert('Couleurs mises à jour sur l’onglet « ' + f.getName() + ' ». Les autres onglets n’ont pas été modifiés.');
}

/* CREATION DES ONGLETS  */

function initialiserOnglets() {
  const ss = SpreadsheetApp.getActive();
  const etat = preparerEtatLocation_(ss);          // titres + calculs automatiques + couleurs de l'état de location
  const locSheet = etat.f;

  const creer = (nom, entetes) => {
    let f = ss.getSheetByName(nom);
    const neuf = !f;
    if (neuf) f = ss.insertSheet(nom);
    if (neuf || f.getLastRow() === 0) {
      f.getRange(1, 1, 1, entetes.length).setValues([entetes])
        .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff').setWrap(true);
      f.setFrozenRows(1);
    }
    return { f: f, neuf: neuf };
  };
  const liste = (f, plage, valeurs) => f.getRange(plage).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(valeurs, true).setAllowInvalid(true).build());

  // --- Voitures (vidange + documents)
  const v = creer('Voitures', ['Plaque', 'Modèle', 'Km actuel', 'Km dernière vidange', 'Intervalle vidange (km)',
    'Prochaine vidange (km)', 'Km restants', 'Statut vidange', 'Assurance expire le', 'Contrôle technique expire le']);
  if (v.neuf) {
    const f = v.f;
    f.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('Dépassée').setBackground('#f4c7c3').setRanges([f.getRange('H2:H300')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('Bientôt').setBackground('#fce8b2').setRanges([f.getRange('H2:H300')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('OK').setBackground('#b7e1cd').setRanges([f.getRange('H2:H300')]).build()
    ]);
  }
  // onglet existant : on n'y touche pas (ni formules, ni listes, ni couleurs). On y ajoute seulement les nouvelles plaques.
  if (v.neuf) {
    v.f.getRange('E2:E300').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.INTERVALLES_VIDANGE.map(String), true).setAllowInvalid(true)
      .setHelpText('Choisis l’intervalle de vidange de cette voiture (ou saisis un autre nombre).').build());
  }
  const nbNouvelles = synchroniserVoitures_(v.f, locSheet);
  if (v.neuf) ecrireFormulesVoitures_(v.f, locSheet.getName());

  // --- Lavages
  const l = creer('Lavages', ['Date', 'Plaque', 'Type de lavage', 'Prix (DA)', 'Laveur', 'Payé au laveur ?']);
  if (l.neuf) {
    l.f.getRange('A2:A2000').setNumberFormat('dd/MM/yyyy');
    l.f.getRange('D2:D2000').setNumberFormat('#,##0.00');
    liste(l.f, 'C2:C2000', ['Extérieur', 'Intérieur', 'Complet']);
    liste(l.f, 'F2:F2000', ['Oui', 'Non']);
  }

  // --- Entretien / réparations / achats
  const e = creer('Entretien', ['Date', 'Plaque', 'Type', 'Description', 'Prix (DA)', 'Km', 'Statut']);
  if (e.neuf) {
    e.f.getRange('A2:A2000').setNumberFormat('dd/MM/yyyy');
    e.f.getRange('E2:E2000').setNumberFormat('#,##0.00');
    liste(e.f, 'C2:C2000', ['Vidange', 'Réparation', 'Achat de pièce', 'Pneus', 'Contrôle technique', 'Assurance', 'Autre']);
    liste(e.f, 'G2:G2000', ['À faire', 'Fait']);
  }

  // --- Réservations (périodes futures)
  const rs = creer(NOM_RESERVATIONS_, ENTETES_RESERVATIONS_);
  if (rs.neuf) {
    const f = rs.f, n = 1000;
    f.getRange('C2:C' + n).setNumberFormat('@');   // heure saisie comme du texte (14h00) : colonne de saisie, sans formule
    liste(f, 'M2:M' + n, ['Confirmée', 'En attente', 'Annulée']);
    [150, 150, 130, 80, 170, 220, 140, 130, 130, 220, 170, 240, 120].forEach((w, i) => f.setColumnWidth(i + 1, w));
    f.setRowHeight(1, 38);
    const zone = f.getRange('A2:M' + n);
    f.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$M2="Annulée"').setFontColor('#9ca3af').setStrikethrough(true).setRanges([zone]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($A2),$A2=TODAY())').setBackground('#fde2e0').setRanges([zone]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($A2),$A2=TODAY()+1)').setBackground('#fff2cc').setRanges([zone]).build()
    ]);
  }
  if (rs.neuf) ecrireFormulesReservations_(rs.f);

  // --- Totaux
  const t = creer('Totaux', ['Indicateur', 'Montant (DA)']);
  if (t.neuf) ecrireTotaux_(t.f, locSheet.getName());

  SpreadsheetApp.getUi().alert('Onglets prêts : ' + locSheet.getName() + ' (titres, calculs, couleurs), Voitures, Lavages, Entretien, ' + NOM_RESERVATIONS_ + ', Totaux.\n\n' +
    '• État de location : titres ajoutés : ' + etat.nbTitres + '. Formules posées dans les cases vides seulement (RESTE et DETTE sont calculés sur le TOTAL HT, à partir de la colonne PAYE). ' +
    'La case « date de restitution » change de couleur toute seule : vert = date pas encore atteinte, orange = date d’aujourd’hui, rouge = date dépassée.\n' +
    '• Voitures : choisis l’intervalle de vidange de chaque voiture dans la liste (8 000, 10 000…). Le km actuel est lu automatiquement dans l’état de location.' + (nbNouvelles ? ' ' + nbNouvelles + ' nouvelle(s) voiture(s) ajoutée(s) depuis l’état de location.' : '') + '\n' +
    '• Réservations : saisis les réservations futures ; l’option « Réservations à venir » se coche dans « Choisir le contenu de l’e-mail ».' + texteDiag_(etat.diag) +
    '\n\nSi des cases affichent une erreur ou ne calculent pas : menu Flotte  > Réparer les formules.');
}

