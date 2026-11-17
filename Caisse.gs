 
const NOM_CAISSE_ = 'Caisse';
const LIGNES_CAISSE_ = 300;     // nombre de lignes de saisie par tableau (revenus d'un côté, dépenses de l'autre)
const CATEGORIES_DEPENSES_ = ['Carburant', 'Lavage', 'Vidange', 'Réparation', 'Pièces', 'Assurance', 'Péage', 'Salaire', 'Loyer', 'Autre'];
 
/** Position des blocs sur la feuille. */
function geometrieCaisse_() {
  const n = LIGNES_CAISSE_;
  const revDebut = 10, revFin = revDebut + n - 1;
  const depBandeau = revFin + 3, depEntete = depBandeau + 1, depDebut = depEntete + 1, depFin = depDebut + n - 1;
  return { n: n, revBandeau: 8, revEntete: 9, revDebut: revDebut, revFin: revFin, depBandeau: depBandeau, depEntete: depEntete, depDebut: depDebut, depFin: depFin };
}
 
function poserFormulesCaisse_(f, col0, debut, fin, modele) {
  const n = fin - debut + 1;
  const rng = f.getRange(debut, col0 + 1, n, 1);
  const vals = rng.getValues(), forms = rng.getFormulas();
  const libre = i => vals[i][0] === '' && forms[i][0] === '';
  let i = 0, nb = 0;
  while (i < n) {
    if (!libre(i)) { i++; continue; }
    let j = i;
    while (j + 1 < n && libre(j + 1)) j++;
    const lot = [];
    for (let r = debut + i; r <= debut + j; r++) lot.push([modele(r)]);
    f.getRange(debut + i, col0 + 1, j - i + 1, 1).setFormulas(lot);
    nb += j - i + 1;
    i = j + 1;
  }
  return nb;
}
 
function preparerCaisse() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const neuve = !ss.getSheetByName(NOM_CAISSE_);
  const f = ss.getSheetByName(NOM_CAISSE_) || ss.insertSheet(NOM_CAISSE_);
  const g = geometrieCaisse_();
  const loc = ss.getSheetByName(CONFIG.FEUILLE_LOCATIONS) || ss.getSheets()[0];
  const L = "'" + String(loc.getName()).replace(/'/g, "''") + "'!";
  const lc = k => L + '$' + CONFIG.COL[k] + ':$' + CONFIG.COL[k];
 
  const derniere = g.depFin + 5;
  if (f.getMaxRows() < derniere) f.insertRowsAfter(f.getMaxRows(), derniere - f.getMaxRows());
  if (f.getMaxColumns() < 7) f.insertColumnsAfter(f.getMaxColumns(), 7 - f.getMaxColumns());
 
  const REV_A = '$A$' + g.revDebut + ':$A$' + g.revFin, REV_F = '$F$' + g.revDebut + ':$F$' + g.revFin;
  const DEP_A = '$A$' + g.depDebut + ':$A$' + g.depFin, DEP_F = '$F$' + g.depDebut + ':$F$' + g.depFin, DEP_G = '$G$' + g.depDebut + ':$G$' + g.depFin;
 
  // ---- titre + liens
  f.getRange('A1').setValue('CAISSE — REVENUS ET DÉPENSES').setFontSize(16).setFontWeight('bold').setFontColor('#111827');
  f.getRange('A2').setFormula('=HYPERLINK("#gid=' + f.getSheetId() + '&range=A' + g.revDebut + '","Aller aux revenus")');
  f.getRange('C2').setFormula('=HYPERLINK("#gid=' + f.getSheetId() + '&range=A' + g.depDebut + '","Aller aux dépenses")');
 
  // ---- résumé en direct : chaque revenu s'ajoute, chaque dépense et chaque acompte se retire
  const debutMois = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)';
  const duMois = (somme, dates) => 'SUMIFS(' + somme + ',' + dates + ',">="&' + debutMois + ',' + dates + ',"<"&EDATE(' + debutMois + ',1))';
  f.getRange('B3:E3').setValues([['REVENUS', 'DÉPENSES', 'ACOMPTES', 'SOLDE DE LA CAISSE']])
    .setFontWeight('bold').setFontColor('#6b7280').setFontSize(10).setHorizontalAlignment('center');
  f.getRange('A4:A5').setValues([['Total général'], ['Ce mois-ci']]).setFontWeight('bold').setFontColor('#374151');
  f.getRange('B4:E4').setFormulas([['=SUM(' + REV_F + ')', '=SUM(' + DEP_F + ')', '=SUM(' + DEP_G + ')', '=B4-C4-D4']]);
  f.getRange('B5:E5').setFormulas([['=' + duMois(REV_F, REV_A), '=' + duMois(DEP_F, DEP_A), '=' + duMois(DEP_G, DEP_A), '=B5-C5-D5']]);
  f.getRange('B4:E5').setNumberFormat('#,##0.00').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  f.getRange('E4:E5').setFontSize(15).setBackground('#f3f4f6');
  f.getRange('A3:E5').setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
  f.getRange('A6').setFormula('=IF(COUNTA(' + REV_F + ',' + DEP_F + ',' + DEP_G + ')-COUNT(' + REV_F + ',' + DEP_F + ',' + DEP_G + ')>0,' +
    '(COUNTA(' + REV_F + ',' + DEP_F + ',' + DEP_G + ')-COUNT(' + REV_F + ',' + DEP_F + ',' + DEP_G + '))&" montant(s) saisi(s) en texte : non comptés. Retape-les en chiffres.","")')
    .setFontColor('#b91c1c').setFontWeight('bold');
 
  // ---- bloc REVENUS
  f.getRange(g.revBandeau, 1, 1, 7).setBackground('#15803d').setFontColor('#ffffff').setFontWeight('bold');
  f.getRange(g.revBandeau, 1).setValue('REVENUS — entrées d’argent').setFontSize(12);
  f.getRange(g.revBandeau, 5, 1, 2).setFormulas([['="TOTAL REVENUS"', '=SUM(' + REV_F + ')']]).setHorizontalAlignment('right');
  f.getRange(g.revBandeau, 6).setNumberFormat('#,##0.00');
  f.getRange(g.revEntete, 1, 1, 6).setValues([['DATE', 'NOM DU CLIENT', 'N° DE CONTRAT', 'VÉHICULE', 'REMARQUE', 'MONTANT (DA)']])
    .setBackground('#111827').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
 
  // ---- bloc DÉPENSES
  f.getRange(g.depBandeau, 1, 1, 7).setBackground('#b91c1c').setFontColor('#ffffff').setFontWeight('bold');
  f.getRange(g.depBandeau, 1).setValue('DÉPENSES ET ACOMPTES — sorties d’argent').setFontSize(12);
  f.getRange(g.depBandeau, 5, 1, 3).setFormulas([['="TOTAL"', '=SUM(' + DEP_F + ')', '=SUM(' + DEP_G + ')']]).setHorizontalAlignment('right');
  f.getRange(g.depBandeau, 6, 1, 2).setNumberFormat('#,##0.00');
  f.getRange(g.depEntete, 1, 1, 7).setValues([['DATE', 'DÉPENSE', 'DÉPENSÉ PAR', 'VÉHICULE (si concerné)', 'REMARQUE', 'MONTANT DE LA DÉPENSE (DA)', 'ACOMPTE (DA)']])
    .setBackground('#111827').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  f.getRange(g.depEntete, 7).setNote('Acompte versé (avance). Il est retiré de la caisse comme une dépense.');
  f.getRange(g.depEntete, 6).setNote('Montant de la dépense (carburant, réparation…). Il est retiré de la caisse.');
  f.getRange(g.revEntete, 6).setNote('Montant encaissé. Il est ajouté à la caisse.');
  f.getRange(g.revEntete, 3).setNote('Saisis le n° de contrat : le nom du client et le véhicule se remplissent tout seuls (tu peux les écraser).');
 
  // ---- formats, listes, bordures
  const zoneRev = f.getRange(g.revDebut, 1, g.n, 6), zoneDep = f.getRange(g.depDebut, 1, g.n, 7);
  [zoneRev, zoneDep].forEach(z => z.setBorder(true, true, true, true, true, true, '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID).setVerticalAlignment('middle'));
  [f.getRange(g.revDebut, 1, g.n, 1), f.getRange(g.depDebut, 1, g.n, 1)].forEach(z => z.setNumberFormat('dd/MM/yyyy').setHorizontalAlignment('center'));
  f.getRange(g.revDebut, 6, g.n, 1).setNumberFormat('#,##0.00');
  f.getRange(g.depDebut, 6, g.n, 2).setNumberFormat('#,##0.00');
  f.getRange(g.depDebut, 2, g.n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(CATEGORIES_DEPENSES_, true).setAllowInvalid(true).setHelpText('Choisis une dépense ou tape la tienne.').build());
  const voitures = ss.getSheetByName('Voitures');
  if (voitures) {
    const liste = SpreadsheetApp.newDataValidation().requireValueInRange(voitures.getRange('B2:B500'), true).setAllowInvalid(true).build();
    f.getRange(g.revDebut, 4, g.n, 1).setDataValidation(liste);
    f.getRange(g.depDebut, 4, g.n, 1).setDataValidation(liste);
  }
 
  // ---- n° de contrat -> nom du client et véhicule (cases vides seulement ; tu peux écraser)
  const rechercher = (k, r) => '=IF(LEN($C' + r + ')<3,"",IFERROR(INDEX(' + lc(k) + ',MATCH("*"&$C' + r + '&"*",' + lc('CONTRAT') + ',0)),""))';
  const nbClients = poserFormulesCaisse_(f, 1, g.revDebut, g.revFin, r => rechercher('CLIENT', r));
  const nbVehicules = poserFormulesCaisse_(f, 3, g.revDebut, g.revFin, r => rechercher('VEHICULE', r));
 
  // ---- largeurs, volets figés, couleurs du solde
  [130, 230, 160, 190, 280, 190, 150].forEach((w, i) => f.setColumnWidth(i + 1, w));
  f.setRowHeight(g.depEntete, 38);
  f.setFrozenRows(6);
  if (neuve || f.getConditionalFormatRules().length === 0) {
    const soldes = [f.getRange('E4:E5')];
    f.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#b91c1c').setRanges(soldes).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor('#15803d').setRanges(soldes).build()
    ]);
  }
  ss.setActiveSheet(f);
 
  ui.alert('Feuille « ' + NOM_CAISSE_ + ' » prête',
    (neuve ? 'Onglet créé.\n\n' : 'Onglet mis à jour (rien de ce que tu as saisi n’a été effacé).\n\n') +
    'En haut, toujours visible : REVENUS, DÉPENSES, ACOMPTES et SOLDE DE LA CAISSE, en direct.\n' +
    '  • chaque revenu AJOUTE son montant ;\n' +
    '  • chaque dépense et chaque acompte RETIRE son montant.\n\n' +
    'Revenus : lignes ' + g.revDebut + ' à ' + g.revFin + ' · Dépenses : lignes ' + g.depDebut + ' à ' + g.depFin + '.\n' +
    'Si une liste est pleine : insère des lignes au MILIEU du tableau (clic droit sur un numéro de ligne > Insérer), les totaux suivent.\n' +
    'Saisis le n° de contrat : le client et le véhicule se remplissent seuls (' + nbClients + ' cases préparées).', ui.ButtonSet.OK);
}
 
function normaliserCaisse_(f, r1, r2, colsDates, colsNombres) {
  const n = r2 - r1 + 1;
  let total = 0;
  const traiter = (col0, estDate) => {
    const rng = f.getRange(r1, col0 + 1, n, 1);
    const vals = rng.getValues(), forms = rng.getFormulas();
    const neuf = vals.map((l, i) => {
      const v = l[0];
      if (forms[i][0] !== '' || typeof v !== 'string' || v.trim() === '') return null;
      const s = v.trim();
      if (estDate) {
        if (!/^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{4}$/.test(s)) return null;
        const d = enDate_(s);
        return d ? d.getTime() / 86400000 + 25569 : null;
      }
      const t = s.replace(/\s*(DA|DZD)\s*$/i, '');
      return /^-?[\d\s\u00A0.,]*\d[\d\s\u00A0.,]*$/.test(t) ? nombre_(t) : null;
    });
    if (neuf.every(x => x === null)) return;
    rng.setNumberFormat(estDate ? 'dd/MM/yyyy' : '#,##0.00');   // d'abord un vrai format (pas « texte »), puis les valeurs
    let i = 0;
    while (i < n) {
      if (neuf[i] === null) { i++; continue; }
      let j = i;
      while (j + 1 < n && neuf[j + 1] !== null) j++;
      f.getRange(r1 + i, col0 + 1, j - i + 1, 1).setValues(neuf.slice(i, j + 1).map(x => [x]));
      total += j - i + 1;
      i = j + 1;
    }
  };
  colsDates.forEach(c => traiter(c, true));
  colsNombres.forEach(c => traiter(c, false));
  return total;
}
 
/** Appelée par onEdit : les montants et les dates saisis en texte deviennent de vrais nombres, donc le solde se calcule toujours. */
function auCalculCaisse_(e) {
  if (!e || !e.range) return;
  const f = e.range.getSheet();
  if (f.getName() !== NOM_CAISSE_) return;
  const g = geometrieCaisse_();
  const r1 = e.range.getRow(), r2 = e.range.getLastRow();
  [[g.revDebut, g.revFin, [5]], [g.depDebut, g.depFin, [5, 6]]].forEach(([debut, fin, nombres]) => {
    const a = Math.max(r1, debut), b = Math.min(r2, fin);
    if (b < a || b - a > 300) return;
    normaliserCaisse_(f, a, b, [0], nombres);
  });
}
 


