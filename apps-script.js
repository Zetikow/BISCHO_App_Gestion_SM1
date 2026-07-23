// ===================================================================
// SCRIPT GOOGLE APPS SCRIPT — Caisse Noire (v2 : rôles, présence, paiements)
// ===================================================================
// COMMENT INSTALLER / METTRE A JOUR :
// 1. Ouvre ton Google Sheet existant > Extensions > Apps Script
// 2. Supprime TOUT le code existant
// 3. Colle TOUT le contenu de ce fichier à la place
// 4. Sélectionne la fonction "setup" en haut, clique "Exécuter" (▶)
//    -> Cela recrée la feuille "Grid" (remise à zéro) ET crée les
//       nouvelles feuilles "Comptes", "Presences", "Paiements"
// 5. Va dans le Google Sheet, onglet "Comptes" : change les codes PIN
//    par défaut (colonne C) et ajuste les rôles si besoin
//    (Joueur / Coach / Admin)
// 6. Si tu avais déjà déployé une Application Web, pas besoin de
//    redéployer : "Déployer > Gérer les déploiements > crayon >
//    Nouvelle version > Déployer" suffit pour appliquer ce code
//    sans changer l'URL existante.
// ===================================================================

const ACTIONS = [
  ["participation mensuelle", 5],
  ["Retard entraînement", 1],
  ["Retard match", 1],
  ["Absence non justifié à l'entrainement", 10],
  ["Oubli de vêtement entraînement", 3],
  ["Oubli de vêtement match", 6],
  ["Absence non justifié au match", 50],
  ["Oubli de chasuble", 2],
  ["Taxer une serviette de douche", 2],
  ["Taxer du savon", 1],
  ["Taxer de la crème", 0.5],
  ["Carton Rouge direct", 7],
  ["2min pour avoir râler", 4],
  ["Carton bleu", 15],
  ["Pas de logo BISCHO pour le déplacement", 5],
  ["Ballon dégueulasse (vraiment !)", 2],
  ["Oubli du ballon (match / entrainement)", 2],
  ["Taxer de l'eau", 1],
  ["Pas présent repas après match domicile", 5],
  ["Nom dans le journal", 2],
  ["Photo dans le journal", 4],
  ["Pire action du match (+ déguisement)", 1],
  ["Meilleure action du match", 2],
  ["Autre (à préciser)", 1],
];

const PLAYERS = ["Thomas L.","JM B.","Hugo R.","Yann N.","Arnaud H.","Victor S.","Romain J.","Volodia M.","Victor P.","Mattéo MP.","Gabriel W.","Basile L.","Brahim I.","Maximilien M.","Arthur M.","Damien P.","Alexis W.","Mathieu S.","Nicolas Z.","Robin S."];

// À exécuter UNE FOIS après la mise à jour du barème (ACTIONS) : réorganise les lignes de la
// feuille Grid pour qu'elles correspondent au nouveau barème. Indispensable car les lignes sont
// associées par POSITION, pas par nom — sans cette étape, les montants existants se
// retrouveraient attribués aux mauvaises actions pour tout le monde.
// Convertit aussi l'historique des anciens retards (comptés à l'occurrence) en équivalent
// euros cumulés, pour ne rien perdre.
function migrateGridBareme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Grid");
  const data = sheet.getDataRange().getValues();
  const nbPlayers = PLAYERS.length;
  const firstPlayerCol = 3;
  const lastPlayerCol = firstPlayerCol + nbPlayers - 1;
  const totalCol = lastPlayerCol + 1;
  const firstDataRow = 2;

  const oldRowsByLabel = {};
  for (let i = 1; i < data.length; i++) {
    const label = data[i][0];
    if (label) oldRowsByLabel[String(label).trim()] = data[i].slice(2, 2 + nbPlayers);
  }
  function getOld(label) {
    return oldRowsByLabel[label] || new Array(nbPlayers).fill(0);
  }

  const newValues = ACTIONS.map(([label]) => {
    let values;
    if (label === "Retard entraînement") {
      const v1 = getOld("Retard entraînement 1 min");
      const v5 = getOld("Retard entraînement 5 min");
      const v10 = getOld("Retard entraînement 10 min ou +");
      values = PLAYERS.map((p, idx) => (Number(v1[idx]) || 0) * 1 + (Number(v5[idx]) || 0) * 2 + (Number(v10[idx]) || 0) * 5);
    } else if (label === "Retard match") {
      const vOld = getOld("Retard match");
      values = PLAYERS.map((p, idx) => (Number(vOld[idx]) || 0) * 5); // ancienne valeur : 5€/occurrence
    } else if (label === "Oubli de vêtement match") {
      values = getOld("Oubli tenu de match");
    } else if (label === "Oubli de vêtement entraînement") {
      values = getOld("Oubli de vêtement hors chasuble");
    } else {
      values = getOld(label);
    }
    return values.map(v => Number(v) || 0);
  });

  const oldLastRow = sheet.getLastRow();
  if (oldLastRow >= firstDataRow) {
    sheet.getRange(firstDataRow, 1, oldLastRow - firstDataRow + 1, totalCol).clearContent();
  }

  const lastDataRow = firstDataRow + ACTIONS.length - 1;
  const totalRow = lastDataRow + 1;

  const rowsToWrite = ACTIONS.map((a, i) => [a[0], a[1], ...newValues[i]]);
  sheet.getRange(firstDataRow, 1, rowsToWrite.length, 2 + nbPlayers).setValues(rowsToWrite);

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const playerRange = `${columnToLetter(firstPlayerCol)}${r}:${columnToLetter(lastPlayerCol)}${r}`;
    sheet.getRange(r, totalCol).setFormula(`=SUM(${playerRange})*B${r}`);
  }

  sheet.getRange(totalRow, 1).setValue("TOTAL Joueur");
  for (let c = firstPlayerCol; c <= lastPlayerCol; c++) {
    const col = columnToLetter(c);
    sheet.getRange(totalRow, c).setFormula(`=SUMPRODUCT(${col}${firstDataRow}:${col}${lastDataRow},$B${firstDataRow}:$B${lastDataRow})`);
  }
  sheet.getRange(totalRow, totalCol).setFormula(`=SUM(${columnToLetter(totalCol)}${firstDataRow}:${columnToLetter(totalCol)}${lastDataRow})`);
  sheet.getRange(totalRow, 1, 1, totalCol).setFontWeight("bold");

  Logger.log(ACTIONS.length + " actions réécrites (contre " + (oldLastRow - firstDataRow) + " avant). Vérifie les montants dans Google Sheets pour confirmer que tout est cohérent avant de continuer.");
}



// ===================== SETUP =====================

function setup() {
  setupGrid();
  setupComptes();
  setupPresences();
  setupPaiements();
  setupEvenements();
  setupPresenceEvenements();
  setupActualites();
  setupCovoiturage();
  setupSupport();
  setupOsteoSlots();
  setupOsteoReservations();
  ensureGridAction("Non renseigné avant dimanche soir", 1);
}

function setupGrid() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Grid");
  if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet("Grid"); }
  sheet.setName("Grid");

  const nbPlayers = PLAYERS.length;
  const firstPlayerCol = 3;
  const lastPlayerCol = firstPlayerCol + nbPlayers - 1;
  const totalCol = lastPlayerCol + 1;
  const firstDataRow = 2;
  const lastDataRow = firstDataRow + ACTIONS.length - 1;
  const totalRow = lastDataRow + 1;

  const header = ["Action", "Valeur", ...PLAYERS, "Total Action"];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  const rows = ACTIONS.map((a, i) => {
    const row = [a[0], a[1]];
    PLAYERS.forEach(() => row.push(i === 0 ? 10 : 0));
    return row;
  });
  sheet.getRange(firstDataRow, 1, rows.length, header.length - 1).setValues(rows);

  let paramSheet = ss.getSheetByName("Paramètres");
  if (!paramSheet) paramSheet = ss.insertSheet("Paramètres");
  paramSheet.clear();
  paramSheet.getRange("A1").setValue("liste du nombre d'action");
  const numberValues = [];
  for (let v = 0; v <= 50; v++) numberValues.push([v]);
  paramSheet.getRange(2, 1, numberValues.length, 1).setValues(numberValues);

  const sourceRange = paramSheet.getRange(2, 1, numberValues.length, 1);
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(firstDataRow, firstPlayerCol, ACTIONS.length, nbPlayers).setDataValidation(validation);

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const playerRange = `${columnToLetter(firstPlayerCol)}${r}:${columnToLetter(lastPlayerCol)}${r}`;
    sheet.getRange(r, totalCol).setFormula(`=SUM(${playerRange})*B${r}`);
  }

  sheet.getRange(totalRow, 1).setValue("TOTAL Joueur");
  for (let c = firstPlayerCol; c <= lastPlayerCol; c++) {
    const col = columnToLetter(c);
    sheet.getRange(totalRow, c).setFormula(`=SUMPRODUCT(${col}${firstDataRow}:${col}${lastDataRow},$B${firstDataRow}:$B${lastDataRow})`);
  }
  sheet.getRange(totalRow, totalCol).setFormula(`=SUM(${columnToLetter(totalCol)}${firstDataRow}:${columnToLetter(totalCol)}${lastDataRow})`);

  sheet.getRange(totalRow, 1, 1, header.length).setFontWeight("bold");
  sheet.getRange(1, 1, 1, header.length).setFontWeight("bold");

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  SpreadsheetApp.flush();
}

// ===================== SCHÉMA COMPTES (une ligne par personne) =====================
// Colonnes : Nom, Code, Roles, Poste, NomComplet, PhotoURL, Email, PushSubIds
// "Roles" contient tous les rôles de la personne dans une seule cellule, au format
// "Joueur:SM1,Coach:U17,Admin:Toutes" — plus besoin de plusieurs lignes reliées par "Personne".
const COL_NOM = 0, COL_CODE = 1, COL_ROLES = 2, COL_POSTE = 3, COL_NOMCOMPLET = 4, COL_PHOTOURL = 5, COL_EMAIL = 6, COL_PUSHSUBIDS = 7;

function parseRoles(cell) {
  return String(cell || "").split(",").map(s => s.trim()).filter(Boolean).map(pair => {
    const idx = pair.indexOf(":");
    if (idx === -1) return { role: pair, equipe: "SM1" };
    return { role: pair.slice(0, idx).trim(), equipe: pair.slice(idx + 1).trim() || "SM1" };
  });
}
function stringifyRoles(arr) {
  return arr.map(r => `${r.role}:${r.equipe}`).join(",");
}
function rowHasRole(row, roleName) {
  return parseRoles(row[COL_ROLES]).some(r => r.role === roleName);
}
function rowEquipesForRole(row, roleName) {
  return parseRoles(row[COL_ROLES]).filter(r => r.role === roleName).map(r => r.equipe);
}

function ensureComptesSchema(sheet) {
  const headers = ["Nom", "Code", "Roles", "Poste", "NomComplet", "PhotoURL", "Email", "PushSubIds"];
  const current = sheet.getRange(1, 1, 1, 8).getValues()[0];
  for (let c = 0; c < 8; c++) {
    if (current[c] !== headers[c]) {
      sheet.getRange(1, c + 1).setValue(headers[c]);
      sheet.getRange(1, c + 1).setFontWeight("bold");
    }
  }
}

function setupComptes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Comptes");
  if (!sheet) sheet = ss.insertSheet("Comptes");
  // Ne PAS effacer si déjà rempli, pour ne pas perdre les codes déjà changés
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 8).setValues([["Nom", "Code", "Roles", "Poste", "NomComplet", "PhotoURL", "Email", "PushSubIds"]]);
    const rows = PLAYERS.map(p => [p, "", "Joueur:SM1", "", "", "", "", ""]);
    rows.push(["Coach", "", "Coach:SM1", "COACH", "", "", "", ""]);
    rows.push(["Admin", "", "Admin:Toutes", "", "", "", "", ""]);
    rows.push(["Bénévole", "", "Bénévole:SM1", "", "", "", "", ""]);
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat("@");
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  ensureComptesSchema(sheet);
}

// À exécuter UNE FOIS depuis l'éditeur pour fusionner les anciennes lignes multiples par
// personne (reliées par "Personne") en UNE SEULE ligne par personne, avec tous ses rôles
// listés dans la colonne "Roles". Fait une sauvegarde complète de l'onglet Comptes avant
// toute modification (nouvel onglet "Comptes_backup_...") — rien n'est perdu en cas de souci.
function migrateComptesToSingleRowPerPerson() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Comptes");
  const data = sheet.getDataRange().getValues();

  const backupName = "Comptes_backup_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const backupSheet = sheet.copyTo(ss);
  backupSheet.setName(backupName);

  // Détecte le format réel en regardant les DONNÉES (pas juste l'en-tête, qui a pu être
  // renommé séparément par ensureComptesSchema sans que les lignes soient migrées) :
  // dans l'ancien format, la colonne B contient un rôle ("Joueur", "Coach"...) ; dans le
  // nouveau, elle ne contient qu'un code à 4 chiffres ou rien.
  const KNOWN_ROLES = ["Joueur", "Coach", "Admin", "Salarié", "Bénévole"];
  const looksOldFormat = data.slice(1).some(row => KNOWN_ROLES.indexOf(String(row[1] || "").trim()) !== -1);
  if (!looksOldFormat) {
    Logger.log("La feuille Comptes semble déjà migrée (aucun rôle trouvé en colonne B). Rien fait. Sauvegarde tout de même créée : " + backupName);
    return;
  }

  // Ancien schéma : Nom(0), Role(1), Code(2), Equipe(3), Personne(4), Poste(5), NomComplet(6), PhotoURL(7), Email(8), PushSubIds(9)
  const groups = {};
  const order = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const personne = row[4] || row[0];
    if (!groups[personne]) { groups[personne] = []; order.push(personne); }
    groups[personne].push(row);
  }

  const newRows = [];
  order.forEach(personne => {
    const rows = groups[personne];
    const roleSet = [];
    let code = "";
    let poste = [];
    let nomComplet = "";
    let photoUrl = "";
    let email = "";
    let pushSubIds = [];
    rows.forEach(row => {
      const role = String(row[1] || "").trim();
      const equipe = String(row[3] || "SM1").trim();
      if (role && !roleSet.some(r => r.role === role && r.equipe === equipe)) roleSet.push({ role, equipe });
      if (!code && row[2]) code = String(row[2]).trim();
      if (row[5]) String(row[5]).split(",").forEach(p => { p = p.trim(); if (p && poste.indexOf(p) === -1) poste.push(p); });
      if (!nomComplet && row[6]) nomComplet = row[6];
      if (!photoUrl && row[7]) photoUrl = row[7];
      if (!email && row[8]) email = row[8];
      if (row[9]) String(row[9]).split(",").forEach(id => { id = id.trim(); if (id && pushSubIds.indexOf(id) === -1) pushSubIds.push(id); });
    });
    newRows.push([personne, code, stringifyRoles(roleSet), poste.join(","), nomComplet, photoUrl, email, pushSubIds.join(",")]);
  });

  sheet.clear();
  const headers = ["Nom", "Code", "Roles", "Poste", "NomComplet", "PhotoURL", "Email", "PushSubIds"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.getRange(1, 2, 1, 1).setNumberFormat("@");
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
    sheet.getRange(2, 2, newRows.length, 1).setNumberFormat("@");
  }
  sheet.setFrozenRows(1);
  Logger.log("Migration terminée : " + newRows.length + " personne(s) consolidée(s) depuis " + (data.length - 1) + " ligne(s) d'origine. Sauvegarde de l'ancien format : " + backupName);
}


// Ajoute la colonne "Score" (ex: "28-24") à la feuille Evenements si elle n'existe pas encore —
// utilisée pour afficher les derniers résultats sans dépendre d'un widget externe.
function ensureEvenementsScoreColumn(sheet) {
  const header = sheet.getRange(1, 1, 1, 8).getValues()[0];
  if (header[7] !== "Score") {
    sheet.getRange(1, 8).setValue("Score");
    sheet.getRange(1, 8).setFontWeight("bold");
  }
}

function setupPresences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Presences");
  if (!sheet) sheet = ss.insertSheet("Presences");
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 3).setValues([["Date", "Joueur", "Present"]]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

function setupPaiements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Paiements");
  if (!sheet) sheet = ss.insertSheet("Paiements");
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 5).setValues([["ID", "Joueur", "Montant", "Date", "Commentaire"]]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  migratePaiements();
}

function setupEvenements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Evenements");
  if (!sheet) sheet = ss.insertSheet("Evenements");
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 7).setValues([["ID", "Date", "Heure", "Type", "Titre", "Lieu", "Equipe"]]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  // Empêche Google Sheets de convertir automatiquement les colonnes Date/Heure
  // en vraies valeurs de type Date/Heure (ce qui cassait l'affichage)
  sheet.getRange(2, 2, 500, 2).setNumberFormat("@");
  migrateEvenementsFormat();
  renameVenueInEvenements("Gymnase Leclerc", "Lustucru Arena");
  ensureEvenementsEquipeColumn(sheet);
}

// Ajoute la colonne "Equipe" si la feuille existait déjà avant son introduction,
// et marque "SM1" par défaut les événements déjà créés (tous étaient SM1 jusqu'ici).
function ensureEvenementsEquipeColumn(sheet) {
  const header = sheet.getRange(1, 1, 1, 7).getValues()[0];
  if (header[6] !== "Equipe") {
    sheet.getRange(1, 7).setValue("Equipe");
    sheet.getRange(1, 7).setFontWeight("bold");
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && !data[i][6]) {
      sheet.getRange(i + 1, 7).setValue("SM1");
    }
  }
}

// Renomme un lieu dans tous les événements existants (utile en cas de renommage du gymnase)
function renameVenueInEvenements(oldName, newName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Evenements");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === oldName) {
      sheet.getRange(i + 1, 6).setValue(newName);
    }
  }
}

// Corrige les événements déjà enregistrés avec le bug de conversion Date/Heure
function migrateEvenementsFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Evenements");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const dateVal = data[i][1];
    const heureVal = data[i][2];
    if (Object.prototype.toString.call(dateVal) === "[object Date]") {
      const y = dateVal.getFullYear(), m = String(dateVal.getMonth() + 1).padStart(2, "0"), d = String(dateVal.getDate()).padStart(2, "0");
      sheet.getRange(i + 1, 2).setNumberFormat("@").setValue(`${y}-${m}-${d}`);
    }
    if (Object.prototype.toString.call(heureVal) === "[object Date]") {
      const h = String(heureVal.getHours()).padStart(2, "0"), min = String(heureVal.getMinutes()).padStart(2, "0");
      sheet.getRange(i + 1, 3).setNumberFormat("@").setValue(`${h}:${min}`);
    }
  }
}

function setupPresenceEvenements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("PresenceEvenements");
  if (!sheet) sheet = ss.insertSheet("PresenceEvenements");
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 4).setValues([["EventID", "Nom", "Present", "Justification"]]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  ensurePresenceEvenementsSchema(sheet);
}

function setupActualites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Actualites");
  if (!sheet) sheet = ss.insertSheet("Actualites");
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 6).setValues([["ID", "Titre", "Scope", "Texte", "Auteur", "Date"]]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

// Covoiturage : une ligne par (événement, personne). JeConduit/BesoinPlace = "Oui"/"" .
function setupCovoiturage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Covoiturage");
  if (!sheet) sheet = ss.insertSheet("Covoiturage");
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 5).setValues([["EventID", "Nom", "JeConduit", "Places", "BesoinPlace"]]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

// Suivi des demandes de support (date, nom, message, réponse — la réponse reste vide pour
// l'instant, prête pour une future fonctionnalité de réponse visible depuis l'appli).
function setupSupport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Support");
  if (!sheet) sheet = ss.insertSheet("Support");
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 5).setValues([["ID", "Date", "Nom", "Message", "Reponse"]]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

// Créneaux de RDV ostéo proposés par Eve (ou l'Admin). RecurrentId regroupe les créneaux créés
// en série (même jour chaque semaine, sur plusieurs semaines).
function setupOsteoSlots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("OsteoSlots");
  if (!sheet) sheet = ss.insertSheet("OsteoSlots");
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 6).setValues([["ID", "Date", "Heure", "Lieu", "Equipe", "RecurrentId"]]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

// Réservations : une ligne par créneau réservé. Motif = raison de consultation, strictement
// privée (jamais montrée aux autres joueurs, seulement à Eve et à la personne concernée).
function setupOsteoReservations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("OsteoReservations");
  if (!sheet) sheet = ss.insertSheet("OsteoReservations");
  if (sheet.getDataRange().getNumRows() <= 1) {
    sheet.getRange(1, 1, 1, 3).setValues([["SlotID", "Nom", "Motif"]]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

// Ajoute la colonne "Justification" si la feuille existait déjà avant son introduction.
function ensurePresenceEvenementsSchema(sheet) {
  const header = sheet.getRange(1, 1, 1, 4).getValues()[0];
  if (header[3] !== "Justification") {
    sheet.getRange(1, 4).setValue("Justification");
    sheet.getRange(1, 4).setFontWeight("bold");
  }
}

// A exécuter une fois si tu avais déjà des paiements enregistrés avant
// l'ajout de la colonne ID (ajoute les ID manquants automatiquement).
function migratePaiements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Paiements");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return;

  if (data[0][0] !== "ID") {
    // Ancien format sans colonne ID : on l'insère en première position
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue("ID");
    for (let i = 1; i < data.length; i++) {
      sheet.getRange(i + 1, 1).setValue("p" + Date.now() + "_" + i);
    }
  } else {
    // Header déjà bon, on comble juste les ID vides éventuels
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) {
        sheet.getRange(i + 1, 1).setValue("p" + Date.now() + "_" + i);
      }
    }
  }
}

function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// ===================== RÈGLE "DISPONIBILITÉS AVANT DIMANCHE SOIR" =====================
// Chaque dimanche soir, sanctionne automatiquement (comme une absence non justifiée) tout
// joueur SM1 qui n'a pas répondu Présent/Absent aux entraînements de la semaine à venir.
// À exécuter une seule fois manuellement : installWeeklyDisponibilitesTrigger()
// (menu Apps Script > sélectionner cette fonction > Exécuter). Elle crée le déclencheur
// hebdomadaire ; pas besoin de la relancer ensuite.

function installWeeklyDisponibilitesTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "checkDisponibilitesDimanche") ScriptApp.deleteTrigger(t);
  });
  // Déclenché juste après minuit dimanche->lundi (l'échéance est "dimanche soir/minuit").
  ScriptApp.newTrigger("checkDisponibilitesDimanche")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(0)
    .create();
}

function checkDisponibilitesDimanche() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evenementsSheet = ss.getSheetByName("Evenements");
  const comptesSheet = ss.getSheetByName("Comptes");
  const presenceEvSheet = ss.getSheetByName("PresenceEvenements");
  const gridSheet = ss.getSheetByName("Grid");
  if (!evenementsSheet || !comptesSheet || !presenceEvSheet || !gridSheet) return;
  ensurePresenceEvenementsSchema(presenceEvSheet);

  const evenements = evenementsSheet.getDataRange().getValues();
  const comptes = comptesSheet.getDataRange().getValues();
  const presenceEv = presenceEvSheet.getDataRange().getValues();

  const now = new Date();
  const in7Days = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  // Entraînements SM1 de la semaine à venir (la caisse noire reste réservée à la SM1 pour le moment)
  const upcomingTrainings = [];
  for (let i = 1; i < evenements.length; i++) {
    const row = evenements[i];
    if (!row[0] || row[3] !== "Entraînement") continue;
    if (String(row[6] || "SM1").trim() !== "SM1") continue;
    const d = new Date(String(row[1]) + "T" + (row[2] || "00:00"));
    if (d >= now && d <= in7Days) upcomingTrainings.push(row);
  }
  if (upcomingTrainings.length === 0) return;

  const responded = new Set();
  for (let i = 1; i < presenceEv.length; i++) {
    if (presenceEv[i][0] && presenceEv[i][1]) responded.add(`${presenceEv[i][0]}_${presenceEv[i][1]}`);
  }

  const joueursSM1 = [];
  for (let i = 1; i < comptes.length; i++) {
    if (rowHasRole(comptes[i], "Joueur") && rowEquipesForRole(comptes[i], "Joueur").indexOf("SM1") !== -1) joueursSM1.push(comptes[i][COL_NOM]);
  }

  upcomingTrainings.forEach(ev => {
    const eventId = ev[0];
    joueursSM1.forEach(nom => {
      const key = `${eventId}_${nom}`;
      if (!responded.has(key)) {
        presenceEvSheet.appendRow([eventId, nom, "Non", "Non renseigné avant l'échéance du dimanche soir (sanction automatique)"]);
        applyCaisseNoireSanction(gridSheet, nom, "Non renseigné avant dimanche soir");
      }
    });
  });
}

// Envoie un rappel par mail aux joueurs SM1 qui n'ont pas encore répondu à un entraînement
// à venir dans les 7 prochains jours — même périmètre exact que checkDisponibilitesDimanche,
// pour que le rappel corresponde toujours à ce qui sera sanctionné.
// mode = "friday" (rappel standard) ou "sunday" (dernier rappel, ton plus urgent).

function sendDisponibilitesReminders(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evenementsSheet = ss.getSheetByName("Evenements");
  const comptesSheet = ss.getSheetByName("Comptes");
  const presenceEvSheet = ss.getSheetByName("PresenceEvenements");
  if (!evenementsSheet || !comptesSheet || !presenceEvSheet) return;
  ensureComptesSchema(comptesSheet); // garantit que la colonne Email existe

  const evenements = evenementsSheet.getDataRange().getValues();
  const comptes = comptesSheet.getDataRange().getValues();
  const presenceEv = presenceEvSheet.getDataRange().getValues();

  const now = new Date();
  const in7Days = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  // Tous les entraînements à venir, toutes équipes confondues (plus limité à la SM1).
  const upcomingTrainings = [];
  for (let i = 1; i < evenements.length; i++) {
    const row = evenements[i];
    if (!row[0] || row[3] !== "Entraînement") continue;
    const equipe = String(row[6] || "SM1").trim();
    const d = new Date(String(row[1]) + "T" + (row[2] || "00:00"));
    if (d >= now && d <= in7Days) upcomingTrainings.push({ id: row[0], titre: row[4] || "Entraînement", date: d, equipe });
  }
  if (upcomingTrainings.length === 0) return;

  const responded = new Set();
  for (let i = 1; i < presenceEv.length; i++) {
    if (presenceEv[i][0] && presenceEv[i][1]) responded.add(`${presenceEv[i][0]}_${presenceEv[i][1]}`);
  }

  // Tous les joueurs, toutes équipes (chacun avec la ou les équipes où il est Joueur).
  const joueurs = [];
  for (let i = 1; i < comptes.length; i++) {
    if (rowHasRole(comptes[i], "Joueur")) {
      rowEquipesForRole(comptes[i], "Joueur").forEach(equipe => {
        joueurs.push({ nom: comptes[i][COL_NOM], email: comptes[i][COL_EMAIL] || "", equipe });
      });
    }
  }
  if (joueurs.length === 0) return;

  const subject = mode === "sunday" ? "⏰ Dernier rappel — Présence à renseigner ce soir" : "Rappel — Présence à renseigner d'ici dimanche";

  joueurs.forEach(j => {
    const pending = upcomingTrainings.filter(t => t.equipe === j.equipe && !responded.has(`${t.id}_${j.nom}`));
    if (pending.length === 0) return;
    if (!j.email) return; // pas d'adresse mail renseignée : rien à envoyer pour cette personne

    // La sanction caisse noire ne concerne que la SM1 — on ne la mentionne pas aux autres équipes.
    const urgencyLine = j.equipe === "SM1"
      ? (mode === "sunday" ? "C'est aujourd'hui le dernier délai : réponds avant minuit pour éviter la sanction." : "Tu as jusqu'à dimanche minuit pour répondre.")
      : (mode === "sunday" ? "C'est aujourd'hui le dernier délai pour répondre, merci de penser à ton coach !" : "Merci de répondre avant dimanche pour que ton coach puisse s'organiser.");
    const sanctionLine = j.equipe === "SM1" ? "Sans réponse, une sanction de 1€ sera automatiquement ajoutée à ta caisse noire.\n\n" : "";

    const liste = pending.map(t => "- " + t.titre + " (" + Utilities.formatDate(t.date, Session.getScriptTimeZone(), "dd/MM 'à' HH:mm") + ")").join("\n");
    const body = "Bonjour " + j.nom + ",\n\n"
      + "Tu n'as pas encore renseigné ta présence pour :\n" + liste + "\n\n"
      + urgencyLine + "\n"
      + sanctionLine
      + "Réponds directement depuis LustuZone.\n\n"
      + "L'équipe " + j.equipe;
    try {
      MailApp.sendEmail(j.email, subject, body, { name: "HBC Bischoffsheim " + j.equipe });
    } catch (err) {
      Logger.log("Erreur envoi mail à " + j.nom + " (" + j.equipe + ") : " + err);
    }
  });
}

// À exécuter manuellement pour tester l'envoi de mail tout de suite, sans dépendre d'un vrai
// entraînement à venir non répondu. Remplace "Maximilien M." par ton propre nom exact si besoin.
function testEmailNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Comptes");
  ensureComptesSchema(sheet);
  const data = sheet.getDataRange().getValues();
  const nomCible = "Maximilien M.";
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_NOM]).trim() === nomCible) {
      const email = data[i][COL_EMAIL];
      if (!email) {
        Logger.log("Aucune adresse mail enregistrée pour ce compte — renseigne-la sur le Profil de l'appli, puis relance ce test.");
        return;
      }
      MailApp.sendEmail(email, "Test LustuZone", "Si tu reçois ce mail, les rappels par mail fonctionnent bien !", { name: "HBC Bischoffsheim SM1" });
      Logger.log("Mail de test envoyé à " + email);
      return;
    }
  }
  Logger.log("Compte introuvable.");
}

function sendDisponibilitesReminderFriday() { sendDisponibilitesReminders("friday"); }
function sendDisponibilitesReminderSunday() { sendDisponibilitesReminders("sunday"); }

// À exécuter UNE FOIS depuis l'éditeur Apps Script pour installer les deux rappels
// (vendredi matin + dimanche matin), en plus du déclencheur de sanction déjà en place.
function installReminderTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "sendDisponibilitesReminderFriday" || fn === "sendDisponibilitesReminderSunday") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("sendDisponibilitesReminderFriday")
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).create();
  ScriptApp.newTrigger("sendDisponibilitesReminderSunday")
    .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();
}

// ===================== SAUVEGARDE AUTOMATIQUE =====================
function getBackupsRootFolder() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty("BACKUPS_ROOT_FOLDER_ID");
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (err) { /* recréé si supprimé */ }
  }
  const folder = DriveApp.createFolder("LustuZone - Sauvegardes");
  props.setProperty("BACKUPS_ROOT_FOLDER_ID", folder.getId());
  return folder;
}

// Copie complète et indépendante de tout le classeur (toutes les feuilles), pour ne perdre
// au pire qu'une semaine de données en cas d'erreur (comme le lancement accidentel de setup()).
// Garde les 12 dernières semaines, supprime automatiquement les plus anciennes.
function backupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const folder = getBackupsRootFolder();
  const name = "Sauvegarde LustuZone — " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  file.makeCopy(name, folder);

  const files = [];
  const iter = folder.getFiles();
  while (iter.hasNext()) files.push(iter.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(12).forEach(f => f.setTrashed(true));
}

// À exécuter UNE FOIS depuis l'éditeur pour installer la sauvegarde automatique hebdomadaire
// (tous les lundis à 4h). Fait aussi une première sauvegarde immédiate.
// ===================== RAPPEL RDV OSTÉO (la veille) =====================
// À exécuter chaque jour (voir installOsteoReminderTrigger) : prévient par mail toute personne
// ayant un RDV ostéo réservé le lendemain.
function sendOsteoReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const slotsSheet = ss.getSheetByName("OsteoSlots");
  const resaSheet = ss.getSheetByName("OsteoReservations");
  const comptesSheet = ss.getSheetByName("Comptes");
  if (!slotsSheet || !resaSheet || !comptesSheet) return;
  ensureComptesSchema(comptesSheet);

  const tz = Session.getScriptTimeZone();
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, tz, "yyyy-MM-dd");

  const slots = slotsSheet.getDataRange().getValues();
  const slotsForTomorrow = {};
  for (let i = 1; i < slots.length; i++) {
    if (String(slots[i][1]).trim() === tomorrowStr) slotsForTomorrow[slots[i][0]] = slots[i];
  }
  if (Object.keys(slotsForTomorrow).length === 0) return;

  const reservations = resaSheet.getDataRange().getValues();
  const comptesData = comptesSheet.getDataRange().getValues();
  function emailFor(nom) {
    for (let i = 1; i < comptesData.length; i++) {
      if (comptesData[i][COL_NOM] === nom) return comptesData[i][COL_EMAIL] || "";
    }
    return "";
  }

  for (let i = 1; i < reservations.length; i++) {
    const slot = slotsForTomorrow[reservations[i][0]];
    if (!slot) continue;
    const nom = reservations[i][1];
    const motif = reservations[i][2] || "";
    const email = emailFor(nom);
    if (!email) continue;
    const heure = slot[2] || "";
    const lieu = slot[3] || "";
    const body = "Bonjour " + nom + ",\n\n"
      + "Petit rappel : tu as rendez-vous avec Eve (ostéopathe du club) demain à " + heure + (lieu ? ", à " + lieu : "") + ".\n\n"
      + (motif ? "Motif indiqué : " + motif + "\n\n" : "")
      + "Besoin d'annuler ? Rends-toi sur l'appli, page RDV Ostéo → Mes RDV.\n\n"
      + "À demain !";
    try {
      MailApp.sendEmail(email, "Rappel : ton RDV ostéo demain" + (heure ? " à " + heure : ""), body, { name: "LustuZone — RDV Ostéo" });
    } catch (err) { /* pas bloquant */ }
  }
}

// À exécuter UNE FOIS depuis l'éditeur pour installer le rappel quotidien (tous les jours à 18h).
function installOsteoReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sendOsteoReminders") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendOsteoReminders").timeBased().everyDays(1).atHour(18).create();
}

function installBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "backupSpreadsheet") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("backupSpreadsheet")
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
  backupSpreadsheet(); // première sauvegarde tout de suite, sans attendre lundi
}

// À exécuter UNE FOIS depuis l'éditeur pour ajouter les comptes parents/joueurs U17 — vérifie
// automatiquement les noms déjà existants (Nom) pour ne jamais créer de doublon, même si
// relancée plusieurs fois par erreur.
// À exécuter UNE FOIS depuis l'éditeur pour créer le compte d'Eve (ostéopathe du club), avec
// le rôle Ostéo (accès Agenda + Actualités de toutes les équipes, et la page RDV Ostéo).
// Vérifie qu'elle n'existe pas déjà, pour ne jamais créer de doublon.
// À exécuter UNE FOIS depuis l'éditeur pour repartir sur une base propre : vide les créneaux et
// réservations ostéo de test, et supprime les actualités "Nouveau(x) créneau(x)..." déjà postées.
// Ne touche à RIEN d'autre (Grid, Comptes, etc. restent intacts).
function resetOsteoTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const slotsSheet = ss.getSheetByName("OsteoSlots");
  if (slotsSheet && slotsSheet.getLastRow() > 1) {
    slotsSheet.getRange(2, 1, slotsSheet.getLastRow() - 1, slotsSheet.getLastColumn()).clearContent();
  }

  const resaSheet = ss.getSheetByName("OsteoReservations");
  if (resaSheet && resaSheet.getLastRow() > 1) {
    resaSheet.getRange(2, 1, resaSheet.getLastRow() - 1, resaSheet.getLastColumn()).clearContent();
  }

  const actualitesSheet = ss.getSheetByName("Actualites");
  if (actualitesSheet) {
    const data = actualitesSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1] || "").indexOf("Nouveau(x) créneau(x) RDV Ostéo") === 0) {
        actualitesSheet.deleteRow(i + 1);
      }
    }
  }

  Logger.log("Nettoyage terminé : créneaux, réservations et actualités de test RDV Ostéo effacés.");
}

function addOsteoAccount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Comptes");
  ensureComptesSchema(sheet);
  const nom = "Eve";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_NOM]).trim() === nom) {
      Logger.log("Le compte '" + nom + "' existe déjà — rien fait.");
      return;
    }
  }
  const row = new Array(8).fill("");
  row[COL_NOM] = nom;
  row[COL_ROLES] = "Ostéo:Toutes";
  row[COL_NOMCOMPLET] = "Eve"; // à compléter avec son nom de famille dans Google Sheets si besoin
  sheet.appendRow(row);
  Logger.log("Compte '" + nom + "' créé avec le rôle Ostéo.");
}

function addU17ParentsAndPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Comptes");
  ensureComptesSchema(sheet);

  const nouveauxComptes = [
  ];

  const data = sheet.getDataRange().getValues();
  const nomsExistants = new Set(data.slice(1).map(r => String(r[COL_NOM]).trim()));

  let ajoutes = 0, ignores = 0;
  nouveauxComptes.forEach(([nom, roles, nomComplet]) => {
    if (nomsExistants.has(nom.trim())) {
      Logger.log("Ignoré (déjà existant) : " + nom);
      ignores++;
      return;
    }
    const row = new Array(8).fill("");
    row[COL_NOM] = nom;
    row[COL_ROLES] = roles;
    row[COL_NOMCOMPLET] = nomComplet;
    sheet.appendRow(row);
    nomsExistants.add(nom.trim());
    ajoutes++;
  });

  Logger.log(ajoutes + " compte(s) ajouté(s), " + ignores + " ignoré(s) car déjà existant(s).");
}


// Incrémente de 1 la case (joueur x action) déjà existante dans Grid — ne modifie pas la structure de la feuille.
function applyCaisseNoireSanction(gridSheet, nom, actionLabel) {
  const data = gridSheet.getDataRange().getValues();
  const header = data[0];
  const playerCol = header.indexOf(nom);
  if (playerCol === -1) return;
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === actionLabel) {
      const currentVal = data[r][playerCol] || 0;
      gridSheet.getRange(r + 1, playerCol + 1).setValue(currentVal + 1);
      return;
    }
  }
}

// Ajoute une nouvelle action à la feuille Grid si elle n'existe pas déjà, sans jamais effacer les données existantes.
function ensureGridAction(label, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Grid");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === label) return; // déjà présente
  }
  const totalRowIdx = data.findIndex(row => row[0] === "TOTAL Joueur");
  const insertAtRow = totalRowIdx === -1 ? sheet.getLastRow() + 1 : totalRowIdx + 1; // ligne 1-indexée, juste avant TOTAL
  if (totalRowIdx !== -1) sheet.insertRowBefore(insertAtRow);
  const nbCols = data[0].length;
  const firstPlayerCol = 3, lastPlayerCol = nbCols - 1, totalCol = nbCols;
  const newRow = new Array(nbCols).fill(0);
  newRow[0] = label;
  newRow[1] = value;
  sheet.getRange(insertAtRow, 1, 1, nbCols).setValues([newRow]);
  const playerRange = `${columnToLetter(firstPlayerCol)}${insertAtRow}:${columnToLetter(lastPlayerCol)}${insertAtRow}`;
  sheet.getRange(insertAtRow, totalCol).setFormula(`=SUM(${playerRange})*B${insertAtRow}`);
}

// ===================== API =====================

// Retourne le détail des rôles d'une personne à partir de sa ligne unique (Nom+Code) —
// un compte multi-rôles cumule automatiquement tous ses droits (parsés depuis la cellule
// "Roles"), sans avoir à choisir une "casquette" à la connexion.
function getSessionRoleDetails(ss, nom, code) {
  const comptes = ss.getSheetByName("Comptes").getDataRange().getValues();
  for (let i = 1; i < comptes.length; i++) {
    if (String(comptes[i][COL_NOM]).trim() === String(nom).trim() && String(comptes[i][COL_CODE]).trim() === String(code).trim()) {
      return parseRoles(comptes[i][COL_ROLES]).map(r => ({ ...r, nom: comptes[i][COL_NOM] }));
    }
  }
  return null;
}

// Renvoie un tableau des rôles (dédupliqués) de la personne, ou null si nom+code invalide.
// Utiliser hasRole(roles, "X") pour vérifier l'appartenance à un rôle donné.
function checkAuth(ss, nom, code) {
  const details = getSessionRoleDetails(ss, nom, code);
  if (!details) return null;
  return [...new Set(details.map(d => d.role))];
}

function hasRole(roles, roleName) {
  return !!roles && roles.indexOf(roleName) !== -1;
}

// Renvoie la liste des équipes où la personne a le rôle donné (ex: hasRole+équipes pour "Coach").
function equipesForRole(ss, nom, code, roleName) {
  const details = getSessionRoleDetails(ss, nom, code);
  if (!details) return [];
  return [...new Set(details.filter(d => d.role === roleName).map(d => d.equipe))];
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = e.parameter.action;

  if (action === "set") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const sheet = ss.getSheetByName("Grid");
    const row = parseInt(e.parameter.row, 10);
    const col = parseInt(e.parameter.col, 10);
    const value = parseInt(e.parameter.value, 10);
    sheet.getRange(row, col).setValue(value);
    return jsonOut({ ok: true });
  }

  if (action === "accountStatus") {
    const nom = e.parameter.nom;
    const comptes = ss.getSheetByName("Comptes").getDataRange().getValues();
    for (let i = 1; i < comptes.length; i++) {
      if (comptes[i][COL_NOM] === nom) {
        const codeVide = !comptes[i][COL_CODE] || String(comptes[i][COL_CODE]).trim() === "";
        return jsonOut({ ok: true, needsSetup: codeVide });
      }
    }
    return jsonOut({ ok: false });
  }

  if (action === "changeCode") {
    const nom = e.parameter.authNom;
    const oldCode = e.parameter.oldCode;
    const role = checkAuth(ss, nom, oldCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const newCode = String(e.parameter.newCode || "");
    if (!/^\d{4}$/.test(newCode)) return jsonOut({ ok: false, error: "format" });
    const sheet = ss.getSheetByName("Comptes");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL_NOM]).trim() === String(nom).trim() && String(data[i][COL_CODE]).trim() === String(oldCode).trim()) {
        sheet.getRange(i + 1, COL_CODE + 1).setNumberFormat("@");
        sheet.getRange(i + 1, COL_CODE + 1).setValue(newCode);
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  if (action === "setCode") {
    const nom = e.parameter.nom;
    const newCode = String(e.parameter.newCode || "");
    if (!/^\d{4}$/.test(newCode)) return jsonOut({ ok: false, error: "format" });
    const sheet = ss.getSheetByName("Comptes");
    ensureComptesSchema(sheet);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL_NOM] === nom) {
        const codeVide = !data[i][COL_CODE] || String(data[i][COL_CODE]).trim() === "";
        if (!codeVide) return jsonOut({ ok: false, error: "already_set" });
        sheet.getRange(i + 1, COL_CODE + 1).setNumberFormat("@");
        sheet.getRange(i + 1, COL_CODE + 1).setValue(newCode);
        const details = getSessionRoleDetails(ss, nom, newCode);
        return jsonOut({ ok: true, roles: details || parseRoles(data[i][COL_ROLES]).map(r => ({ ...r, nom })) });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  if (action === "listNoms") {
    const comptes = ss.getSheetByName("Comptes").getDataRange().getValues();
    const seen = new Set();
    const noms = [];
    for (let i = 1; i < comptes.length; i++) {
      if (comptes[i][COL_NOM] && !seen.has(comptes[i][COL_NOM])) {
        seen.add(comptes[i][COL_NOM]);
        noms.push({ nom: comptes[i][COL_NOM], role: comptes[i][COL_ROLES] });
      }
    }
    return jsonOut({ ok: true, noms });
  }

  if (action === "setPoste") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const nom = e.parameter.nom;
    const isSelfEdit = e.parameter.authNom === nom;
    if (!hasRole(role, "Admin") && !isSelfEdit) {
      return jsonOut({ ok: false, error: "forbidden" });
    }
    const sheet = ss.getSheetByName("Comptes");
    ensureComptesSchema(sheet);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const matches = isSelfEdit
        ? (String(data[i][COL_NOM]).trim() === String(nom).trim() && String(data[i][COL_CODE]).trim() === String(e.parameter.authCode).trim())
        : (String(data[i][COL_NOM]).trim() === String(nom).trim());
      if (matches) {
        sheet.getRange(i + 1, COL_POSTE + 1).setValue(e.parameter.poste || "");
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  if (action === "setEmail") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const nom = e.parameter.nom;
    const isSelfEdit = e.parameter.authNom === nom;
    if (!hasRole(role, "Admin") && !isSelfEdit) {
      return jsonOut({ ok: false, error: "forbidden" });
    }
    const sheet = ss.getSheetByName("Comptes");
    ensureComptesSchema(sheet);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const matches = isSelfEdit
        ? (String(data[i][COL_NOM]).trim() === String(nom).trim() && String(data[i][COL_CODE]).trim() === String(e.parameter.authCode).trim())
        : (String(data[i][COL_NOM]).trim() === String(nom).trim());
      if (matches) {
        sheet.getRange(i + 1, COL_EMAIL + 1).setValue(e.parameter.email || "");
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  // Envoie une question/demande de support directement à l'adresse de gestion du club, et
  // l'enregistre dans un onglet de suivi (visible ensuite par la personne dans l'appli).
  // Prévient tous les Admin ayant une adresse mail renseignée qu'un joueur dit avoir payé sa
  // cotisation — aucune détection automatique du paiement, juste une action à vérifier/valider.
  if (action === "notifyPaymentClaim") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const nom = e.parameter.authNom;
    const montant = e.parameter.montant || "?";
    const comptesSheet = ss.getSheetByName("Comptes");
    ensureComptesSchema(comptesSheet);
    const data = comptesSheet.getDataRange().getValues();
    const adminEmails = [];
    for (let i = 1; i < data.length; i++) {
      if (rowHasRole(data[i], "Admin") && data[i][COL_EMAIL]) adminEmails.push(data[i][COL_EMAIL]);
    }
    const body = nom + " indique avoir réglé sa cotisation (" + montant + " €).\n\n"
      + "Merci de vérifier la réception du paiement puis de l'enregistrer dans LustuZone (page Paiements → \"+ Ajouter un paiement\").";
    [...new Set(adminEmails)].forEach(em => {
      try { MailApp.sendEmail(em, "LustuZone — Paiement à valider : " + nom, body, { name: "LustuZone — Paiements" }); } catch (err) { /* pas bloquant */ }
    });
    return jsonOut({ ok: true });
  }

  if (action === "sendSupportMessage") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const message = (e.parameter.message || "").trim();
    if (!message) return jsonOut({ ok: false, error: "empty" });
    const comptesSheet = ss.getSheetByName("Comptes");
    ensureComptesSchema(comptesSheet);
    const data = comptesSheet.getDataRange().getValues();
    let expediteurEmail = "";
    const adminEmails = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL_NOM] === e.parameter.authNom) expediteurEmail = data[i][COL_EMAIL] || "";
      if (rowHasRole(data[i], "Admin") && data[i][COL_EMAIL]) adminEmails.push(data[i][COL_EMAIL]);
    }
    const body = "Nouveau message depuis LustuZone :\n\n"
      + "De : " + e.parameter.authNom + " (" + role.join(", ") + ")\n"
      + (expediteurEmail ? "Adresse mail renseignée : " + expediteurEmail + "\n" : "")
      + "\nMessage :\n" + message
      + "\n\nRéponds directement depuis l'appli (menu Support, réservé aux Admin).";
    try {
      const mailOptions = { name: "LustuZone — Support" };
      if (expediteurEmail) mailOptions.replyTo = expediteurEmail;
      MailApp.sendEmail("hbcb.gestion@gmail.com", "LustuZone — Question de " + e.parameter.authNom, body, mailOptions);
      // Prévient en plus chaque Admin ayant une adresse mail personnelle renseignée, pour ne pas
      // dépendre uniquement de la surveillance de la boîte mail partagée du club.
      [...new Set(adminEmails)].filter(em => em !== "hbcb.gestion@gmail.com").forEach(em => {
        try { MailApp.sendEmail(em, "LustuZone — Question de " + e.parameter.authNom, body, mailOptions); } catch (err) { /* pas bloquant */ }
      });
    } catch (err) {
      return jsonOut({ ok: false, error: "send_failed" });
    }
    try {
      const supportSheet = ss.getSheetByName("Support");
      setupSupport();
      const id = "s" + Date.now() + "_" + Math.floor(Math.random() * 1000);
      const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      const row = supportSheet.getLastRow() + 1;
      supportSheet.getRange(row, 2).setNumberFormat("@"); // force le texte, sinon Sheets convertit en date réelle
      supportSheet.getRange(row, 1, 1, 5).setValues([[id, now, e.parameter.authNom, message, ""]]);
    } catch (err) {
      Logger.log("Erreur enregistrement suivi support : " + err); // le mail est déjà parti, pas bloquant
    }
    return jsonOut({ ok: true });
  }

  // Historique des demandes de support d'une personne, du plus récent au plus ancien.
// Normalise la date d'une demande de support en texte, que la cellule contienne du texte
// (cas normal) ou que Google Sheets l'ait converti en vraie date (anciennes lignes).
function formatSupportDateValue(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  }
  return String(val || "");
}

  if (action === "getMySupportHistory") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const supportSheet = ss.getSheetByName("Support");
    if (!supportSheet) return jsonOut({ ok: true, history: [] });
    const data = supportSheet.getDataRange().getValues();
    const history = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === e.parameter.authNom) {
        history.push({ date: formatSupportDateValue(data[i][1]), message: data[i][3], reponse: data[i][4] || "" });
      }
    }
    history.reverse();
    return jsonOut({ ok: true, history });
  }

  // Toutes les demandes de support, toutes personnes confondues — réservé à l'Admin.
  if (action === "getAllSupportRequests") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const supportSheet = ss.getSheetByName("Support");
    if (!supportSheet) return jsonOut({ ok: true, requests: [] });
    const data = supportSheet.getDataRange().getValues();
    const requests = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) requests.push({ id: data[i][0], date: formatSupportDateValue(data[i][1]), nom: data[i][2], message: data[i][3], reponse: data[i][4] || "" });
    }
    requests.reverse();
    return jsonOut({ ok: true, requests });
  }

  // Enregistre la réponse de l'Admin à une demande, et prévient l'auteur par mail si possible.
  if (action === "replySupportMessage") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const id = e.parameter.id;
    const reponse = (e.parameter.reponse || "").trim();
    if (!reponse) return jsonOut({ ok: false, error: "empty" });
    const supportSheet = ss.getSheetByName("Support");
    if (!supportSheet) return jsonOut({ ok: false, error: "not_found" });
    const data = supportSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        supportSheet.getRange(i + 1, 5).setValue(reponse);
        const nomAuteur = data[i][2];
        const comptesSheet = ss.getSheetByName("Comptes");
        ensureComptesSchema(comptesSheet);
        const comptesData = comptesSheet.getDataRange().getValues();
        let auteurEmail = "";
        for (let k = 1; k < comptesData.length; k++) {
          if (comptesData[k][COL_NOM] === nomAuteur) { auteurEmail = comptesData[k][COL_EMAIL] || ""; break; }
        }
        if (auteurEmail) {
          try {
            MailApp.sendEmail(auteurEmail, "LustuZone — Réponse à ta question",
              "Bonjour " + nomAuteur + ",\n\nTu as reçu une réponse à ta question :\n\n" + reponse + "\n\nConsulte-la aussi depuis l'appli (menu Support).",
              { name: "LustuZone — Support" });
          } catch (err) { /* pas bloquant */ }
        }
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  if (action === "login") {
    const nom = e.parameter.nom;
    const code = e.parameter.code;
    const details = getSessionRoleDetails(ss, nom, code);
    if (details) return jsonOut({ ok: true, roles: details });
    return jsonOut({ ok: false });
  }

  // Covoiturage : soi-même (joueur adulte, coach), l'Admin, ou un Parent déclaré pour cette
  // personne précise (rôle "Parent:NomEnfant" dans sa cellule Roles) peuvent renseigner.
  // Crée un ou plusieurs créneaux de RDV ostéo (récurrence hebdomadaire possible). Réservé au
  // rôle Ostéo et à l'Admin. Peut aussi publier une actualité générale pour l'annoncer.
  if (action === "addOsteoSlot") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Ostéo") && !hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    setupOsteoSlots();
    const sheet = ss.getSheetByName("OsteoSlots");
    const date = e.parameter.date; // "YYYY-MM-DD"
    const heure = e.parameter.heure || "";
    const lieu = e.parameter.lieu || "";
    const equipe = e.parameter.equipe || "Toutes";
    const semaines = Math.max(1, parseInt(e.parameter.semaines, 10) || 1);
    const recurrentId = semaines > 1 ? ("r" + Date.now()) : "";
    const createdIds = [];

    for (let w = 0; w < semaines; w++) {
      const d = new Date(date + "T00:00:00");
      d.setDate(d.getDate() + w * 7);
      const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const id = "osl" + Date.now() + "_" + w + "_" + Math.floor(Math.random() * 1000);
      const row = sheet.getLastRow() + 1;
      sheet.getRange(row, 2).setNumberFormat("@"); // force le texte, sinon Sheets convertit en date réelle
      sheet.getRange(row, 3).setNumberFormat("@"); // idem pour l'heure, sinon Sheets la convertit en horodatage
      sheet.getRange(row, 1, 1, 6).setValues([[id, dateStr, heure, lieu, equipe, recurrentId]]);
      createdIds.push(id);
    }

    if (e.parameter.publierActualite === "1") {
      try {
        const actualitesSheet = ss.getSheetByName("Actualites");
        const actId = "a" + Date.now();
        const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
        const titre = "Nouveau(x) créneau(x) RDV Ostéo disponible(s)";
        const texte = "Un ou plusieurs créneaux de RDV avec Eve (ostéopathe du club) sont maintenant ouverts à la réservation" + (semaines > 1 ? (", chaque semaine sur " + semaines + " semaines") : "") + ". Rendez-vous sur la page RDV Ostéo pour réserver.";
        actualitesSheet.appendRow([actId, titre, (equipe === "Toutes" ? "Générale" : equipe), texte, e.parameter.authNom, now]);
      } catch (err) {
        Logger.log("Erreur création actualité RDV Ostéo : " + err); // pas bloquant pour la création du créneau, mais visible dans le journal
      }
    }

    return jsonOut({ ok: true, ids: createdIds });
  }

  // Réserve un créneau — pour soi-même, ou pour quelqu'un d'autre si Admin/Ostéo (utile en cas
  // de besoin d'aide). Refuse si le créneau est déjà pris (vérifié côté serveur).
  if (action === "reserveOsteoSlot") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const nom = e.parameter.nom || e.parameter.authNom;
    if (nom !== e.parameter.authNom && !hasRole(role, "Admin") && !hasRole(role, "Ostéo")) {
      return jsonOut({ ok: false, error: "forbidden" });
    }
    setupOsteoReservations();
    const slotId = e.parameter.slotId;
    const motif = e.parameter.motif || "";
    const resaSheet = ss.getSheetByName("OsteoReservations");
    const data = resaSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === slotId) return jsonOut({ ok: false, error: "already_taken" });
    }
    resaSheet.appendRow([slotId, nom, motif]);
    return jsonOut({ ok: true });
  }

  // Annule sa propre réservation — le créneau redevient automatiquement disponible pour les
  // autres, sans aucune action supplémentaire côté Ostéo/Admin.
  if (action === "cancelOsteoReservation") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const nom = e.parameter.nom || e.parameter.authNom;
    if (nom !== e.parameter.authNom && !hasRole(role, "Admin") && !hasRole(role, "Ostéo")) {
      return jsonOut({ ok: false, error: "forbidden" });
    }
    setupOsteoReservations();
    const slotId = e.parameter.slotId;
    const resaSheet = ss.getSheetByName("OsteoReservations");
    const data = resaSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === slotId && data[i][1] === nom) {
        resaSheet.deleteRow(i + 1);
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  // Réservé à Ostéo/Admin : annule la réservation en cours sur un créneau et la réattribue
  // directement à quelqu'un d'autre jugé prioritaire.
  if (action === "reassignOsteoSlotPriority") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Ostéo") && !hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    setupOsteoReservations();
    const slotId = e.parameter.slotId;
    const newNom = e.parameter.newNom;
    const message = (e.parameter.message || "").trim();
    const resaSheet = ss.getSheetByName("OsteoReservations");
    const data = resaSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === slotId) { resaSheet.deleteRow(i + 1); break; }
    }
    resaSheet.appendRow([slotId, newNom, ""]);

    try {
      const slotsSheet = ss.getSheetByName("OsteoSlots");
      const slotsData = slotsSheet.getDataRange().getValues();
      const slot = slotsData.find(r => r[0] === slotId);
      const comptesSheet = ss.getSheetByName("Comptes");
      ensureComptesSchema(comptesSheet);
      const comptesData = comptesSheet.getDataRange().getValues();
      let email = "";
      for (let i = 1; i < comptesData.length; i++) {
        if (comptesData[i][COL_NOM] === newNom) { email = comptesData[i][COL_EMAIL] || ""; break; }
      }
      if (slot && email) {
        const dateStr = slot[1] instanceof Date ? Utilities.formatDate(slot[1], Session.getScriptTimeZone(), "dd/MM/yyyy") : slot[1];
        const heure = slot[2] instanceof Date ? Utilities.formatDate(slot[2], Session.getScriptTimeZone(), "HH:mm") : slot[2];
        const lieu = slot[3] || "";
        const body = "Bonjour " + newNom + ",\n\n"
          + "Je t'ai réservé en priorité un créneau de RDV ostéo le " + dateStr + " à " + heure + (lieu ? ", à " + lieu : "") + ".\n\n"
          + (message ? message + "\n\n" : "")
          + "Tu peux consulter ou annuler ce RDV depuis l'appli, page RDV Ostéo → Mes RDV.\n\n"
          + "À bientôt,\nEve Ostéo";
        MailApp.sendEmail(email, "RDV ostéo prioritaire", body, { name: "Eve Ostéo" });
      }
    } catch (err) {
      Logger.log("Erreur envoi mail réassignation ostéo : " + err); // pas bloquant
    }

    return jsonOut({ ok: true });
  }

  // Réservé à Ostéo/Admin : supprime un créneau entièrement (et sa réservation associée si elle existe).
  if (action === "deleteOsteoSlot") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Ostéo") && !hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    setupOsteoSlots();
    setupOsteoReservations();
    const slotId = e.parameter.slotId;
    const slotsSheet = ss.getSheetByName("OsteoSlots");
    const slotsData = slotsSheet.getDataRange().getValues();
    for (let i = 1; i < slotsData.length; i++) {
      if (slotsData[i][0] === slotId) { slotsSheet.deleteRow(i + 1); break; }
    }
    const resaSheet = ss.getSheetByName("OsteoReservations");
    const resaData = resaSheet.getDataRange().getValues();
    for (let i = 1; i < resaData.length; i++) {
      if (resaData[i][0] === slotId) { resaSheet.deleteRow(i + 1); break; }
    }
    return jsonOut({ ok: true });
  }

  if (action === "setCovoiturage") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const nom = e.parameter.nom;
    let autorise = hasRole(role, "Admin") || e.parameter.authNom === nom;
    if (!autorise) {
      const details = getSessionRoleDetails(ss, e.parameter.authNom, e.parameter.authCode);
      autorise = !!details && details.some(d => d.role === "Parent" && d.equipe === nom);
    }
    if (!autorise) return jsonOut({ ok: false, error: "forbidden" });

    const sheet = ss.getSheetByName("Covoiturage");
    const eventId = e.parameter.eventId;
    const jeConduit = e.parameter.jeConduit || "";
    const places = e.parameter.places || "";
    const besoinPlace = e.parameter.besoinPlace || "";
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === eventId && data[i][1] === nom) {
        sheet.getRange(i + 1, 3, 1, 3).setValues([[jeConduit, places, besoinPlace]]);
        return jsonOut({ ok: true });
      }
    }
    sheet.appendRow([eventId, nom, jeConduit, places, besoinPlace]);
    return jsonOut({ ok: true });
  }

  if (action === "setPresence") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const joueur = e.parameter.joueur;
    // Un joueur ne peut renseigner QUE sa propre présence. Seul un Admin peut le faire pour un autre.
    if (!hasRole(role, "Admin") && e.parameter.authNom !== joueur) {
      return jsonOut({ ok: false, error: "forbidden" });
    }
    const sheet = ss.getSheetByName("Presences");
    const date = e.parameter.date;
    const present = e.parameter.present;
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === date && data[i][1] === joueur) {
        sheet.getRange(i + 1, 3).setValue(present);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([date, joueur, present]);
    return jsonOut({ ok: true });
  }

  if (action === "addPaiement") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const sheet = ss.getSheetByName("Paiements");
    const joueur = e.parameter.joueur;
    const montant = parseFloat(e.parameter.montant);
    const commentaire = e.parameter.commentaire || "";
    const id = "p" + Date.now() + "_" + Math.floor(Math.random() * 1000);
    sheet.appendRow([id, joueur, montant, new Date().toISOString().slice(0, 10), commentaire]);
    return jsonOut({ ok: true, id });
  }

  if (action === "updatePaiement") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const sheet = ss.getSheetByName("Paiements");
    const id = e.parameter.id;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 2).setValue(e.parameter.joueur);
        sheet.getRange(i + 1, 3).setValue(parseFloat(e.parameter.montant));
        sheet.getRange(i + 1, 5).setValue(e.parameter.commentaire || "");
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  if (action === "deletePaiement") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const sheet = ss.getSheetByName("Paiements");
    const id = e.parameter.id;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  if (action === "generateSeasonTrainings") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const sheet = ss.getSheetByName("Evenements");
    const data = sheet.getDataRange().getValues();
    const equipe = e.parameter.equipe || "SM1";
    const existingKeys = new Set();
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === "Entraînement") existingKeys.add(`${data[i][1]}_${data[i][6] || "SM1"}`);
    }

    // Mardi et jeudi ont chacun leur propre date de départ, heure et lieu (équipes/gymnases différents possibles).
    const startMardi = e.parameter.startMardi ? new Date(e.parameter.startMardi) : null;
    const startJeudi = e.parameter.startJeudi ? new Date(e.parameter.startJeudi) : null;
    const end = new Date(e.parameter.end);
    end.setHours(0, 0, 0, 0);
    if (startMardi) startMardi.setHours(0, 0, 0, 0);
    if (startJeudi) startJeudi.setHours(0, 0, 0, 0);
    const heureMardi = e.parameter.heureMardi || "20:00";
    const heureJeudi = e.parameter.heureJeudi || "20:00";
    const lieuMardi = e.parameter.lieuMardi || e.parameter.lieu || "";
    const lieuJeudi = e.parameter.lieuJeudi || e.parameter.lieu || "";

    const earliestStart = [startMardi, startJeudi].filter(Boolean).sort((a, b) => a - b)[0];
    if (!earliestStart) return jsonOut({ ok: false, error: "no_start_date" });

    let created = 0;
    let d = new Date(earliestStart);
    while (d <= end) {
      const dow = d.getDay();
      const isMardi = dow === 2 && startMardi && d >= startMardi;
      const isJeudi = dow === 4 && startJeudi && d >= startJeudi;
      if (isMardi || isJeudi) {
        const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, "0"), dd = String(d.getDate()).padStart(2, "0");
        const dateStr = `${y}-${m}-${dd}`;
        const heure = isMardi ? heureMardi : heureJeudi;
        const lieu = isMardi ? lieuMardi : lieuJeudi;
        const key = `${dateStr}_${equipe}`;
        if (!existingKeys.has(key)) {
          const id = "e" + Date.now() + "_" + Math.floor(Math.random() * 1000000);
          const row = sheet.getLastRow() + 1;
          sheet.getRange(row, 1).setValue(id);
          sheet.getRange(row, 2).setNumberFormat("@").setValue(dateStr);
          sheet.getRange(row, 3).setNumberFormat("@").setValue(heure);
          sheet.getRange(row, 4).setValue("Entraînement");
          sheet.getRange(row, 5).setValue("Entraînement");
          sheet.getRange(row, 6).setValue(lieu);
          sheet.getRange(row, 7).setValue(equipe);
          existingKeys.add(key);
          created++;
        }
      }
      d.setDate(d.getDate() + 1);
    }
    return jsonOut({ ok: true, created });
  }

  if (action === "addEvenement") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Coach") && !hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const sheet = ss.getSheetByName("Evenements");
    ensureEvenementsScoreColumn(sheet);
    const id = "e" + Date.now() + "_" + Math.floor(Math.random() * 1000);
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setValue(id);
    sheet.getRange(row, 2).setNumberFormat("@").setValue(e.parameter.date || "");
    sheet.getRange(row, 3).setNumberFormat("@").setValue(e.parameter.heure || "");
    sheet.getRange(row, 4).setValue(e.parameter.type || "Autre");
    sheet.getRange(row, 5).setValue(e.parameter.titre || "");
    sheet.getRange(row, 6).setValue(e.parameter.lieu || "");
    const equipesCoach = equipesForRole(ss, e.parameter.authNom, e.parameter.authCode, "Coach");
    sheet.getRange(row, 7).setValue(e.parameter.equipe || equipesCoach[0] || "SM1");
    return jsonOut({ ok: true, id });
  }

  if (action === "updateEvenement") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Coach") && !hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const sheet = ss.getSheetByName("Evenements");
    ensureEvenementsScoreColumn(sheet);
    const id = e.parameter.id;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        const row = i + 1;
        sheet.getRange(row, 2).setNumberFormat("@").setValue(e.parameter.date || "");
        sheet.getRange(row, 3).setNumberFormat("@").setValue(e.parameter.heure || "");
        sheet.getRange(row, 4).setValue(e.parameter.type || "Autre");
        sheet.getRange(row, 5).setValue(e.parameter.titre || "");
        sheet.getRange(row, 6).setValue(e.parameter.lieu || "");
        if (Object.prototype.hasOwnProperty.call(e.parameter, "equipe")) {
          sheet.getRange(row, 7).setValue(e.parameter.equipe || "SM1");
        }
        if (Object.prototype.hasOwnProperty.call(e.parameter, "score")) {
          sheet.getRange(row, 8).setValue(e.parameter.score || "");
        }
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  if (action === "deleteEvenement") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Coach") && !hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const sheet = ss.getSheetByName("Evenements");
    const id = e.parameter.id;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  if (action === "setPresenceEvenement") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    if (hasRole(role, "Bénévole") && !hasRole(role, "Coach") && !hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const nom = e.parameter.nom;
    if (!hasRole(role, "Admin") && !hasRole(role, "Coach") && e.parameter.authNom !== nom) {
      return jsonOut({ ok: false, error: "forbidden" });
    }
    const sheet = ss.getSheetByName("PresenceEvenements");
    ensurePresenceEvenementsSchema(sheet);
    const eventId = e.parameter.eventId;
    const present = e.parameter.present;
    const hasJustification = Object.prototype.hasOwnProperty.call(e.parameter, "justification");
    const justification = hasJustification ? e.parameter.justification : null;
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === eventId && data[i][1] === nom) {
        sheet.getRange(i + 1, 3).setValue(present);
        if (hasJustification) sheet.getRange(i + 1, 4).setValue(justification);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([eventId, nom, present, hasJustification ? justification : ""]);
    return jsonOut({ ok: true });
  }

  if (action === "addActualite") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Coach") && !hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const scope = e.parameter.scope || "Générale";
    if (hasRole(role, "Coach") && !hasRole(role, "Admin")) {
      const equipesCoach = equipesForRole(ss, e.parameter.authNom, e.parameter.authCode, "Coach");
      if (equipesCoach.indexOf(scope) === -1) return jsonOut({ ok: false, error: "forbidden_scope" });
    }
    const sheet = ss.getSheetByName("Actualites");
    const id = "a" + Date.now() + "_" + Math.floor(Math.random() * 1000);
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheet.appendRow([id, e.parameter.titre || "", scope, e.parameter.texte || "", e.parameter.authNom, dateStr]);
    return jsonOut({ ok: true, id });
  }

  if (action === "deleteActualite") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!hasRole(role, "Coach") && !hasRole(role, "Admin")) return jsonOut({ ok: false, error: "forbidden" });
    const sheet = ss.getSheetByName("Actualites");
    const data = sheet.getDataRange().getValues();
    const id = e.parameter.id;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        if (hasRole(role, "Coach") && !hasRole(role, "Admin")) {
          const equipesCoach = equipesForRole(ss, e.parameter.authNom, e.parameter.authCode, "Coach");
          if (equipesCoach.indexOf(data[i][2]) === -1) return jsonOut({ ok: false, error: "forbidden_scope" });
        }
        sheet.deleteRow(i + 1);
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ ok: false, error: "not_found" });
  }

  // Liste (et crée si besoin) le dossier photo d'un match précis : Photos/{equipe}/{date titre}.
  // Le dossier n'est créé qu'à la demande (pas pour tous les matchs d'un coup), pour rester rapide.
  if (action === "photosListForMatch") {
    const role = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "auth" });
    const equipe = e.parameter.equipe || "SM1";
    const folderName = matchFolderName(e.parameter.date, e.parameter.titre);
    const root = getPhotosRootFolder();
    const teamFolder = getOrCreateSubfolder(root, equipe);
    const matchFolder = getOrCreateSubfolder(teamFolder, folderName);

    const photos = [];
    const fileIter = matchFolder.getFiles();
    while (fileIter.hasNext()) {
      const file = fileIter.next();
      const mime = file.getMimeType();
      if (mime.indexOf("image/") !== 0) continue;
      photos.push(driveImageInfo(file));
    }
    return jsonOut({ ok: true, folderId: matchFolder.getId(), folderName, photos });
  }

  if (action === "salariesList") {
    const role = checkSalarieAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "forbidden" });
    const root = getSalariesRootFolder();
    const folderId = e.parameter.folderId || root.getId();
    let folder;
    try { folder = DriveApp.getFolderById(folderId); } catch (err) { return jsonOut({ ok: false, error: "not_found" }); }

    const folders = [];
    const fIter = folder.getFolders();
    while (fIter.hasNext()) {
      const f = fIter.next();
      folders.push({ id: f.getId(), name: f.getName() });
    }

    // Note perf : on évite getOwner() (un des appels DriveApp les plus lents) et on ne
    // recalcule plus le fil d'Ariane côté serveur (l'appli le suit elle-même côté client) —
    // ça évite plusieurs allers-retours Drive à chaque navigation.
    const files = [];
    const fileIter = folder.getFiles();
    while (fileIter.hasNext()) {
      const file = fileIter.next();
      const info = driveFileTypeInfo(file);
      files.push({
        id: file.getId(),
        name: info.isLink ? file.getName().replace(/\.url$/, "") : file.getName(),
        iconType: info.iconType,
        isLink: info.isLink,
        linkUrl: info.linkUrl,
        previewUrl: info.previewUrl,
        imageUrl: info.imageUrl,
        viewUrl: info.viewUrl,
        sizeLabel: info.isLink ? "" : formatBytes(file.getSize()),
        date: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
      });
    }

    return jsonOut({ ok: true, currentFolder: { id: folder.getId(), name: folder.getId() === root.getId() ? "Racine" : folder.getName() },
      folders, files });
  }

  if (action === "salariesCreateFolder") {
    const role = checkSalarieAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "forbidden" });
    const root = getSalariesRootFolder();
    const parentId = e.parameter.parentId || root.getId();
    let parent;
    try { parent = DriveApp.getFolderById(parentId); } catch (err) { return jsonOut({ ok: false, error: "not_found" }); }
    const name = (e.parameter.name || "Nouveau dossier").trim();
    const newFolder = parent.createFolder(name);
    return jsonOut({ ok: true, id: newFolder.getId() });
  }

  if (action === "salariesUploadFile") {
    const role = checkSalarieAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "forbidden" });
    const root = getSalariesRootFolder();
    const parentId = e.parameter.parentId || root.getId();
    let parent;
    try { parent = DriveApp.getFolderById(parentId); } catch (err) { return jsonOut({ ok: false, error: "not_found" }); }
    try {
      const bytes = Utilities.base64Decode(e.parameter.base64);
      const blob = Utilities.newBlob(bytes, e.parameter.mimeType || "application/octet-stream", e.parameter.filename || "fichier");
      const file = parent.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return jsonOut({ ok: true, id: file.getId() });
    } catch (err) {
      return jsonOut({ ok: false, error: "upload_failed", detail: String(err) });
    }
  }

  if (action === "salariesAddLink") {
    const role = checkSalarieAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "forbidden" });
    const root = getSalariesRootFolder();
    const parentId = e.parameter.parentId || root.getId();
    let parent;
    try { parent = DriveApp.getFolderById(parentId); } catch (err) { return jsonOut({ ok: false, error: "not_found" }); }
    const titre = (e.parameter.titre || "Lien").trim();
    const url = e.parameter.url || "";
    const file = parent.createFile(titre + ".url", url, MimeType.PLAIN_TEXT);
    file.setDescription("LUSTUZONE_LINK::" + url);
    return jsonOut({ ok: true, id: file.getId() });
  }

  if (action === "salariesDelete") {
    const role = checkSalarieAuth(ss, e.parameter.authNom, e.parameter.authCode);
    if (!role) return jsonOut({ ok: false, error: "forbidden" });
    try {
      if (e.parameter.type === "folder") {
        DriveApp.getFolderById(e.parameter.id).setTrashed(true);
      } else {
        DriveApp.getFileById(e.parameter.id).setTrashed(true);
      }
      return jsonOut({ ok: true });
    } catch (err) {
      return jsonOut({ ok: false, error: "not_found" });
    }
  }

  // action=getAll (ou par défaut) -> tout en un seul appel
  // Sécurité : authentification obligatoire, codes PIN jamais renvoyés, Paiements réservé à l'Admin.
  const callerRole = checkAuth(ss, e.parameter.authNom, e.parameter.authCode);
  if (!callerRole) return jsonOut({ ok: false, error: "auth" });

  const grid = ss.getSheetByName("Grid").getDataRange().getValues();
  const comptesRaw = ss.getSheetByName("Comptes").getDataRange().getValues();
  const comptes = comptesRaw.map((row, i) => {
    if (i === 0) return row; // en-tête
    const copy = row.slice();
    copy[COL_CODE] = ""; // ne jamais renvoyer les codes PIN au client, même le sien
    return copy;
  });
  const presences = ss.getSheetByName("Presences").getDataRange().getValues();
  const paiements = hasRole(callerRole, "Admin") ? ss.getSheetByName("Paiements").getDataRange().getValues() : [];
  const evenements = ss.getSheetByName("Evenements").getDataRange().getValues();
  const presenceEvenements = ss.getSheetByName("PresenceEvenements").getDataRange().getValues();
  const actualitesSheet = ss.getSheetByName("Actualites");
  const actualites = actualitesSheet ? actualitesSheet.getDataRange().getValues() : [];
  const covoiturageSheet = ss.getSheetByName("Covoiturage");
  const covoiturage = covoiturageSheet ? covoiturageSheet.getDataRange().getValues() : [];
  const osteoSlotsSheet = ss.getSheetByName("OsteoSlots");
  const osteoSlotsRaw = osteoSlotsSheet ? osteoSlotsSheet.getDataRange().getValues() : [];
  const osteoSlots = osteoSlotsRaw.map((row, i) => {
    if (i === 0) return row;
    const copy = row.slice();
    if (copy[1] instanceof Date) copy[1] = Utilities.formatDate(copy[1], Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (copy[2] instanceof Date) copy[2] = Utilities.formatDate(copy[2], Session.getScriptTimeZone(), "HH:mm");
    return copy;
  });
  const osteoReservationsSheet = ss.getSheetByName("OsteoReservations");
  const osteoReservations = osteoReservationsSheet ? osteoReservationsSheet.getDataRange().getValues() : [];

  return jsonOut({ ok: true, grid, comptes, presences, paiements, evenements, presenceEvenements, actualites, covoiturage, osteoSlots, osteoReservations });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================== ESPACE SALARIÉS (Google Drive réel) =====================
// Aucune feuille supplémentaire : les dossiers/fichiers vivent directement sur Drive.
// Le dossier racine est créé automatiquement au premier appel, et son ID est
// mémorisé dans les propriétés du script (PropertiesService), pas dans une feuille.

function checkSalarieAuth(ss, nom, code) {
  const role = checkAuth(ss, nom, code);
  if (!hasRole(role, "Salarié") && !hasRole(role, "Admin")) return null;
  return role;
}

function getSalariesRootFolder() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty("SALARIES_ROOT_FOLDER_ID");
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (err) { /* le dossier a été supprimé : on en recrée un */ }
  }
  const folder = DriveApp.createFolder("LustuZone - Espace salariés");
  props.setProperty("SALARIES_ROOT_FOLDER_ID", folder.getId());
  return folder;
}

// ===================== GALERIE PHOTOS (Google Drive réel) =====================
// Même principe que l'Espace Salariés : un dossier racine, puis un sous-dossier par équipe,
// puis un sous-dossier par match (créé à la demande, nommé d'après sa date + son titre).
// L'Admin/Salarié dépose directement les photos dans ces dossiers depuis Google Drive.
function getPhotosRootFolder() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty("PHOTOS_ROOT_FOLDER_ID");
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (err) { /* recréé si supprimé */ }
  }
  const folder = DriveApp.createFolder("LustuZone - Photos");
  props.setProperty("PHOTOS_ROOT_FOLDER_ID", folder.getId());
  return folder;
}

function getOrCreateSubfolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

function matchFolderName(dateStr, titre) {
  const d = String(dateStr || "").slice(0, 10).split("-");
  const dateLabel = d.length === 3 ? `${d[2]}-${d[1]}` : String(dateStr || "");
  return `${dateLabel} ${titre || "Match"}`.trim();
}

function driveImageInfo(file) {
  const id = file.getId();
  return {
    id,
    name: file.getName(),
    imageUrl: `https://lh3.googleusercontent.com/d/${id}=s1200`,
    viewUrl: `https://drive.google.com/file/d/${id}/view`,
  };
}

// À exécuter UNE FOIS manuellement depuis l'éditeur Apps Script (menu déroulant en haut,
// sélectionner cette fonction, "Exécuter") : ça déclenche l'écran d'autorisation Google Drive,
// crée le dossier racine, et affiche son lien dans les journaux d'exécution (Exécutions / Logs).
function initialiserEspaceSalaries() {
  const folder = getSalariesRootFolder();
  Logger.log("Dossier de l'Espace salariés créé : " + folder.getUrl());
}

// Détecte le type d'un fichier Drive pour choisir l'icône et l'URL d'aperçu adaptées.
// Un "lien externe" est représenté par un petit fichier texte dont la description
// commence par LUSTUZONE_LINK:: — pas besoin de feuille séparée pour les stocker.
function driveFileTypeInfo(file) {
  const desc = file.getDescription() || "";
  if (desc.indexOf("LUSTUZONE_LINK::") === 0) {
    return { iconType: "link", isLink: true, linkUrl: desc.substring("LUSTUZONE_LINK::".length), previewUrl: null, imageUrl: null, viewUrl: null };
  }
  const mime = file.getMimeType();
  const id = file.getId();
  let iconType = "file";
  let previewUrl = `https://drive.google.com/file/d/${id}/preview`;
  let imageUrl = null;
  const viewUrl = `https://drive.google.com/file/d/${id}/view?usp=drivesdk`;

  if (mime.indexOf("image/") === 0) {
    iconType = "image";
    // Lien direct vers l'image (pas un iframe Drive) : évite le souci mobile où le navigateur
    // tente d'ouvrir l'appli Google Drive dans l'iframe et n'affiche qu'un logo bloqué.
    imageUrl = `https://lh3.googleusercontent.com/d/${id}=s1600`;
  }
  else if (mime.indexOf("video/") === 0) iconType = "video";
  else if (mime === "application/pdf") iconType = "pdf";
  else if (mime === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || mime === "application/msword" || mime === "application/vnd.google-apps.document") iconType = "docx";
  else if (mime === "application/vnd.openxmlformats-officedocument.presentationml.presentation" || mime === "application/vnd.ms-powerpoint" || mime === "application/vnd.google-apps.presentation") iconType = "pptx";
  else if (mime === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || mime === "application/vnd.ms-excel" || mime === "application/vnd.google-apps.spreadsheet") iconType = "xlsx";

  if (mime === "application/vnd.google-apps.document") previewUrl = `https://docs.google.com/document/d/${id}/preview`;
  else if (mime === "application/vnd.google-apps.presentation") previewUrl = `https://docs.google.com/presentation/d/${id}/embed`;
  else if (mime === "application/vnd.google-apps.spreadsheet") previewUrl = `https://docs.google.com/spreadsheets/d/${id}/preview`;

  return { iconType, isLink: false, linkUrl: null, previewUrl, imageUrl, viewUrl };
}

function formatBytes(bytes) {
  if (!bytes) return "0 Ko";
  if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + " Ko";
  return (bytes / (1024 * 1024)).toFixed(1) + " Mo";
}

// Les requêtes POST (upload de fichier, payload potentiellement volumineux) sont
// simplement redirigées vers la même logique que doGet, via un objet compatible.
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { body = {}; }
  return doGet({ parameter: body });
}
