// ===================================================================
// SETUP — point d'entrée unique pour initialiser toutes les feuilles.
// À exécuter UNE FOIS depuis l'éditeur (sélectionner "setup" dans le
// menu déroulant en haut, puis "Exécuter") pour un nouveau classeur.
// Sur un classeur existant, relancer setup() ne réinitialise que Grid
// (voir setupGrid) ; les autres setupXxx() n'écrivent rien si la
// feuille contient déjà des données.
//
// Pour un club qui ne veut pas un module donné (ex. Ostéo), retirer
// simplement l'appel setupXxx() correspondant ci-dessous, en plus de
// supprimer le fichier .gs du module.
// ===================================================================

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
  ensureGridAction("Non renseigné avant dimanche soir", 1); // lié à Notifications.gs (checkDisponibilitesDimanche)
}

// ===================== DONNÉES DE DÉPART SPÉCIFIQUES À CE CLUB =====================
// À exécuter UNE FOIS depuis l'éditeur pour ajouter les comptes parents/joueurs U17 — vérifie
// automatiquement les noms déjà existants (Nom) pour ne jamais créer de doublon, même si
// relancée plusieurs fois par erreur.
// Entièrement propre à l'effectif U17 du club Bischo : à supprimer (ou remplacer par votre
// propre liste) pour une autre association.
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
