// ===================================================================
// PAIEMENTS — suivi des cotisations réglées (feuille "Paiements",
// réservée à l'Admin).
// ===================================================================

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

// ===================== ACTIONS API =====================

function api_addPaiement(ss, e) {
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

function api_updatePaiement(ss, e) {
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

function api_deletePaiement(ss, e) {
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

// Prévient tous les Admin ayant une adresse mail renseignée qu'un joueur dit avoir payé sa
// cotisation — aucune détection automatique du paiement, juste une action à vérifier/valider.
function api_notifyPaymentClaim(ss, e) {
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
