// ===================================================================
// PRÉSENCES — présence brute par date/joueur (feuille "Presences", peu
// utilisée) et présence par événement avec justification (feuille
// "PresenceEvenements", utilisée par l'Agenda).
// ===================================================================

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

// Ajoute la colonne "Justification" si la feuille existait déjà avant son introduction.
function ensurePresenceEvenementsSchema(sheet) {
  const header = sheet.getRange(1, 1, 1, 4).getValues()[0];
  if (header[3] !== "Justification") {
    sheet.getRange(1, 4).setValue("Justification");
    sheet.getRange(1, 4).setFontWeight("bold");
  }
}

// ===================== ACTIONS API =====================

function api_setPresence(ss, e) {
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

function api_setPresenceEvenement(ss, e) {
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
