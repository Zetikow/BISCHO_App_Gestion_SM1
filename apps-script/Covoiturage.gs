// ===================================================================
// COVOITURAGE — une ligne par (événement, personne). feuille "Covoiturage".
// ===================================================================

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

// ===================== ACTION API =====================

// Covoiturage : soi-même (joueur adulte, coach), l'Admin, ou un Parent déclaré pour cette
// personne précise (rôle "Parent:NomEnfant" dans sa cellule Roles) peuvent renseigner.
function api_setCovoiturage(ss, e) {
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
