// ===================================================================
// GALERIE PHOTOS (Google Drive réel) — Même principe que l'Espace
// Salariés : un dossier racine, puis un sous-dossier par équipe, puis
// un sous-dossier par match (créé à la demande, nommé d'après sa date
// + son titre). L'Admin/Salarié dépose directement les photos dans ces
// dossiers depuis Google Drive.
// ===================================================================

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

// ===================== ACTION API =====================

// Liste (et crée si besoin) le dossier photo d'un match précis : Photos/{equipe}/{date titre}.
// Le dossier n'est créé qu'à la demande (pas pour tous les matchs d'un coup), pour rester rapide.
function api_photosListForMatch(ss, e) {
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
