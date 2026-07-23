// ===================================================================
// ESPACE SALARIÉS (Google Drive réel) — Aucune feuille supplémentaire :
// les dossiers/fichiers vivent directement sur Drive. Le dossier racine
// est créé automatiquement au premier appel, et son ID est mémorisé
// dans les propriétés du script (PropertiesService), pas dans une feuille.
// Module optionnel : à supprimer avec son fichier frontend
// js/modules/salaries.js pour un club qui n'a pas de salarié.
// ===================================================================

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

// ===================== ACTIONS API =====================

function api_salariesList(ss, e) {
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

function api_salariesCreateFolder(ss, e) {
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

function api_salariesUploadFile(ss, e) {
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

function api_salariesAddLink(ss, e) {
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

function api_salariesDelete(ss, e) {
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
