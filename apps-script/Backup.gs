// ===================================================================
// SAUVEGARDE AUTOMATIQUE — copie hebdomadaire complète du classeur.
// ===================================================================

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
function installBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "backupSpreadsheet") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("backupSpreadsheet")
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
  backupSpreadsheet(); // première sauvegarde tout de suite, sans attendre lundi
}
