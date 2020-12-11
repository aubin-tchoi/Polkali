/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)  */

/* Quand un mail arrive dans la boîte mail,
- on extrait ses pièces jointes et on met le mail dans la corbeille   (retrieveMailAttachments)
- on enregistres ces pièces jointes sur le Drive                      (saveOnDrive)
- on les ouvre au format Google Docs                                  (openAsDocs)
- on en extrait les données utile pour la BDD                         (readData)
- on enregistre ces données dans la BDD                               (writeOnSheet)
- on supprime les pièces jointes                                      (deleteFiles)                 */

function addToBdd() {
  
  function retrieveMailAttachments() {
    const thread = GmailApp.getInboxThreads()[0],
      attachments = thread.getMessages()[0].getAttachments();
    thread.moveToTrash();
    return attachments.map(attachment => attachment.copyBlob());
  }

  function saveOnDrive(attachments, folder_id) {
    let folder = DriveApp.getFolderById(folder_id),
      fileIDs = [];
    attachments.forEach(function(attachment) {let file = folder.createFile(attachment); fileIDs.push(file.getId());});
    return fileIDs;
  }

  function openAsDocs(fileIDs) {
    let docFiles = [];
    fileIDs.forEach(function(fileID) {docFiles.push(DocumentApp.openById(fileID));});
    return docFiles;
  }

  function readData(docFiles) {
    let data = [];
    docFiles.forEach();
    return data;
  }

  function writeOnSheet(sheet_id, data) {
    const sheet = SpreadsheetApp.openById(sheet_id).getSheets()[0];
    sheet.getRange((sheet.getLastRow() + 1), 1, data.length(), data[0].length()).setValues(data);
  }

  function deleteFiles(folder_id) {
    let folder = DriveApp.getFolderById(folder_id),
      files = folder.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      file.setTrashed(true);
    }
  }
  
  const folder_id = "",
    sheet_id = "";

  let attachments = retrieveMailAttachments(),
    fileIDs = saveOnDrive(attachments, folder_id),
    docFiles = openAsDocs(fileIDs),
    data = readData(docFiles);
  writeOnSheet(sheet_id, data);
  deleteFiles(folder_id);
}
