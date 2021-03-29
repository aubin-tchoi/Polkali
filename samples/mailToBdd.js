/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)  */

/* Quand un mail arrive dans la boîte mail,
- on extrait ses pièces jointes et on met le mail dans la corbeille   (retrieveMailAttachments)
- on enregistres ces pièces jointes sur le Drive                      (saveOnDrive)
- on les ouvre au format Google Docs                                  (openAsDocs)
- on en extrait les données utile pour la BDD                         (readData)
- on enregistre ces données dans la BDD                               (writeOnSheet)
- on supprime les pièces jointes                                      (deleteFiles)                 */

/* La fonction readData est ainsi à adapter en fonction du formatage des mails à lire,
    On peut également choisir de lire des infos directement dans le mail plutôt que les pièces jointes en adaptant ce script
    (on remplace l'extraction, l'enregistrement et la lecture des pièces jointes par une lecture du message html du mail). */

const FOLDER_ID = "", // Where you save the attachments (knowing that they will be deleted at the end of the process)
    SPREADSHEET_ID = "", // The id of the database
    SHEET_NAME = ""; // The name of the sheet containing the database

function addToBdd() {
    function retrieveMailAttachments() {
        // Reading the most recent message in the most recent thread
        const thread = GmailApp.getInboxThreads()[0],
            attachments = thread.getMessages()[0].getAttachments();
        thread.moveToTrash();
        return attachments.map(attachment => attachment.copyBlob());
    }

    function saveOnDrive(attachments, folder_id) {
        let folder = DriveApp.getFolderById(folder_id),
            fileIDs = [];
        attachments.forEach(function (attachment) {
            let file = folder.createFile(attachment);
            fileIDs.push(file.getId());
        });
        return fileIDs;
    }

    function openAsDocs(fileIDs) {
        let docFiles = [];
        fileIDs.forEach(function (fileID) {
            docFiles.push(DocumentApp.openById(fileID));
        });
        return docFiles;
    }

    // Here is the part that you should adapt depending on the format of the mail
    function readData(docFiles) {
        let data = [];
        docFiles.forEach();
        return data;
    }

    function writeOnSheet(spreadsheetId, sheetName, data) {
        const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
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

    let attachments = retrieveMailAttachments(),
        fileIDs = saveOnDrive(attachments, FOLDER_ID),
        docFiles = openAsDocs(fileIDs),
        data = readData(docFiles);
    writeOnSheet(SPREADSHEET_ID, SHEET_NAME, data);
    deleteFiles(FOLDER_ID);
}
