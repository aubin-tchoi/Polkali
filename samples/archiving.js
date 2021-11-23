/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Making a copy of current spreadsheet in a given folder for archiving or backup-ing purposes
function archiving() {
  const ssOld = SpreadsheetApp.getActiveSpreadsheet(),
    ssNew = SpreadsheetApp.create(ssOld.getName()),
    ui = SpreadsheetApp.getUi();

  // Copying ssOld's content to ssNew
  ssOld.getSheets().forEach(function (s) {
    s.copyTo(ssNew);
  });

  while (true) {
    try {
    // Retrieving destination folder ID
    let drive_id = ui.prompt("Archivage",
      "Entrez l'ID du dossier Drive de destination.",
      ui.ButtonSet.OK_CANCEL).getResponseText();

    // Moving ssNew to the folder with drive_id
    var file = DriveApp.getFileById(ssNew.getId());
    DriveApp.getFolderById(drive_id).addFile(file);
    break;
    }
    catch(e) {
      ui.alert("L'ID entré est invalide, veuillez recommencer.");
    }
  }

  // Removing initial sheet in ssNew and renaming its shee
  ssNew.deleteSheet(ssNew.getSheetByName("Feuille 1"));
  ssNew.getSheets().forEach(function (s) {
    s.setName(s.getName().match(/Copie de (.+)/)[1]);
  });

  // Removing ssOld if asked to
  let remove = ui.alert("Archivage",
    "Le fichier a été archivé dans le dossier, souhaitez-vous effacer son contenu ?",
    ui.ButtonSet.YES_NO);
  if (remove == ui.Button.YES) {
    ssOld.getSheets().forEach(function (s) {
      s.getRange(2, 1, (s.getLastRow() - 1), s.getLastColumn())
        .clearContent()
        .setBackground('white');
    });
  }
}

//Bonnjoueueerrre