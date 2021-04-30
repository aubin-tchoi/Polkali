/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

/** Main function */
function duplication(IDs) {
  let sourceFolder = DriveApp.getFolderById(IDs.source),
    targetFolder = DriveApp.getFolderById(IDs.target);
  duplicateFolder(sourceFolder, targetFolder);
}

// Including a file
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// GET method
function doGet(e) {
  return HtmlService.createTemplateFromFile("home").evaluate();
}

// POST method
function doPost(e) {
return HtmlService.createTemplateFromFile("home").evaluate();
}
