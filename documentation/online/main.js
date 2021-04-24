/**
 * Webapp to host some documentation on how KPI are implemented
 */

// GET method
function doGet(e) {
    return HtmlService.createTemplateFromFile("accueil").evaluate();
}

// POST method
function doPost(e) {
  return HtmlService.createTemplateFromFile("accueil").evaluate();
}

// Including a file
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
