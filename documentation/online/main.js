/**
 * Webapp to host some documentation on how KPI are implemented
 */

// GET method
function doGet(e) {
    return HtmlService.createTemplateFromFile("layout").evaluate();
}

// POST method
function doPost(e) {
  return HtmlService.createTemplateFromFile("layout").evaluate();
}

// Including a file
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function generateSlides() {
  let KPI = Aubin.generateKPI();
  KPI.slides();
}

function sendMail(mailAddress) {
  let KPI = Aubin.generateKPI();
  KPI.mail(mailAddress);
}
