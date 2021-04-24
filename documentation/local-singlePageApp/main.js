/**
 * Webapp to host some documentation on how KPI are implemented
 */

let Route = {};
Route.path = function(route, callback) {
  Route[route] = callback;
}

// GET method
function doGet(e) {
  Logger.log(e);
  Route.path("accueil", () => HtmlService.createTemplateFromFile("accueil").evaluate());
  ["dataTables", "features", "main", "charts", "parameters", "tools"].forEach(file => {
    Route.path(file, load(file));
  })
  
  if (Route[e.parameters.v]) {
    return Route[e.parameters.v]();
  }
  else {
    return HtmlService.createTemplateFromFile("accueil").evaluate();
  }
}

// POST method
function doPost(e) {
  Logger.log(e);
  return HtmlService.createHtmlOutputFromFile("accueil");
}

// Including a file
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function load(page) {
  return HtmlService.createTemplateFromFile(`files/${page}`).evaluate();
}

function redirection(page) {
  doGet({parameters: {v: page}});
}
