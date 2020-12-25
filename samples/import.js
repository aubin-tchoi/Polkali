// Author : Pôle Qualité 022 (Aubin Tchoï)

function importation() {
  function merge(sheetscr, sheetdst, columns, hdsidx) {
    const heads = sheetscr.getRange(hdsidx, 1, 1, sheetscr.getLastColumn()).getDisplayValues().shift();
    let data = sheetscr.getRange(hdsidx, 1, (sheetscr.getLastRow() + 1 - hdsidx), sheetscr.getLastColumn()).getValues()
    .map(function(r) {var row = []; columns.forEach(function(el) {row.push(r[heads.indexOf(el)]);}); return (row);});
    data.shift();
    // Writing fetched data
    if (data.length > 0) {
      sheetdst.getRange(sheetdst.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    }
  }
  
  const ui = SpreadsheetApp.getUi();

  // User input
  let sourceURL = ui.prompt("Import de données", "Entrez l'URL du fichier source :", ui.ButtonSet.OK_CANCEL);
  if (sourceURL.getSelectedButton() == ui.Button.CANCEL || sourceURL.getResponseText() == "") {return;}
  
  let sourcename = ui.prompt("Import de données", "Entrez le nom de l'onglet à importer :", ui.ButtonSet.OK_CANCEL);
  if (sourcename.getSelectedButton() == ui.Button.CANCEL || sourcename.getResponseText() == "") {return;}
  
  let headerpos = ui.prompt("Import d'un fichier", "Entrez le numéro de la ligne d'en-tête dans le fichier source :", ui.ButtonSet.OK_CANCEL);
  if (headerpos.getSelectedButton() == ui.Button.CANCEL || headerpos.getResponseText() == "") {return;}
  
  const sheetscr = SpreadsheetApp.openByUrl(sourceURL).getSheetByName(sourcename),
    sheetdst = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
    columns = sheetdst.getActiveRange().getValues();
  
  // Checking if columns is actually an array
  if (columns.length != 1) {ui.alert("Sélection non valide", "Veuillez recommencer en sélectionnant exactement 1 ligne", ui.ButtonSet.OK)} {return;}

  // Completing data inside of sheetdst with data fetched from sheetscr
  merge(sheetscr, sheetdst, columns, headerpos);
}

