/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Uses function displayLoadingScreen taken from ressources.js

const SOURCE = {
    ssId: "",
    sheetName: ""
  },
  DESTINATION = {
    ssId: "",
    sheetName: "",
    url: ""
  },
  IMAGES = {
    thumbsUp: "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/thumbsUp.png"
  };


// Update a destination sheet with active selection, uses 
function updateOnSelec() {
  // Updating sheet sheet with row row
  function update(sheet, row) {
    // You can add a validation here (return; if row doesn't verify a certain set of conditions)
    // Creating a new row with the destination sheet's header's informations
    let rnew = [],
      heads = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().shift();
    heads.forEach(function (el) {
      rnew.push(row[el]);
    });
    sheet.getRange((sheet.getLastRow() + 1), 1, 1, rnew.length).setValues([rnew]);
  }

  // -- main --
  const sheetscr = SpreadsheetApp.openById(SOURCE["ssId"]).getSheetByName(SOURCE["sheetName"]),
    sheetdst = SpreadsheetApp.openById(DESTINATION["ssId"]).getSheetByName(DESTINATION["sheetName"]),
    ui = SpreadsheetApp.getUi();

  // Confirm selection
  let confirmSelection = ui.alert("Synchronisation des données", "Vous devez préalablement sélectionner la ligne à synchroniser (par exemple en cliquant sur le numéro à gauche). \n Confirmez-vous votre sélection ?", ui.ButtonSet.YES_NO);
  if (confirmSelection == ui.Button.NO) {
    return;
  }

  // Loading screen
  displayLoadingScreen();

  // heads is the sheet's header, data a 2D-array representation of the selected values, and obj its 'array of js objects' representation
  let heads = sheetscr.getRange(1, 1, 1, sheetscr.getLastColumn()).getValues().shift(),
    data = sheetscr.getActiveRange().getDisplayValues(),
    obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Checking the selection again (empty or not)
  if (obj.length == 0) {
    ui.alert("Pas de ligne sélectionnée, veuillez recommencer.");
    return;
  }

  // Synchronizing each selected row
  obj.forEach(function (row) {
    update(sheetdst, row);
  });

  // Confirmation
  let operationSuccess = HtmlService
    .createHtmlOutput(`<span style='font-size: 16pt;'> <span style="font-family: 'roboto', sans-serif;">Le <a href="${DESTINATION["url"]}">sheets</a> a été complété avec succès.<br/><br/> La bise</span></span><p style="text-align:center;"><img src="${IMAGES["thumbsUp"]}" alt="Thumbs up" width="130" height="131"></p>`);
  ui.showModalDialog(operationSuccess, "Synchronisation des données");
}
