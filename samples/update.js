// Author : Pôle Qualité 022 (Aubin Tchoï)

// Lets you update a destination sheet with active selection
function update_onSelec() {
  // Updating sheet sheet with row row
  function update(sheet, row) {
    // You can add a validation here (return; if row doesn't verify a certain set of conditions)
    // Creating a new row with the destination sheet's header's informations
    let rnew = []; heads.forEach(function(el) {rnew.push(row[el]);});
    sheet.getRange((sheet.getLastRow() + 1), 1, 1, rnew.length).setValues([rnew]);
  }
  
  // Displaying a loading screen
  function display_LoadingScreen() {
    let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/loadingScreen.gif" alt="Loading" width="531" height="299">`)
    .setWidth(540)
    .setHeight(350);
    ui.showModelessDialog(htmlLoading, "Synchronisation des données.."); // Not clean bc ui is not defined here but who cares
  }
  
  // -- main --
  const sheetscr = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
      sheetdst = SpreadsheetApp.openById("").getSheetByName(""), // Insert destination spreadsheet ID and sheet name here
      ui = SpreadsheetApp.getUi();
  
  // Confirm selection
  let confirm_selection = ui.alert("Synchronisation des données", "Vous devez préalablement sélectionner la ligne à synchroniser (par exemple en cliquant sur le numéro à gauche). \n Confirmez-vous votre sélection ?", ui.ButtonSet.YES_NO);
  if (confirm_selection == ui.Button.NO) {return;}
  
  // Loading screen
  display_LoadingScreen();
  
  // heads is the sheet's header, data a 2D-array representation of the selected values, and obj its 'array of js object' representation
  let heads = sheetscr.getRange(1, 1, 1, sheetscr.getLastColumn()).getValues().shift(),
      data = sheetscr.getActiveRange().getDisplayValues(),
      obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  
  // Checking the selection again (empty or not)
  if (obj.length == 0) {ui.alert("Pas de ligne sélectionnée, veuillez recommencer."); return;}
  
  // Synchronizing each selected row
  obj.forEach(function(row) {
    update(sheetdst, row);
  });
  
  // Confirmation
  let img_url = "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/thumbsUp.png",
      sheet_url = "",
      operation_success = HtmlService
  .createHtmlOutput(`<span style='font-size: 16pt;'> <span style="font-family: 'roboto', sans-serif;">Le <a href="${sheet_url}">sheets</a> a été complété avec succès.<br/><br/> La bise</span></span><p style="text-align:center;"><img src="${img_url}" alt="Thumbs up" width="130" height="131"></p>`);
  ui.showModalDialog(operation_success, "Synchronisation des données");
}

