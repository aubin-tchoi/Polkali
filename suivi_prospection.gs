// Author : Pôle Qualité 022 (Aubin Tchoï)

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Synchronisation')
  .addItem("Vers le suivi des études", "sync_onSelec")
  .addToUi();
}

function sync_onSelec() {
  // Returning the last row of a sheet (not possible for this particular sheet for some unkown reasons)
  function manually_getLastRow(sheet) {
    let data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    for (let row=1;row<sheet.getLastRow();row++) {
      if (data[row][0] == "" || row > 200) {
        return row;
      }
    }      
  }
  
  // Updagint sheet sheet with row row
  function update(sheet, row) {
    // Checking whether PEP actually did obtain the mission or not
    if (row["État"] != "Etude obtenue ") {
      ui.alert("Entrée invalide", `L'étude confiée par l'entreprise ${row["Entreprise"]} ne correspond pas à une étude obtenue.`, ui.ButtonSet.OK);
      return;
    }
    
    // ref is gonna be the mission's reference, obtained by incrementing the most recent one
    const heads = sheet.getRange(1, 2, 1, (sheet.getLastColumn() - 1)).getValues().shift(),
        last_row = manually_getLastRow(sheet),
        ref = `'20e${Math.floor(sheet.getRange(last_row, 2, 1, 1).getValues()[0][0].toString().match(/([0-9]+)/gi)[1]) + 1}`;
    
    // Creating a new row with the destination sheet's header's informations
    let rnew = [row["Entreprise"]]; heads.forEach(function(el) {rnew.push(el == "Référence de l'étude" ? ref : (el == "État") ? "Devis accepté" : row[el]);});
    sheet.getRange((last_row + 1), 1, 1, rnew.length).setValues([rnew]);
  }
  
  // Displaying a loading screen
  function display_LoadingScreen() {
    let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://www.demilked.com/magazine/wp-content/uploads/2016/06/gif-animations-replace-loading-screen-14.gif" alt="Loading" width="531" height="299">`)
    .setWidth(540)
    .setHeight(350);
    ui.showModelessDialog(htmlLoading, "Synchronisation des données.."); // Not clean bc ui is not defined here but who cares
  }
  
  // -- MAIN --
  const sheetscr = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
      sheetdst = SpreadsheetApp.openById("1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04").getSheetByName("Suivi"),
      ui = SpreadsheetApp.getUi(),
      htmlLoading = HtmlService
  .createHtmlOutput(`<img src="https://www.demilked.com/magazine/wp-content/uploads/2016/06/gif-animations-replace-loading-screen-14.gif" alt="Loading" width="531" height="299">`)
  .setWidth(540)
  .setHeight(350);
  
  // Confirm selection
  let confirm_selection = ui.alert("Synchronisation des données", "Avez-vous bien sélectionné les lignes à synchroniser ?", ui.ButtonSet.YES_NO);
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
  let img_url = "https://fcbk.su/_data/stickers/tonton_friends_returns/tonton_friends_returns_06.png",
      sheet_url = "https://docs.google.com/spreadsheets/d/1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04/edit#gid=0",
      operation_success = HtmlService
  .createHtmlOutput(`<span style='font-size: 16pt;'> <span style="font-family: 'roboto', sans-serif;">L'opération a été effectuée avec succès, veuillez remplir manuellement le nom de l'étude dans le <a href="${sheet_url}">suivi des études</a>.<br/><br/> La bise</span></span><p style="text-align:center;"><img src="${img_url}" alt="C'est la PEP qui régale !" width="130" height="131"></p>`);
  ui.showModalDialog(operation_success, "Synchronisation des données");
}


