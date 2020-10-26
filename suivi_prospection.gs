//  ----- main -----

// Auteur : Aubin Tchoï
//          Directeur Qualité - 022


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Synchronisation')
  .addItem("Suivi des études", "sync_onselec")
  .addToUi();
}

function sync_onselec() {
  
  function update(sheet, row) {
    const heads = sheet.getRange(1, 2, 1, (sheet.getLastColumn() - 1)).getValues().shift(),
        ref = `20e${Math.floor(sheet.getRange(sheet.getLastRow(), 2, 1, 1).getValues()[0][0].toString().match(/e([0-9]+)/gi)[1]) + 1}`;
    let rnew = [];
    heads.forEach(function(el) {rnew.push(row[el]);});
  }
  
  const sheetscr = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
      sheetdst = SpreadsheetApp.openById("1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04").getSheetByName("Suivi");
  
  let heads = sheetscr.getRange(1, 1, 1, sheet1.getLastColumn()).getValues().shift(),
      data = sheetscr.getActiveRange().getDisplayValues(),
      obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  
  if (obj.length == 0) {
    err = Browser.msgBox("Pas de ligne sélectionnée", Browser.Buttons.OK);
    return;
  }
  
  obj.forEach(function(row) {
    update(sheetdst, row);
  });
}

function test() {
  let obj = {"a":2};
  Logger.log(obj["b"]);
  Logger.log("20e35".match(/[0-9]+e([0-9]+)/gi)[1]);
}
