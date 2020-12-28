/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

/* Must be bound to a spreadsheet that contains a sheet name "Dashboard" in which are given a list of mail adresses
  starting at cell B5 (1 row per adress). Cells B4 to E4 should contain this header : ["Date d'envoi", "Prénom", "Nom", "Adresse mail"]
  This sheet must contain a cell "Poule" next to which will be recorded the votes.
  The name of the Gmail Draft used is read from cell B1 (because it is next to a cell "Modèle Gmail", you can move both cells if you want) */

/* There are only 3 consts variables : folderId, spreadsheetId and colors, the folder to this Id must contain 1 folder for each group. */

const folderId = "1nHfPZR10ZCFx-Nro546WjRwAmih2_bfq",
  spreadsheetId = "11gnPSgy85927z95yxoXb_iS6B2YNoTZumkQooWndjgE",
  colors = {
    winner: "#00ff00",
    duplicate: "#ea9999",
    loser: "#d9ead3",
    groups: "#ffe599",
    contestants: "#fff2cc"
  };

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bracket')
    .addItem("Phases de poules", "firstRound")
    .addSeparator()
    .addItem("Match nul", "tie")
    .addSeparator()
    .addItem("Round suivant", "nextRound")
    .addToUi();
}

// Generates the bracket
function firstRound() {
  // Adds a line to the dashboard
  function formatSheet(spreadsheet, poulNum, poulSize, marker) {
    const sheet = spreadsheet.getSheetByName("Dashboard"),
      [startRow, startColumn] = detectPos(sheet, marker);

    Logger.log(`Row offset : ${startRow}`);
    Logger.log(`Column offset : ${startColumn}`);

    // First line
    for (let poule = 1; poule <= poulNum; poule++) {
      let poulValues = sheet.getRange(startRow, (startColumn + 1 + (poule - 1) * poulSize), 1, poulSize).getValues();
      poulValues[0][0] = `Poule ${poule}`;
      sheet.getRange(startRow, (startColumn + 1 + (poule - 1) * poulSize), 1, poulSize)
        .setValues(poulValues)
        .setBackground(colors["groups"])
        .setHorizontalAlignment("center")
        .merge();
      sheet.getRange((startRow + 1), (startColumn + 1 + (poule - 1) * poulSize), 1, poulSize)
        .setValues([Array.from(Array(poulSize).keys()).map(x => x + 1)])
        .setBackground(colors["contestants"])
        .setHorizontalAlignment("center");
      sheet.setColumnWidths((startColumn + 1), poulNum * poulSize, 33);
    }
    return [startRow, startColumn];
  }

  displayLoadingScreen("Génération du Forms ..")

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId),
    [editLink, publishedLink, poulNum, poulSize] = genForms(folderId, spreadsheet, () => true),
    [startRow, startColumn] = formatSheet(spreadsheet, poulNum, poulSize, "Poule"),
    ui = SpreadsheetApp.getUi();
  let nSent = 0;

  if (ui.alert("Bracket", "Souhaitez-vous envoyer le lien du Google Form par mail aux participants ?", ui.ButtonSet.YES_NO) == ui.Button.YES) {
    // Loading screen
    displayLoadingScreen("Envoi du lien ..")
    nSent = mailLink(spreadsheet, publishedLink);
  }

  // Adding a trigger to the dashboard to update it after each vote
  ScriptApp.newTrigger("updateFirstRound")
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();

  // Confirmation message (also ends the loading screen)
  let htmlOutput = HtmlService.createHtmlOutput(`<span style="font-family: 'trebuchet ms', sans-serif;">Voci le lien éditeur : <a href = "${editLink}">éditer le Form</a>.<br/>
                                                <br/> Voici le lien lecteur : <a href = "${publishedLink}">répondre au Form</a>.<br/>
                                                <br/> Le lien lecteur a été envoyé à ${nSent} personne${nSent >= 2 ? "s" : ""}.</span>`);
  ui.showModelessDialog(htmlOutput, "Succès de l'opération");
}

// Solve tie
function tie() {
  
}

// Generates a form for the next round (supposing that there is no tie)
function nextRound() {
  // Finds the two winners in each group and returning the corresponding indexes
  function findWinners(targetRange, poulNum, poulSize) {
    let targetValues = targetRange.getDisplayValues().shift(),
      big1 = Array.from(Array(poulNum).keys()).map(poul => Math.max(...(targetValues.slice(poul * poulSize, (poul + 1) * poulSize)))),
      big2 = Array.from(Array(poulNum).keys()).map(poul => Math.max(...(targetValues.slice(poul * poulSize, (poul + 1) * poulSize).filter(value => value) != big1[poul]))),
      bigIdx1 = big1.map((value, index) => targetValues.slice(index * poulSize, (index + 1) * poulSize).indexOf(value)),
      bigIdx2 = big2.map((value, index) => targetValues.slice(index * poulSize, (index + 1) * poulSize).indexOf(value));

    Logger.log(`Biggest values for each group : ${big1}`);
    Logger.log(`Second biggest values for each group : ${big2}`);

    return [bigIdx1, bigIdx2];
  }

  const sheet = SpreadsheetApp.openById(spreadsheetId),
    folder = DriveApp.getFolderById(folderId),
    ui = SpreadsheetApp.getUi(),
    [startRow, startColumn] = detectPos(sheet, "Poule"),
    poulNum = Math.max(...sheet.getRange(startRow, (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0].map(el => parseInt(el.replace(/[^0-9]+/gi, ""), 10) || 0)),
    poulSize = Math.max(...sheet.getRange((startRow + 1), (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0]),
    [idx1, idx2] = findWinners(sheet.getRange(sheet.getLastRow(), (startColumn + 1), 1, poulNum * poulSize), poulNum, poulSize);

  displayLoadingScreen("Génération du Forms ..")

  let filter = (idx => idx1.includes(idx) || idx2.includes(idx)),
    [editLink, publishedLink, poulNumSmall, poulSizeSmall] = genForms(folderId, spreadsheet, filter),
    nSent = 0;

  if (ui.alert("Bracket", "Souhaitez-vous envoyer le lien du Google Form par mail aux participants ?", ui.ButtonSet.YES_NO) == ui.Button.YES) {
    // Loading screen
    displayLoadingScreen("Envoi du lien ..")
    nSent = mailLink(spreadsheet, publishedLink);
    // Adding a trigger to the dashboard to update it after each vote
    ScriptApp.newTrigger("updateNextRound")
      .forSpreadsheet(spreadsheet)
      .onFormSubmit()
      .create();

    // Confirmation message (also ends the loading screen)
    let htmlOutput = HtmlService.createHtmlOutput(`<span style="font-family: 'trebuchet ms', sans-serif;">Voci le lien éditeur : <a href = "${editLink}">éditer le Form</a>.<br/>
                                              <br/> Voici le lien lecteur : <a href = "${publishedLink}">répondre au Form</a>.<br/>
                                              <br/> Le lien lecteur a été envoyé à ${nSent} personne${nSent >= 2 ? "s" : ""}.</span>`);
    ui.showModelessDialog(htmlOutput, "Succès de l'opération");
  }
}
