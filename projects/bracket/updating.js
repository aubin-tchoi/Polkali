/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

// onFormSubmit trigger function

// Sets a different background color for the 2 biggest values in each group + borders around each group
function highlightWinners(targetRange, poulNum, poulSize) {
  let targetValues = targetRange.getDisplayValues().shift(),
    big1 = Array.from(Array(poulNum).keys()).map(poul => Math.max(...(targetValues.slice(poul * poulSize, (poul + 1) * poulSize)))),
    big2 = Array.from(Array(poulNum).keys()).map(poul => Math.max(...(targetValues.slice(poul * poulSize, (poul + 1) * poulSize).filter(value => value != big1[poul])))),
    duplicate1 = Array.from(Array(poulNum).keys()).map(poul => (targetValues.slice(poul * poulSize, (poul + 1) * poulSize).filter(value => value == big1[poul]).length == 1 ? colors["winner"] : colors["duplicate"])),
    duplicate2 = Array.from(Array(poulNum).keys()).map(poul => (targetValues.slice(poul * poulSize, (poul + 1) * poulSize).filter(value => value == big2[poul]).length == 1 ? colors["winner"] : colors["duplicate"])),
    backgrounds = targetValues.map((value, index) => (value == big1[Math.floor(index / poulSize)] ? duplicate1[Math.floor(index / poulSize)] : value == big2[Math.floor(index / poulSize)] ? duplicate2[Math.floor(index / poulSize)] : colors["loser"]));

  Logger.log(`Biggest values for each group : ${big1}`);
  Logger.log(`Second biggest values for each group : ${big2}`);
  Logger.log(`Backgound colors : ${backgrounds}`);

  targetRange.setBackgrounds([backgrounds]);

  // Setting borders for each group
  for (let poul = 0; poul < poulNum; poul++) {
    sheet.getRange(startRow, (startColumn + 1 + poul * poulSize), (numberVotes + 3), (poulSize)).setBorder(true, true, true, true, false, false);
  }
}

// updates the sheet based on the first form
function updateFirstRound() {
  // Retrieving useful data
  const dataSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Round 1"),
    numberVotes = (dataSheet.getLastRow() - 1),
    sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Dashboard"),
    [startRow, startColumn] = detectPos(sheet, "Poule"),
    poulNum = Math.max(...sheet.getRange(startRow, (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0].map(el => parseInt(el.replace(/[^0-9]+/gi, ""), 10) || 0)),
    poulSize = Math.max(...sheet.getRange((startRow + 1), (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0]),
    targetRange = sheet.getRange((startRow + 1 + numberVotes + 1), (startColumn + 1), 1, poulNum * poulSize);

  // Retrieving the most recent vote and normalizing it (the total might be lower than 10 since we're only taking integers, if someones messed up his vote it might count less !)
  let unnormalizedVote = dataSheet.getRange(dataSheet.getLastRow(), 4, 1, (dataSheet.getLastColumn() - 3)).getValues().shift(),
    normalization = Array.from(Array(poulNum).keys()).map(poul => unnormalizedVote.slice(poul * poulSize, (poul + 1) * poulSize).reduce((a, b) => a + b, 0)),
    normalizedVote = unnormalizedVote.map((value, index) => parseInt(value * 10 / normalization[Math.floor(index / poulSize)], 10));

  Logger.log(`Number of groups : ${poulNum}`);
  Logger.log(`Size of a group : ${poulSize}`);
  Logger.log(`Unnormalized most recent vote : ${unnormalizedVote}`);
  Logger.log(`Normalized most recent vote : ${normalizedVote}`);

  // Adding the last vote to the dashboard
  sheet.getRange((startRow + 1 + numberVotes), (startColumn + 1), 1, poulNum * poulSize).setValues([normalizedVote]).setBackground("white").setHorizontalAlignment("center");

  // The last row contains the sum of all vote for given candidate
  let formulas = Array.from(Array(poulNum * poulSize).keys()).map(() => `=SUM(R[-${numberVotes}]C[0]:R[-1]C[0])`);
  Logger.log(`Formulas : ${formulas}`);
  targetRange.setFormulasR1C1([formulas]).setBackground(colors["loser"]).setHorizontalAlignment("center");

  // Identifying the 2 leading contestants
  highlightWinners(targetRange, poulNum, poulSize);

  // Moving the "Total" cell one row beneath where it was
  sheet.getRange((startRow + 1 + numberVotes), startColumn, 2, 1).setValues([[""],["Total"]]).setBackgrounds([["white"],[colors["groups"]]]);
}

function updateNextRound() {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId),
    dataSheet = spreadsheet.getSheetByName(`Round ${Math.max(...(spreadsheet.getSheets().map(sheet => sheet.getName()).map(el => parseInt(el.replace(/[^0-9]+/gi, ""), 10) || 0)))}`),
    numberVotes = (dataSheet.getLastRow() - 1),
    sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Dashboard"),
    [startRow, startColumn] = detectPos(sheet, "Poule"),
    poulNum = Math.max(...sheet.getRange(startRow, (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0].map(el => parseInt(el.replace(/[^0-9]+/gi, ""), 10) || 0)),
    poulSize = Math.max(...sheet.getRange((startRow + 1), (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0]),
    targetRange = sheet.getRange((startRow + 1 + numberVotes + 1), (startColumn + 1), 1, poulNum * poulSize);
}