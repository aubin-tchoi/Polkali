/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

// Loading screen
function displayLoadingScreen(msg) {
  let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="442" height="249">`)
    .setWidth(450)
    .setHeight(325);
  SpreadsheetApp.getUi().showModelessDialog(htmlLoading, msg);
}

// Detects the position of an element or a color in a sheet
function detectColor(sheet, color) {
  let data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getBackgrounds(),
    row = (data.indexOf(data.filter(r => r.includes(color))[0]) + 1),
    column = (data[row - 1].indexOf(color) + 1);
  return [row, column];
}

// Finds the number of the next round
function findRoundNumber(sheet) {
  const [row, col] = detectColor(sheet, MARKERS["round"]);
  return (parseInt(sheet.getRange(row, col).getValue().replace(/[^0-9]+/gi, ""), 10) + 1);
}

// Finds groupNumber and groupSize (only works with a formated sheet)
function findGroups(sheet) {
  const [startRow, startColumn] = detectColor(sheet, MARKERS["groups"]),
  groupNumber = Math.max(...sheet.getRange(startRow, (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0].map(el => parseInt(el.replace(/[^0-9]+/gi, ""), 10) || 0)),
  groupSize = Math.max(...sheet.getRange((startRow + 1), (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0]);
  return [groupNumber, groupSize];
}

// Creating a new array with blanks for eliminated contestants
function addLosers(row, reference) {
  let filterLosers = (idx => reference.getValues()[0][idx] == "" || [COLORS["loser"], COLORS["contestants"]].includes(reference.getBackgrounds()[0][idx])),
    newRow = [];
  for (let idx = 0; idx < reference.length; idx++) {
    newRow.push(filterLosers(idx) ? "" : row.shift());
  }
  return newRow;
}
