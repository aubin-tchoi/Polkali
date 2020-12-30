/* For any further question, please contact Aubin Tchoï */

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

// Finds the size of a column
function findColumnDepth(sheet, begin, column) {
  const values = sheet.getRange(begin, column, sheet.getLastRow(), 1).getValues(); 
  for (let depth = 0; depth < sheet.getLastRow() - begin + 2; depth++) {
    if (values[depth][0] == "") {
      return depth;
    }
  }
  return (sheet.getLastRow() - begin + 2);
}

// Finds the number of the next round
function findRoundNumber(sheet) {
  let [row, col] = [0, 0];
  try {
    [row, col] = detectColor(sheet, MARKERS["nextRound"]);
  }
  catch(e) {
    [row, col] = detectColor(sheet, MARKERS["currentRound"]);
  }
  return ((parseInt(sheet.getRange(row, col).getValue().replace(/[^0-9]+/gi, ""), 10) + 1) || 1);
}

// Finds groupNumber and groupSize (only works with a formated sheet)
function findGroups(sheet) {
  const [startRow, startColumn] = detectColor(sheet, MARKERS["groups"]),
    groupNumber = Math.max(...(sheet.getRange(startRow, (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0].map(el => parseInt(el.toString().replace(/[^0-9]+/gi, ""), 10) || 0))),
    groupSize = Math.max(...sheet.getRange((startRow + 1), (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0]);
  return [groupNumber, groupSize];
}

// Creating a new array with blanks for eliminated contestants
function addLosers(row, reference) {
  let filterLosers = (idx => reference.getValues()[0][idx].toString() == "" || reference.getBackgrounds()[0][idx] == COLORS["loser"]),
    newRow = [];
  for (let idx = 0; idx < reference.getValues()[0].length; idx++) {
    newRow.push(filterLosers(idx) ? "" : row.shift());
  }
  return newRow;
}
