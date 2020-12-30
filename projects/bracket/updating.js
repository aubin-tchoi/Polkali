/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

// onFormSubmit trigger function

// Adds the most recent vote to our dahsboard
function addLatestVote(sheet, row, groupNumber, groupSize, startColumn) {
  // Normalizes an array in each slice of size groupSize
  function normalize(array, groupSize) {
    let sum = 0,
      newArray = [];
    for (let group = 0; group < array.length / groupSize; group++) {
      sum = array.slice(group * groupSize, (group + 1) * groupSize).reduce((acc, val) => acc + val, 0);
      newArray = newArray.concat(array.slice(group * groupSize, (group + 1) * groupSize)
        .map(val => parseInt(val * GAME_PARAMETERS["pointsPool"] / sum, 10)));
    }
    return newArray;
  }

  const reference = sheet.getRange(sheet.getLastRow(), (startColumn + 1), 1, groupNumber * groupSize);

  Logger.log(`Unnormalized vote : ${row}`);
  row = normalize(row, groupSize);
  Logger.log(`Normalized vote : ${row}`);

  // Writing the new values on sheet
  sheet.getRange((sheet.getLastRow() + 1), (startColumn + 1), 1, (groupNumber * groupSize)).setValues([addLosers(row, reference)]);
}

// Adds a "Total Round n" line
function addSumFormula(sheet, row, groupNumber, groupSize, startColumn) {
  const reference = sheet.getRange(sheet.getLastRow(), (startColumn + 1), 1, groupNumber * groupSize),
    numberVotes = detectColor(sheet, MARKERS["nextRound"])[0] - detectColor(sheet, MARKERS["currentRound"]) - 1,
    formulas = Array.from(Array(row.length).keys()).map((_, idx) => `=SUM(R[-${numberVotes}]C[0]:R[-1]C[0])`);

  // Adding the formulas to the dashboard
  sheet.getRange((sheet.getLastRow() + 1), (startColumn + 1), 1, (groupNumber * groupSize)).setValues([addLosers(formulas, reference)]);

  // Writing on the left cells
  sheet.getRange((sheet.getLastRow() - 1), startColumn, 2, 1)
  .setValues([[""], [`Total R${findRoundNumber(sheet)}`]])
  .setBackgrounds([["white"], [MARKERS["nextRound"]]]);
}

// Sets a different background color for the 2 biggest values in each group
function highlightWinners(targetRange, groupNumber, groupSize) {
  let targetValues = targetRange.getDisplayValues()[0],
    backgrounds = [];

  // Slicing up the range by group
  for (let group = 0; group < groupNumber; group++) {
    let slice = targetValues.slice(group * groupSize, (group + 1) * groupSize),
      backgroundSlice = Array.from(Array(groupSize).keys()),
      biggestValue = 0,
      biggestIndex = groupSize;

    // Looking for the two biggest values
    for (let k = 0; k < GAME_PARAMETERS["winnersNumber"]; k++) {
      let newIdx = slice.indexOf(Math.max(...slice));
      biggestIndex = newIdx >= biggestIndex ? newIdx + 1 : newIdx;
      backgroundSlice[biggestIndex] = COLORS["winner"];
      biggestValue = slice.splice(biggestIndex, 1);
    }

    // Checking for tie
    if (Math.max(...slice) == biggestValue) {
      let newIdx = slice.indexOf(biggestValue);
      backgroundSlice[biggestIndex] = COLORS["tie"];
      backgroundSlice[newIdx >= biggestIndex ? newIdx + 1 : newIdx] = COLORS["tie"];
    }

    // Adding the slice to backgrounds
    backgrounds = backgrounds.concat(backgroundSlice);
  }
  targetRange.setBackgrounds([backgrounds]);
}

// Sets borders around each group
function borderGroups(sheet, position, groupNumber, groupSize) {
  for (let group = 0; group < groupNumber; group++) {
    sheet.getRange(position[0], (position[1] + group * groupSize), (sheetdst.getLastRow() - position[0]), groupSize)
      .setBorder(true, true, true, true, false, false);
  }
}

// Updates the dashboard after each new vote
function updateSheet() {
  const sheetsrc = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
    data = sheetsrc.getRange(sheetsrc.getLastRow(), 3, 1, (sheetsrc.getLastColumn() - 2)).getValues()[0],
    sheetdst = SpreadsheetApp.openById(IDS["dashboard"]).getSheetByName("Dashboard"),

    [startRow, startColumn] = detectColor(sheetdst, MARKERS["currentRound"]),
    [groupNumber, groupSize] = findGroups(sheetdst);

  addLatestVote(sheet, data, groupNumber, groupSize, startColumn);
  addSumFormula(sheet, data, groupNumber, groupSize, startColumn);
  highlightWinners(sheet, data, groupNumber, groupSize);
  borderGroups(sheet, [startRow + 1, startColumn + 1], groupNumber, groupSize);
}