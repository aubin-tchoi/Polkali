/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

// onFormSubmit trigger function

// Adds the most recent vote to our dahsboard
function addLatestVote(sheet, row, groupNumber, groupSize) {
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

  const reference = dashboard.getRange(dashboard.getLastRow(), (startColumn + 1), 1, groupNumber * groupSize);

  Logger.log(`Unnormalized vote : ${row}`);
  row = normalize(row, groupSize);
  Logger.log(`Normalized vote : ${row}`);

  // Writing the new values on sheet
  sheet.getRange((sheet.getLastRow() + 1), (startColumn + 1), 1, (groupNumber * groupSize)).setValues([addLosers(row, reference)]);
}

function addSumFormula(sheet, position) {
  
}

// Sets a different background color for the 2 biggest values in each group + borders around each group
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

function updateSheet() {
  const sheetsrc = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
    data = sheetsrc.getRange(sheetsrc.getLastRow(), 3, 1, (sheetsrc.getLastColumn() - 2)).getValues()[0],
    sheetdst = SpreadsheetApp.openById(IDS["dashboard"]).getSheetByName("Dashbord"),

    [startRow, startColumn] = detectColor(sheetdst, MARKERS["currentRound"]),
    [groupNumber, groupSize] = findGroups(sheetdst);

  addLatestVote(sheet, data, groupNumber, groupSize);
  addSumFormula(sheet, [startRow + 1, startColumn + 1]);
  highlightWinners(sheet, data, groupNumber, groupSize);
  borderGroups(sheet, [startRow + 1, startColumn + 1], groupNumber, groupSize);
}