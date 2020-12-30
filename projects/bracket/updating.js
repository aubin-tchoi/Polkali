/* For any further question, please contact Aubin Tchoï */

// onFormSubmit trigger function

// Adds the most recent vote to our dahsboard
function addLatestVote(sheet, row, voteNumber, groupNumber, groupSize, startRow, startColumn) {
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
  
  const reference = sheet.getRange((startRow + voteNumber), (startColumn + 1), 1, groupNumber * groupSize);

  Logger.log(`Unnormalized vote : ${row}`);
  row = normalize(row, groupSize);
  Logger.log(`Normalized vote : ${row}`);

  // Writing the new values on sheet
  sheet.getRange((startRow + voteNumber + 1), (startColumn + 1), 1, (groupNumber * groupSize))
    .setBackground("white")
    .setValues([addLosers(row, reference)]);
}

// Adds a "Total Round n" line
function addSumFormula(sheet, row, voteNumber, groupNumber, groupSize, startRow, startColumn) {
  const reference = sheet.getRange((startRow + voteNumber), (startColumn + 1), 1, groupNumber * groupSize),
    formulas = Array.from(Array(row.length).keys()).map((_, idx) => `=SUM(R[-${voteNumber + 1}]C[0]:R[-1]C[0])`);

  // Adding the formulas to the dashboard
  sheet.getRange((startRow + voteNumber + 2), (startColumn + 1), 1, (groupNumber * groupSize)).setValues([addLosers(formulas, reference)]);

  // Writing on the left cells
  if (voteNumber > 0) {
    sheet.getRange((startRow + voteNumber + 1), startColumn, 2, 1)
      .setValues([[""], [`Total R${findRoundNumber(sheet)}`]])
      .setBackgrounds([["white"], [MARKERS["nextRound"]]]);
  }
  else {
    sheet.getRange((startRow + 2), startColumn, 1, 1)
      .setValues([[`Total R${findRoundNumber(sheet)}`]])
      .setBackgrounds([[MARKERS["nextRound"]]]);
  }
}

// Sets a different background color for the 2 biggest values in each group
function highlightWinners(targetRange, groupNumber, groupSize) {
  let targetValues = targetRange.getDisplayValues()[0],
    backgrounds = [];

  // Slicing up the range by group
  for (let group = 0; group < groupNumber; group++) {
    let slice = targetValues.slice(group * groupSize, (group + 1) * groupSize),
      backgroundSlice = Array.from(Array(groupSize).keys()).map(() => "white"),
      biggestValue = 0,
      absoluteIndex = groupSize,
      offset = 1;

    // Looking for the two biggest values (only works if there is no more than 2 winners bc of the "+ 1" in line 67)
    for (let k = 0; k < GAME_PARAMETERS["winnersNumber"]; k++) {
      let relativeIndex = slice.indexOf(Math.max(...slice).toString());
      absoluteIndex = relativeIndex + offset > absoluteIndex ? relativeIndex + offset++ : relativeIndex;
      backgroundSlice[absoluteIndex] = COLORS["winner"];
      biggestValue = slice.splice(relativeIndex, 1);
    }

    // Checking for tie
    for (let k = 0; k < groupSize - GAME_PARAMETERS["winnersNumber"]; k++) {
      if (Math.max(...slice) == biggestValue) {
        let relativeIndex = slice.indexOf(biggestValue.toString());
        backgroundSlice[absoluteIndex] = COLORS["tie"];
        absoluteIndex = relativeIndex + offset >= absoluteIndex ? relativeIndex + offset++ : relativeIndex;
        backgroundSlice[absoluteIndex] = COLORS["tie"];
        biggestValue = slice.splice(relativeIndex, 1);
      }
    }

    // Adding the slice to backgrounds
    backgrounds = backgrounds.concat(backgroundSlice);
  }
  targetRange.setBackgrounds([backgrounds]);
}

// Sets borders around each group and centers numbers in their cells
function borderGroups(sheet, voteNumber, groupNumber, groupSize, startRow, startColumn) {
  for (let group = 0; group < groupNumber; group++) {
    sheet.getRange(startRow, (startColumn + group * groupSize), (voteNumber + 2), groupSize)
      .setBorder(true, true, true, true, false, false)
      .setHorizontalAlignment("center");
  }
}

// Updates the dashboard after each new vote
function updateSheet() {
  const sheetsrc = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
    data = sheetsrc.getRange(sheetsrc.getLastRow(), 4, 1, (sheetsrc.getLastColumn() - 3)).getValues()[0],
    voteNumber = (sheetsrc.getLastRow() - 2),
    sheetdst = SpreadsheetApp.openById(IDS["dashboard"]).getSheetByName("Dashboard"),

    [startRow, startColumn] = detectColor(sheetdst, MARKERS["currentRound"]),
    [groupNumber, groupSize] = findGroups(sheetdst);

  addLatestVote(sheetdst, data, voteNumber, groupNumber, groupSize, startRow, startColumn);
  addSumFormula(sheetdst, data, voteNumber, groupNumber, groupSize, startRow, startColumn);
  highlightWinners(sheetdst.getRange((startRow + voteNumber + 2), (startColumn + 1), 1, groupNumber * groupSize), groupNumber, groupSize);
  borderGroups(sheetdst, voteNumber, groupNumber, groupSize, (startRow + 1), (startColumn + 1));
}
