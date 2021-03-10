/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Spreadsheet IDs and sheet names are hard-coded
const SOURCE = {
  ssId: "",
  sheetName: ""
},
DESTINATION = {
  ssId: "",
  sheetName: ""
};

// Synchronizes SOURCE with DESTINATION
function sync() {
  function getFilters(sheet) {
    let filterObj = {};
    for (let col = 1; col <= sheet.getLastColumn(); col++) {
      if (sheet.getFilter().getColumnFilterCriteria(col) != null) {
        filterObj[col] = sheet.getFilter().getColumnFilterCriteria(col);
      }
    }
    return filterObj;
  }

  function copyData(sheetscr, sheetdst) {
    let text = sheetscr.getRange(1, 1, (sheetscr.getLastRow() + 30), sheetscr.getLastColumn()).getValues(),
      colors = sheetscr.getRange(1, 1, (sheetscr.getLastRow() + 30), sheetscr.getLastColumn()).getBackgrounds(),
      filters = getFilters(sheetdst);

    // Removing the filters on each column
    Object.keys(filters).forEach(function (col) {
      sheetdst.getFilter().removeColumnFilterCriteria(col);
    });

    sheetdst.getRange(1, 1, text.length, text[0].length)
      .clearContent()
      .setValues(text)
      .setBackgrounds(colors)
      .setFontColors(text.map(row => row.map(_ => "black")));

    // Setting the filters back
    Object.keys(filters).forEach(function (col) {
      sheetdst.getFilter().setColumnFilterCriteria(col, filters[col]);
    });
  }

  const sheetscr = SpreadsheetApp.openById(SOURCE["ssId"]).getSheetByName(SOURCE["sheetName"]),
    sheetdst = SpreadsheetApp.openById(DESTINATION["ssId"]).getSheetByName(DESTINATION["sheetName"]);

  copyData(sheetscr, sheetdst);
}
