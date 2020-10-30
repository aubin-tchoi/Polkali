// Lets you synchronize a sheet with another sheet


function sync() {
  
  function getFilters(sheet) {
    let filterObj = {};
    for (let col=1;col<=sheet.getLastColumn();col++) {
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
    Object.keys(filters).forEach(function(col) {sheetdst.getFilter().removeColumnFilterCriteria(col);});
    
    sheetdst.getRange(1, 1, text.length, text[0].length)
    .clearContent()
    .setValues(text)
    .setBackgrounds(colors)
    .setFontColors(text.map(row => row.map(element => "black")));
    
    // Setting the filters back
    Object.keys(filters).forEach(function(col) {sheetdst.getFilter().setColumnFilterCriteria(col, filters[col]);});
  }
  
  const planscr = SpreadsheetApp.openById("").getSheetByName(""),
      plandst = SpreadsheetApp.openById("").getSheetByName("");
  
  copyData(planscr, plandst);
}
