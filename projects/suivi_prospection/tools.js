/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Various tools

// Returning the last non-empty row of a sheet (not possible with getLastRow for this particular sheet bc of data validation)
function manuallyGetLastRow(sheet) {
    let data = sheet.getRange(1, 1, sheet.getLastRow(), 5).getValues();
    for (let row = 4; row < sheet.getLastRow(); row++) {
        if ((data[row][0] == "" && data[row][1] == "" && data[row][2] == "" && data[row][4] == "") || row > 200) {
            return row;
        }
    }
}

// Loading screen
function displayLoadingScreen(msg) {
    let htmlLoading = HtmlService
        .createHtmlOutput(`<img src="${IMGS["loadingScreen"]}" alt="Loading" width="442" height="249">`)
        .setWidth(450)
        .setHeight(280);
    ui.showModelessDialog(htmlLoading, msg);
}

// Getting an array of all unique values inside of a set of data for 1 information (heads.indexOf(str) being the index of the column that contains the data inside of each row)
function uniqueValues(str, data) {
    let val = [];
    for (let row = 0; row < data.length; row++) {
        if (data[row][str] != "") {
            if (val.indexOf(data[row][str]) == -1) {
                val.push(data[row][str]);
            }
        }
    }
    return val;
}

// Computes a / b * 100 if b !=0, and returns 0 otherwise
const prcnt = (a, b) => (parseInt(b, 10) == 0) ? 0 : a / b * 100;