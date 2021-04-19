/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// ----- Data extraction tools -----

// I chose to keep the data as an array of objects (other possibility : object of objects, the keys being the months' names and the values the data in each row)
function extractSheetData(spreadsheetId, sheetName, pos) {
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName),
        data = sheet.getRange(pos.data.x, pos.data.y, (manuallyGetLastRow(sheet, pos.data.x, pos.trustColumn) - pos.data.x + 1), (sheet.getLastColumn() - pos.data.y + 1)).getValues(),
        heads = sheet.getRange(pos.header.x, pos.header.y, 1, (sheet.getLastColumn() - pos.header.y + 1)).getValues().shift();
    return data.map(row => heads.reduce((obj, key, idx) => (obj[key] = (row[idx] != "") ? row[idx] : obj[key] || '', obj), {}));
}

// Returning the last non-empty row of a sheet (not possible with getLastRow for this particular sheet bc of data validation)
function manuallyGetLastRow(sheet, start, trustColumn) {
    // 100 is considered as a higher bound for the number of rows inside of this sheet
    let data = sheet.getRange(start, trustColumn, (100 - start + 1), 1).getValues();
    for (row = 0; row < 100; row++) {
        if ((data[row][0]) == "") {
            return row + start;
        }
    }
}


// ----- Computation tools -----

// Getting an array of all unique values inside of a set of data for 1 information
const uniqueValues = (key, data) => data.map((row) => row[key]).filter((val, idx, arr) => arr.indexOf(val) == idx);

// Converting a chart object into an image
function convertChart(chart, title, htmlOutput, attachments) {
    // Loading the blob for this chart (this part takes the most time)
    let chartBlob = chart.getAs('image/png');
    // Adding the chart to the HtmlOutput
    htmlOutput.append(`<img border="1" src="data:image/png;base64,${encodeURI(Utilities.base64Encode(chartBlob.getBytes()))}">`);
    // Adding the chart to the attachments
    attachments.push(chartBlob.setName(title));
}

// Computes a / b * 100 if b !=0, and returns 0 otherwise
const prcnt = (a, b) => (parseInt(b, 10) == 0) ? 0 : a / b * 100;

// Checks if row was written in month month or not
const sameMonth = (row, month) => parseInt(row[HEADS.premierContact].getMonth(), 10) == month.month && parseInt(row[HEADS.premierContact].getFullYear(), 10) == month.year;


// ----- User's side features -----

// Loading screen
function displayLoadingScreen(msg) {
    ui.showModelessDialog(HTML_CONTENT.loadingScreen, msg);
}

// Retrieving a user query and checking if it is correct if message contains a incorrectInput attribute (assuming it matches with a folder's ID)
const userQuery = (message) => {
    let response = "";
    if (message.incorrectInput != undefined) {
        while (true) {
            try {
                response = ui.prompt(message.title, message.query, ui.ButtonSet.OK).getResponseText();
                // Checking whether the input is valid or not
                DriveApp.getFolderById(response);
                return response;
            } catch (e) {
                ui.alert(message.title, message.incorrectInput, ui.ButtonSet.OK);
            }
        }
    } else {
        return ui.prompt(message.title, message.query, ui.ButtonSet.OK).getResponseText();
    }
}

// Logs the execution time and resets the checkpoint
const measureTime = (checkpoint, message) => {
    let newCheckpoint = new Date();
    Logger.log(`It took ${(newCheckpoint.getTime() - checkpoint.getTime()) / 1000} s to ${message}.`);
    return newCheckpoint;
};

// There will be a time when this decorator will be put to a good use
const logTime = (message) => ((target, name, descriptor) => {
    let initialTime = new Date();
    descriptor.value();
    measureTime(initialTime, message);
});
