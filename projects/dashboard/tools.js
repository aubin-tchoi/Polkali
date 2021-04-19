/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Various tools

// Returning the last non-empty row of a sheet (not possible with getLastRow for this particular sheet bc of data validation)
function manuallyGetLastRow(sheet) {
    let data = sheet.getRange(1, 1, sheet.getLastRow(), 5).getValues();
    for (let row = 4, lastRow = sheet.getLastRow(); row < lastRow; row++) {
        if ((data[row][0] == "" && data[row][1] == "" && data[row][2] == "" && data[row][4] == "") || row > 200) {
            return row;
        }
    }
}

// Loading screen
function displayLoadingScreen(msg) {
    ui.showModelessDialog(HTML_CONTENT["loadingScreen"], msg);
}

// Getting an array of all unique values inside of a set of data for 1 information
const uniqueValues = (key, data) => data.map((row) => row[key]).filter((val, idx, arr) => arr.indexOf(val) == idx);

// Converting a chart object into an image
function convertChart(chart, title, htmlOutput, attachments) {
    // Adding the chart to the HtmlOutput
    let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
        imageUrl = "data:image/png;base64," + encodeURI(imageData);
    htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

    // Adding the chart to the attachments
    let imageDatamail = chart.getAs('image/png').getBytes(),
        imgblob = Utilities.newBlob(imageDatamail, "image/png", title);
    attachments.push(imgblob);
}

// Computes a / b * 100 if b !=0, and returns 0 otherwise
const prcnt = (a, b) => (parseInt(b, 10) == 0) ? 0 : a / b * 100;

// Checks if row was written in month month or not
const sameMonth = (row, month) => parseInt(row[HEADS["premierContact"]].getMonth(), 10) == month["month"] && parseInt(row[HEADS["premierContact"]].getFullYear(), 10) == month["year"];

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
            } catch(e) {
            ui.alert(message.title, message.incorrectInput, ui.ButtonSet.OK);
            }
        }
    } else {
        return ui.prompt(message.title, message.query, ui.ButtonSet.OK).getResponseText();
    }
}
