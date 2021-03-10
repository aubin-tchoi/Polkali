// Checking if a str can be converted into a float
function isNumeric(str) {
    if (typeof str != "string") {
        return false;
    } // Only processes strings
    return !isNaN(str) && // use type coercion to parse the _entirety_ of the string (`parseFloat` alone does not do this)...
        !isNaN(parseFloat(str)) // ...and ensure strings of whitespace fail
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

// Converting an array into a multi-indexed-line str
function listToQuery(arr) {
    let query = "";
    arr.forEach(function (el, idx) {
        query += `\n ${idx + 1}. ${el}`
    });
    return query;
}

// Loading screen
function displayLoadingScreen(msg) {
    let htmlLoading = HtmlService
        .createHtmlOutput(`<img src="${IMAGES["loadingScreen"]}" alt="Loading" width="440" height="245">`)
        .setWidth(450)
        .setHeight(250);
    ui.showModelessDialog(htmlLoading, msg);
}

// Choosing how this question should be treated (PieChart, BarChart, ..)
function questionType(question, data) {
    // Excluding some questions
    if (uniqueValues(question, data).every(element => element == "")) {
        return "TextResponse";
    }

    // Less than 10 distinct values, and all values are integers
    if (uniqueValues(question, data).length <= 10 && uniqueValues(question, data).every(element => isNumeric(element))) {
        return "ColumnChart";
    }

    // Less than 6 distinct values (not integers)
    if (uniqueValues(question, data).length <= 6) {
        return "PieChart";
    }

    // Too many different responses, it has to be a qualitative question
    else {
        return "TextResponse";
    }
}