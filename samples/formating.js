/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Uses functions displayLoadingScreen, properCase and addSpaces (taken from ressources.js)

// Returns a date
function formatDate(phrase) {
  return new Date(phrase);
}

// Takes a str, erases "+33", ".", "/", "-" and converts it into an integer
function formatAndConvert(str) {
  return Math.floor(str.toString().replace(/\s|\/|\+33|\.|-/g, ""));
}

// Formats selected range
function format_SelectRange() {
  // data stores the data contained in the selected range only
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
    data = sheet.getActiveRange().getValues(),
    start_column = sheet.getActiveRange().getColumn(),
    ui = SpreadsheetApp.getUi();

  // Confirm selection
  let confirm_selection = ui.alert("Formatage des données", "Avez-vous bien sélectionné les cellules à modifier ?", ui.ButtonSet.YES_NO);
  if (confirm_selection == ui.Button.NO) {
    return;
  }

  // User input
  let flag = true;
  while (flag) {
    var query = ui.prompt("Formatage des données",
      "Entrez le numéro correspondant au type de formatage nécessaire (1, 2, 3 ou 4) : \n 1. Nom/Prénom \n 2. Numéro de téléphone avec espaces \n 3. Numéro de téléphone sans espaces \n 4. Date",
      ui.ButtonSet.OK_CANCEL);
    if (query.getSelectedButton() == ui.Button.CANCEL) {
      return;
    }
    if (["1", "2", "3", "4"].indexOf(query.getResponseText()) != -1) {
      flag = false;
    } else {
      ui.alert("Formatage des données", "Entrée non valide, veuillez recommencer.", ui.ButtonSet.OK)
    }
  }

  // Loading screen
  displayLoadingScreen("Formatage des données..");

  // Formating the data
  if (query.getResponseText() == "1") {
    data.forEach(function (row, row_idx) {
      row.forEach(function (el, col_idx) {
        data[row_idx][col_idx] = (el == "" || !isNaN(formatAndConvert(el)) ? el : properCase(el))
      })
    });
  } else if (query.getResponseText() == "2") {
    data.forEach(function (row, row_idx) {
      row.forEach(function (el, col_idx) {
        data[row_idx][col_idx] = (el == "" || isNaN(formatAndConvert(el)) ? el : add_spaces("'0" + formatAndConvert(el)))
      })
    });
  } else if (query.getResponseText() == "3") {
    data.forEach(function (row, row_idx) {
      row.forEach(function (el, col_idx) {
        data[row_idx][col_idx] = (el == "" || isNaN(formatAndConvert(el)) ? el : "'0" + formatAndConvert(el))
      })
    });
  } else if (query.getResponseText() == "4") {
    data.forEach(function (row, row_idx) {
      row.forEach(function (el, col_idx) {
        data[row_idx][col_idx] = (el == "" || isNaN(formatAndConvert(el)) ? el : formatDate(el));
      })
    });
  }

  // Confirmation message
  let confirm_modif = ui.alert("Formatage des données", `Les colonnes ${start_column} à ${start_column + data[0].length - 1} seront formatées. \n Confirmez-vous cette sélection ?`, ui.ButtonSet.YES_NO);
  if (confirm_modif == ui.Button.NO) {
    return;
  }

  // Writing the formated values
  sheet.getActiveRange().setValues(data);
}