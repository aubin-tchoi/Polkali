// Author : Pôle Qualité 022 (Aubin Tchoï)

// Formats selected range
function format_SelectRange() {
  // Capital letter at the beginning of each word
  function properCase(phrase) {
    return (phrase.toLowerCase().replace(/([^A-Za-zÀ-ÖØ-öø-ÿ])([A-Za-zÀ-ÖØ-öø-ÿ])(?=[A-Za-zÀ-ÖØ-öø-ÿ]{2})|^([A-Za-zÀ-ÖØ-öø-ÿ])/g, function(_, g1, g2, g3) {
      return (typeof g1 === 'undefined') ? g3.toUpperCase() : g1 + g2.toUpperCase(); }));
  }
  
  // Space every 2 char
  function addSpaces(phrase) {
    return x.toString().replace(/\B(?=(\d{2})+(?!\d))/g, " ");
  }
  
  // Returns a date
  function formatDate(phrase) {
    return new Date(phrase);
  }
  
  // data stores the data contained in the selected range only
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
      data = sheet.getActiveRange().getValues(),
      start_column = sheet.getActiveRange().getColumn(),
      ui = SpreadsheetApp.getUi(),
      htmlLoading = HtmlService
  .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/loadingScreen.gif" alt="Loading" width="885" height="498">`)
  .setWidth(900)
  .setHeight(500);
  
  // Confirm selection
  let confirm_selection = ui.alert("Formatage des données", "Avez-vous bien sélectionné les cellules à modifier ?", ui.ButtonSet.YES_NO);
  if (confirm_selection == ui.Button.NO) {return;}
  
  // User input
  let flag = true;
  while (flag) {
    var query = ui.prompt("Formatage des données", 
                          "Entrez le numéro correspondant au type de formatage nécessaire (1, 2, 3 ou 4) : \n 1. Nom/Prénom \n 2. Numéro de téléphone avec espaces \n 3. Numéro de téléphone sans espaces \n 4. Date",
                          ui.ButtonSet.OK_CANCEL);
    if (query.getSelectedButton() == ui.Button.CANCEL) {return;}
    if (["1", "2", "3", "4"].indexOf(query.getResponseText()) != -1) {flag = false;}
    else {ui.alert("Formatage des données", "Numéro non valide, veuillez recommencer.", ui.ButtonSet.OK)}
  }
  
  // Loading screen
  ui.showModelessDialog(htmlLoading, "Formatage des données..");
  
  // Formating the data
  if (query.getResponseText() == "1") {
    data.forEach(function(row, row_idx) {row.forEach(function(el, col_idx) {data[row_idx][col_idx] = (el == "" || !isNaN(Math.floor(el.toString().replace(/\s|\/|\+33|\.|-/g, "")))) ? el : properCase(el)})});
  }
  else if (query.getResponseText() == "2") {
    data.forEach(function(row, row_idx) {row.forEach(function(el, col_idx) {data[row_idx][col_idx] = (el == "" || isNaN(Math.floor(el.toString().replace(/\s|\/|\+33|\.|-/g, "")))) ? el : add_spaces("'0" + Math.floor(el.toString().replace(/\s|\/|\+33|\.|-/g, "")))})});
  }
  else if (query.getResponseText() == "3") {
    data.forEach(function(row, row_idx) {row.forEach(function(el, col_idx) {data[row_idx][col_idx] = (el == "" || isNaN(Math.floor(el.toString().replace(/\s|\/|\+33|\.|-/g, "")))) ? el : "'0" + Math.floor(el.toString().replace(/\s|\/|\+33|\.|-/g, ""))})});
  }
  else if (query.getResponseText() == "4") {
    data.forEach(function(row, row_idx) {row.forEach(function(el, col_idx) {data[row_idx][col_idx] = (el == "" || isNaN(Math.floor(el.toString().replace(/\s|\/|\+33|\.|-/g, "")))) ? el : formatDate(el);})});
  }
  
  // Confirmation message
  let confirm_modif = ui.alert("Formatage des données", `Les colonnes ${start_column} à ${start_column + data[0].length - 1} seront formatées. \n Confirmez-vous cette sélection ?`, ui.ButtonSet.YES_NO);
  if (confirm_modif == ui.Button.NO) {return;}
  
  // Writing the formated values
  sheet.getActiveRange().setValues(data);
}