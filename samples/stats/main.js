/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Check https://github.com/aubin-tchoi/Polkali/tree/main/projects/suivi_prospection for a fully operational application

const DRIVE_FOLDERS = {
    stats: "",
  },
  MAIL_ADRESS = "",
  NAME_SHEET = {
    source: "",
    destination: ""
  },
  DIMS = {
    width: 750,
    height: 400
  },
  IMAGES = {
    loadingScreen: "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/loadingScreen.gif"
  }
ui = SpreadsheetApp.getUi();

// Retrieves data from designated sheets
function statsMerged(filtered) {
  // Initialization
  const ui = SpreadsheetApp.getUi(),
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = ss.getSheetByName(NAME_SHEET["source"]);
  let data = sheet.getRange(1, 2, sheet.getLastRow(), sheet.getLastColumn() - 1).getValues(),
    heads = data.shift();
  data = data.map(r => heads.reduce((o, k, i) => (o[k] = (r[i] != "") ? r[i] : o[k] || '', o), {}));

  let respQuali = [];

  // Filtering by class if requested
  if (filtered) {
    let key = ui.prompt("Filtrage des données", "Suivant quelle information souhaitez-vous filtrer les données ? (Entrez un nom de colonne)", ui.ButtonSet.OK).getResponseText();
    data = filterByKey(data, heads, key);
  }

  displayLoadingScreen("Chargement des diagrammes..");

  // Final outputs (displaying the charts on screen & mail content)  
  let htmlOutput = HtmlService
    .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">Voici les statistiques dressées sur l'ensemble des ${data.length} lignes de données :<br/></span> </span> <br/>`)
    .setWidth(800)
    .setHeight(465);

  let htmlMail = HtmlService.createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">&nbsp; &nbsp;
  Bonjour, <br/><br/> Voici les diagrammes récapitulatifs des ${data.length} lignes de données.<br/> <br/>
  Bonne journée !</span> </span>`),
    attachments = [];

  heads.forEach(function (question) {
    Logger.log(`Question : ${question}, type : ${questionType(question, data)}`);
    if (questionType(question, data) == "PieChart") {
      [htmlOutput, attachments] = createPieChart(question, data, htmlOutput, attachments);
    } else if (questionType(question, data) == "TextResponse") {
      respQuali.push(question);
    } else if (questionType(question, data) == "ColumnChart") {
      [htmlOutput, attachments] = createColumnChart(question, data, htmlOutput, attachments);
    }
  });

  sendMail(htmlMail, MAIL_ADRESS, "Statistiques groupées", attachments);
  saveOnDrive(DRIVE_FOLDERS["stats"], attachments);
  rewrite(data, respQuali, ss);

  // Final display of the charts
  ui.showModalDialog(htmlOutput, "Réponses aux questionnaires");
}

function statsMergedFetchAllSheets() {
  statsMerged(false);
}

function statsMergedFiltered() {
  statsMerged(true);
}
