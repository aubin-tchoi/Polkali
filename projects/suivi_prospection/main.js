/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */


// Main file (triggers + main function for KPI and synchronization)

const ui = SpreadsheetApp.getUi(),
  COLORS = {
    burgundy: "#8E3232",
    gold: "#FFBE2B",
    grey: "#404040",
    lightGrey: "#A29386"
  },
  SHEETS = {
    id: "1lJhJuZxUt_8_mVLXe5tazXPrb2Z3wr0M49rho974sNQ",
    name: "Suivi"
  },
  HEADS = {
    état: "État",
    devis: "Devis réalisé",
    entreprise: "Entreprise"
  },
  DIMS = {
    width: 1000,
    height: 500
  },
  HTML_CONTENT = {
    display: `<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">KPI KPI KPI<br/></span> </span> <br/>`,
    mail: `<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">&nbsp; &nbsp; Bonjour, <br/><br/>Voici les KPI portant sur la prospection.<br/> <br/>Bonne journée !</span> </span>`
  },
  DRIVE = {
    folderId: "1dPi0dht-q_rI8fUmheA1j861huYPPcAy"
  }
monthNames = ["Jan", "Fév", "Mars", "Avr", "Mai", "Juin", "Juil", "Août", "Sept", "Oct", "Nov", "Déc"],
  stateList = ["Premier RDV réalisé", "Devis rédigé et envoyé", "En négociation", "Etude obtenue"];

/* ----- Triggers ----- */
function onOpen() {
  ui.createMenu('KPI')
    .addItem("Synchronisation de la sélection avec le suivi d'études", "syncOnSelec")
    .addItem("Affichage des KPI", "displayKPI")
    .addToUi();
}

/*
function onEdit(e) {
  let range = e.range;
  if (range.getNumColumns() == 1 && range.getNumRows() == 1 && e.value == stateList[3]) {
    syncOnEdit(range.getRow());
  }
} */

/* ------ KPI ----- */
function generateKPI(mailResults, saveResults, displayResults) {
  // Initialization
  const sheet = SpreadsheetApp.openById(SHEETS["id"]).getSheetByName(SHEETS["name"]),
    data = sheet.getRange(4, 2, (manuallyGetLastRow(sheet) - 3), 12).getValues(),
    heads = sheet.getRange(1, 2, 1, 12).getValues().shift();

  // I chose to keep the data as an array of objects (other possibility : object of objects, the keys being the months' names and the values the data in each row)
  let obj = data.map(r => heads.reduce((o, k, i) => (o[k] = r[i] || 0, o), {})).filter(row => row["Premier contact"] != "");

  // console.log
  Logger.log(obj);
  Logger.log(heads);

  displayLoadingScreen("Chargement des KPI..");

  // Final outputs (displaying the charts on screen & mail content)  
  let htmlOutput = HtmlService
    .createHtmlOutput(HTML_CONTENT["display"])
    .setWidth(1015)
    .setHeight(515),
    htmlMail = HtmlService.createHtmlOutput(HTML_CONTENT["mail"]),
    attachments = [],
    conversionChart = [];

  // /!\ LE PROCHAIN MANDAT REMMETTEZ MAI JUIN JUILLET (ou mettez mois + année pour couvrir plusieurs années)
  let monthList = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 0];


  /* ------------ KPI (Contacts par mois) ------------ */
  let dataTable1 = Charts.newDataTable();
  // Columns
  dataTable1.addColumn(Charts.ColumnType.STRING, "Mois");
  stateList.forEach(function (state) {
    dataTable1.addColumn(Charts.ColumnType.NUMBER, state);
  });
  // Rows
  monthList.forEach(function (month) {
    let dataRow = [monthNames[month]];
    let conversionChartRow = [];
    stateList.forEach(function (state) {
      let val = obj.filter(row => parseInt(row["Premier contact"].getMonth(), 10) == month && stateList.indexOf(row[HEADS["état"]]) >= stateList.indexOf(state)).length;
      if (state == "Devis rédigé et envoyé" || state == "En négociation") {
        val += obj.filter(row => parseInt(row["Premier contact"].getMonth(), 10) == month && (row[HEADS["état"]] == "Sans suite" || row[HEADS["état"]] == "A relancer") && row[HEADS["devis"]]).length
      } else if (state == "Premier RDV réalisé") {
        val += obj.filter(row => parseInt(row["Premier contact"].getMonth(), 10) == month && (row[HEADS["état"]] == "Sans suite" || row[HEADS["état"]] == "A relancer")).length
      }
      conversionChartRow.push(val);
      dataRow.push(val);
    });
    conversionChart.push(conversionChartRow);
    dataTable1.addRow(dataRow);
  });
  Logger.log(conversionChart);
  // Appending the corresponding ColumnChart
  [htmlOutput, attachments] = createColumnChart(dataTable1, "Contacts", htmlOutput, attachments, DIMS["width"], DIMS["height"]);

  /* ------------ KPI (Taux de conversion) ------------ */
  let dataTable2 = Charts.newDataTable();
  // Columns
  dataTable2.addColumn(Charts.ColumnType.STRING, "Mois");
  dataTable2.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${stateList[0]} et ${stateList[1]}`);
  dataTable2.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${stateList[1]} et ${stateList[2]}`);
  dataTable2.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${stateList[2]} et ${stateList[3]}`);
  // Rows
  const prcnt = (a, b) => (parseInt(b, 10) == 0) ? 0 : a / b * 100;
  for (let idx in monthList) {
    Logger.log(`Taux de conversion : ${[prcnt(conversionChart[idx][1], conversionChart[idx][0]), prcnt(conversionChart[idx][2], conversionChart[idx][1]), prcnt(conversionChart[idx][3], conversionChart[idx][2])]}`);
  }
  monthList.forEach(function (month, idx) {
    dataTable2.addRow([monthNames[month], prcnt(conversionChart[idx][1], conversionChart[idx][0]), prcnt(conversionChart[idx][2], conversionChart[idx][1]), prcnt(conversionChart[idx][3], conversionChart[idx][2])]);
  });
  // Creating the corresponding LineChart
  [htmlOutput, attachments] = createLineChart(dataTable2, ["#8E3232", "#FFBE2B", "#404040"], "Taux de conversion", htmlOutput, attachments, DIMS["width"], DIMS["height"]);

  /* ------------ KPI (CA) ------------ */
  let dataTable3 = Charts.newDataTable();
  // Columns
  dataTable3.addColumn(Charts.ColumnType.STRING, "Mois");
  dataTable3.addColumn(Charts.ColumnType.NUMBER, "CA sur étude potentielle");
  dataTable3.addColumn(Charts.ColumnType.NUMBER, "CA sur étude signée");
  // Rows
  monthList.forEach(function (month) {
    let ca_pot = 0,
      ca_sig = 0;
    obj.filter(row => parseInt(row["Premier contact"].getMonth(), 10) == month && row["État"] != "Etude obtenue").forEach(function (row) {
      if (Object.keys(row).includes("Prix potentiel de l'étude  € (HT)")) {
        ca_pot += parseInt(row["Prix potentiel de l'étude  € (HT)"], 10) * ((parseInt(row["Pourcentage de confiance à la conversion en réelle étude"], 10) > 0) ? parseInt(row["Pourcentage de confiance à la conversion en réelle étude"], 10) / 100 : 1);
      }
    });
    obj.filter(row => parseInt(row["Premier contact"].getMonth(), 10) == month && row["État"] == "Etude obtenue").forEach(function (row) {
      if (Object.keys(row).includes("Prix potentiel de l'étude  € (HT)")) {
        ca_sig += parseInt(row["Prix potentiel de l'étude  € (HT)"], 10);
      }
    });
    dataTable3.addRow([monthNames[month], ca_pot, ca_sig]);
  });
  // Creating the corresponding ColumnCharts
  [htmlOutput, attachments] = createColumnChart(dataTable3, "Chiffre d'affaires", htmlOutput, attachments, DIMS["width"], DIMS["height"]);

  /* ------------ KPI (Type de contact) ------------ */
  [htmlOutput, attachments] = createPieChartUniqueval("Type de contact", obj, htmlOutput, attachments, DIMS["width"], DIMS["height"]);

  /* ------------ KPI (Taux de conversion par type de contact) ------------ */
  let dataTable4 = Charts.newDataTable();
  // Columns
  dataTable4.addColumn(Charts.ColumnType.STRING, "Type de contact");
  dataTable4.addColumn(Charts.ColumnType.NUMBER, "Nombre de contacts");
  dataTable4.addColumn(Charts.ColumnType.NUMBER, "Taux de conversion");
  // Rows
  uniqueValues("Type de contact", obj).forEach(function (value) {
    let objFiltered = obj.filter(row => row["Type de contact"] == value);
    let conversionRate = prcnt(objFiltered.filter(row => row["État"] == "Etude obtenue").length, objFiltered.length);
    dataTable4.addRow([value, objFiltered.length, conversionRate]);
  });
  // Creating the corresponding ColumnCharts
  [htmlOutput, attachments] = createColumnChart(dataTable4, "Taux de conversion par type de contact", htmlOutput, attachments, DIMS["width"], DIMS["height"]);


  // Sending graphs by mail (mail adress will be prompted)
  if (mailResults) {
    sendMail(htmlMail, "KPI", attachments);
  }
  // Saving graphs in a Drive folder
  if (saveResults) {
    let folderId = ui.prompt("Enregistrement des images sur le Drive", "Entrez l'id du dossier de destination :", ui.ButtonSet.OK).getResponseText();
    saveOnDrive(attachments, folderId);
  }
  // Displaying the charts on screen
  if (displayResults) {
    ui.showModalDialog(htmlOutput, "KPI");
  }
}

function displayKPI() {
  let mailResults = ui.alert("Envoi des diagrammes par mail", "Souhaitez vous recevoir les diagrammes par mail ?", ui.ButtonSet.YES_NO) == ui.Button.YES,
    saveResults = ui.alert("Enregistrement des images sur le Drive", "Souhaitez vous enregistrer les images sur le Drive ?", ui.ButtonSet.YES_NO) == ui.Button.YES;
  generateKPI(mailResults, saveResults, true);
}

function autoSaveKPI() {
  generateKPI(false, true, false);
}