/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Two parts in this script : synchronizing with "Suivi d'études" and producing some KPIs

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('KPI')
    .addItem("Synchronisation de la sélection avec le suivi d'études", "syncOnSelec")
    .addItem("Affichage des KPI", "statsKPI")
    .addToUi();
}

// Returning the last non-empty row of a sheet (not possible with getLastRow for this particular sheet bc of data validation)
function manuallyGetLastRow(sheet) {
  let data = sheet.getRange(1, 1, sheet.getLastRow(), 5).getValues();
  for (let row = 4; row < sheet.getLastRow(); row++) {
    if ((data[row][0] == "" && data[row][1] == "" && data[row][2] == "" && data[row][4] == "") || row > 200) {
      return row;
    }
  }
}


function onEdit(e) {
  let range = e.range;
  if (range.getNumColumns() == 1 && range.getNumRows() == 1 && e.value == "Etude obtenue") {
    syncOnEdit(range.getRow());
  }
}

function syncOnEdit(rowidx) {
  // Updating sheet sheet with row row
  function update(sheet, row) {
    // Checking whether PEP actually did obtain the mission or not
    if (row["État"] != "Etude obtenue ") {
      return;
    }

    // ref is gonna be the mission's reference, obtained by incrementing the most recent one
    const heads = sheet.getRange(1, 2, 1, (sheet.getLastColumn() - 1)).getValues().shift(),
      lastRow = manuallyGetLastRow(sheet),
      ref = `'20e${Math.floor(sheet.getRange(lastRow, 2, 1, 1).getValues()[0][0].toString().match(/([0-9]+)/gi)[1]) + 1}`;

    // Creating a new row with the destination sheet's header's informations
    let rnew = [row["Entreprise"]];
    heads.forEach(function (el) {
      rnew.push(el == "Référence de l'étude" ? ref : (el == "État") ? "Devis accepté" : row[el]);
    });
    sheet.getRange((lastRow + 1), 1, 1, rnew.length).setValues([rnew]);
  }

  // -- MAIN --
  const sheetscr = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
    sheetdst = SpreadsheetApp.openById("1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04").getSheetByName("Suivi");

  // heads is the sheet's header, data a js object of the edited row
  let heads = sheetscr.getRange(1, 1, 1, sheetscr.getLastColumn()).getValues().shift(),
    data = sheetscr.Range(rowidx, 1, 1, sheetscr.getLastColumn()).getValues().map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {}))).shift();

  // Synchronizing edited row
  update(sheetdst, data);
}

/* ------ updating ------ */
function syncOnSelec() {
  // Loading screen
  function displayLoadingScreen(msg) {
    let htmlLoading = HtmlService
      .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="442" height="249">`)
      .setWidth(450)
      .setHeight(280);
    ui.showModelessDialog(htmlLoading, msg);
  }

  // Updating sheet sheet with row row
  function update(sheet, row) {
    // Checking whether PEP actually did obtain the mission or not
    if (row["État"] != "Etude obtenue") {
      ui.alert("Entrée invalide", `L'étude confiée par l'entreprise ${row["Entreprise"]} ne correspond pas à une étude obtenue.`, ui.ButtonSet.OK);
      return;
    }

    // ref is gonna be the mission's reference, obtained by incrementing the most recent one
    const heads = sheet.getRange(1, 2, 1, (sheet.getLastColumn() - 1)).getValues().shift(),
      lastRow = manuallyGetLastRow(sheet),
      ref = `'20e${Math.floor(sheet.getRange(lastRow, 2, 1, 1).getValues()[0][0].toString().match(/([0-9]+)/gi)[1]) + 1}`;

    // Creating a new row with the destination sheet's header's informations
    let rnew = [row["Entreprise"]];
    heads.forEach(function (el) {
      rnew.push(el == "Référence de l'étude" ? ref : (el == "État") ? "Devis accepté" : row[el]);
    });
    sheet.getRange((lastRow + 1), 1, 1, rnew.length).setValues([rnew]);
  }

  // -- MAIN --
  const sheetscr = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
    sheetdst = SpreadsheetApp.openById("1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04").getSheetByName("Suivi"),
    ui = SpreadsheetApp.getUi();

  // Confirm selection
  let confirmSelection = ui.alert("Synchronisation des données", "Vous devez préalablement sélectionner la ligne à synchroniser (par exemple en cliquant sur le numéro à gauche). \n Confirmez-vous votre sélection ?", ui.ButtonSet.YES_NO);
  if (confirmSelection == ui.Button.NO) {
    return;
  }

  // Loading screen
  displayLoadingScreen("Synchronisation ..");

  // heads is the sheet's header, data a 2D-array representation of the selected values, and obj its 'array of js object' representation
  let heads = sheetscr.getRange(1, 1, 1, sheetscr.getLastColumn()).getValues().shift(),
    data = sheetscr.getActiveRange().getValues(),
    obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Checking the selection again (empty or not)
  if (obj.length == 0) {
    ui.alert("Pas de ligne sélectionnée, veuillez recommencer.");
    return;
  }

  // Synchronizing each selected row
  obj.forEach(function (row) {
    update(sheetdst, row);
  });

  // Confirmation
  let imgUrl = "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Ddabong.png",
    sheetUrl = "https://docs.google.com/spreadsheets/d/1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04/edit#gid=0",
    operationSuccess = HtmlService
    .createHtmlOutput(`<span style='font-size: 16pt;'> <span style="font-family: 'roboto', sans-serif;">L'opération a été effectuée avec succès, veuillez remplir manuellement le nom de l'étude dans le <a href="${sheetUrl}">suivi des études</a>.<br/><br/> La bise</span></span><p style="text-align:center;"><img src="${imgUrl}" alt="C'est la PEP qui régale !" width="130" height="131"></p>`);
  ui.showModalDialog(operationSuccess, "Synchronisation des données");
}


/* ------ KPI KPI KPI KPI KPI KPI KPI KPI ----- */

function statsKPI() {
  // Loading screen
  function displayLoadingScreen(msg) {
    let htmlLoading = HtmlService
      .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="442" height="249">`)
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

  // Query and send mail to designated adress
  function sendMail(htmlOutput, subject, attachments) {
    let query = ui.alert("Envoi des diagrammes par mail", "Souhaitez vous recevoir les diagrammes par mail ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
      let adress = ui.prompt("Envoi des diagrammes par mail", "Entrez l'adresse mail de destination :", ui.ButtonSet.OK).getResponseText(),
        msgHtml = htmlOutput.getContent(),
        msgPlain = htmlOutput.getContent().replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/ig, "");
      GmailApp.sendEmail(adress, subject, msgPlain, {
        htmlBody: msgHtml,
        attachments: attachments
      });
      ui.alert("Envoi des diagrammes par mail", `Les diagrammes ont été envoyés par mail à : ${adress}.`, ui.ButtonSet.OK);
    }
  }
  // Query and save data in designated Drive folder
  function saveOnDrive(imageBlobs) {
    let query = ui.alert("Enregistrement des images sur le Drive", "Souhaitez vous enregistrer les images sur le Drive ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
      try {
        let folderId = ui.prompt("Enregistrement des images sur le Drive", "Entrez l'id du dossier de destination :", ui.ButtonSet.OK).getResponseText();
        displayLoadingScreen("Enregistrement des images sur le Drive..");
        // Folder will be dated with current date
        let today = new Date();
        today = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
        let folder = DriveApp.getFolderById(folderId).createFolder(`Statistiques groupées ${today}`);
        imageBlobs.forEach(function (f) {
          folder.createFile(f);
        });
        ui.alert("Enregistrement des images sur le Drive", `Les images ont été enregistrées à ladresse suivante : ${folder.getUrl()}`, ui.ButtonSet.OK);
      } catch (e) {
        Logger.log(`Erreur lors de l'enregistrement des images sur le Drive : ${e}.`);
        ui.alert("Erreur lors de l'enregistrement des images sur le Drive.");
      }
    }
  }

  // Creating a ColumnChart
  function createColumnChart(dataTable, title, htmlOutput, attachments, width, height) {
    try {
      // Creating a ColumnChart with data from dataTable
      let chart = Charts.newColumnChart()
        .setDataTable(dataTable)
        .setOption('legend', {
          textStyle: {
            font: 'trebuchet ms',
            fontSize: 11
          }
        })
        .setOption('colors', ["#8E3232", "#FFBE2B", "#404040", "#A29386"])
        .setTitle(title)
        .setDimensions(width, height)
        .build();

      // Adding the chart to the HtmlOutput
      let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
        imageUrl = "data:image/png;base64," + encodeURI(imageData);
      htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

      // Adding the chart to the attachments
      let imageDatamail = chart.getAs('image/png').getBytes(),
        imgblob = Utilities.newBlob(imageDatamail, "image/png", "Contacts prospection");
      attachments.push(imgblob);

      return [htmlOutput, attachments];
    } catch (e) {
      Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
  }

  // Creating a PieChart with unique values
  function createPieChartUniqueval(title, data, htmlOutput, attachments, width, height) {
    try {
      // Creating a DataTable with the proportion of each contact type
      let dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType.STRING, title);
      dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion");
      Logger.log(`Valeurs uniques : ${uniqueValues(title, data)}`);
      uniqueValues(title, data).forEach(function (val) {
        dataTable.addRow([val, data.filter(row => row[title] == val).length / data.length]);
      });

      // Creating a PieChart with data from dataTable
      let chart = Charts.newPieChart()
        .setDataTable(dataTable)
        .setOption('legend', {
          textStyle: {
            font: 'trebuchet ms',
            fontSize: 11
          }
        })
        .setOption('colors', ["#8E3232", "#FFBE2B", "#404040", "#A29386"])
        .setTitle(title)
        .setDimensions(width, height)
        .build();

      // Adding the chart to the HtmlOutput
      let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
        imageUrl = "data:image/png;base64," + encodeURI(imageData);
      htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

      // Adding the chart to the attachments
      let imageDatamail = chart.getAs('image/png').getBytes(),
        imgblob = Utilities.newBlob(imageDatamail, "image/png", title);
      attachments.push(imgblob);

      return [htmlOutput, attachments];
    } catch (e) {
      Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
  }

  // Creating a LineChart
  function createLineChart(dataTable, colors, title, htmlOutput, attachments, width, height) {
    try {
      // Creating a LineChart with data from dataTable
      let chart = Charts.newLineChart()
        .setDataTable(dataTable)
        .setOption('legend', {
          textStyle: {
            font: 'trebuchet ms',
            fontSize: 11
          }
        })
        .setOption('curveType', 'function')
        .setOption('pointShape', 'square')
        .setTitle(title)
        .setDimensions(width, height)
        .setColors(colors)
        .build();

      // Adding the chart to the HtmlOutput
      let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
        imageUrl = "data:image/png;base64," + encodeURI(imageData);
      htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

      // Adding the chart to the attachments
      let imageDatamail = chart.getAs('image/png').getBytes(),
        imgblob = Utilities.newBlob(imageDatamail, "image/png", `KPIs par mois`);
      attachments.push(imgblob);

      return [htmlOutput, attachments];
    } catch (e) {
      Logger.log(`Could not create graph for info : ${title},, error : ${e}`);
    }
  }

  // Initialization
  const ui = SpreadsheetApp.getUi(),
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suivi"),
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
    .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">KPI KPI KPI<br/></span> </span> <br/>`)
    .setWidth(1015)
    .setHeight(515);

  // Preparing the contents of the email
  let htmlMail = HtmlService.createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">&nbsp; &nbsp; Bonjour, <br/><br/>Voici les KPI tant attendus.<br/> <br/>Bonne journée !</span> </span>`),
    attachments = [],
    conversionChart = [];

  // /!\ LE PROCHAIN MANDAT REMMETTEZ MAI JUIN JUILLET (ou mettez mois + année pour couvrir plusieurs années)
  let monthList = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 0],
    monthNames = ["Jan", "Fév", "Mars", "Avr", "Mai", "Juin", "Juil", "Août", "Sept", "Oct", "Nov", "Déc"],
    stateList = ["Premier RDV réalisé", "Devis rédigé et envoyé", "En négociation", "Etude obtenue"],
    colorList = ["#8E3232", "#FFBE2B", "#404040", "#A29386"];


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
      let val = obj.filter(row => parseInt(row["Premier contact"].getMonth(), 10) == month && stateList.indexOf(row["État"]) >= stateList.indexOf(state)).length;
      if (state == "Devis rédigé et envoyé" || state == "En négociation") {
        val += obj.filter(row => parseInt(row["Premier contact"].getMonth(), 10) == month && (row["État"] == "Sans suite" || row["État"] == "A relancer") && row["Devis réalisé"]).length
      } else if (state == "Premier RDV réalisé") {
        val += obj.filter(row => parseInt(row["Premier contact"].getMonth(), 10) == month && (row["État"] == "Sans suite" || row["État"] == "A relancer")).length
      }
      conversionChartRow.push(val);
      dataRow.push(val);
    });
    conversionChart.push(conversionChartRow);
    dataTable1.addRow(dataRow);
  });
  Logger.log(conversionChart);
  // Appending the corresponding ColumnChart
  [htmlOutput, attachments] = createColumnChart(dataTable1, "Contacts", htmlOutput, attachments, 1000, 500);


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
  [htmlOutput, attachments] = createLineChart(dataTable2, ["#8E3232", "#FFBE2B", "#404040"], "Taux de conversion", htmlOutput, attachments, 1000, 500);


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
  [htmlOutput, attachments] = createColumnChart(dataTable3, "Chiffre d'affaires", htmlOutput, attachments, 1000, 500);


  /* ------------ KPI (Type de contact) ------------ */

  [htmlOutput, attachments] = createPieChartUniqueval("Type de contact", obj, htmlOutput, attachments, 1000, 500);


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
  [htmlOutput, attachments] = createColumnChart(dataTable4, "Taux de conversion par type de contact", htmlOutput, attachments, 1000, 500);

  // Saving the data
  sendMail(htmlMail, "KPI", attachments);
  saveOnDrive(attachments);

  // Final display of the charts
  ui.showModalDialog(htmlOutput, "KPI");
}
