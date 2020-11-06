// Author : Pôle Qualité 022 (Aubin Tchoï)

// Two parts in this script : updating the data from (hidden) sheet "Base calcul KPI" and then using this set of data to produce some graphs (current month + last month, evolution)

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('KPI')
  .addItem("Obtention des graphes", "sync_onSelec")
  .addToUi();
}

// Returning the last non-empty row of a sheet (not possible with getLastRow for this particular sheet bc of data validation)
function manually_getLastRow(sheet) {
  let data = sheet.getRange(1, 1, sheet.getLastRow(), 5).getValues();
  for (let row=1;row<sheet.getLastRow();row++) {
    if ((data[row][0] == "" && data[row][1] == "" && data[row][2] == "" && data[row][4] == "")|| row > 200) {
      return row;
    }
  }
}

function onEdit(e) {
  let range = e.range,
    colidx_ContactType = 3;
  if (range.getNumColumns() == 1 && range.getNumRows() == 1) {

  }
  // Check here if getColumn() returns an A1 notation or an 0-array one
  // Data in sheet "Suivi" have been modified
  if (range.getSheet().getName() == "Suivi") {
    
    if (range.getColumn() > 1 && range.getColumn() <= colidx_ContactType && (range.getColumn() + range.getNumColumns() - 1) >= colidx_ContactType) {
      
    }
  }
}


function sync_onEdit(rowidx) {
  // Returning the last non-empty row of a sheet (not possible with getLastRow for this particular sheet for some unkown reasons)
  function manually_getLastRow(sheet) {
    let data = sheet.getRange(1, 1, sheet.getLastRow(), 5).getValues();
    for (let row=1;row<sheet.getLastRow();row++) {
      if ((data[row][0] == "" && data[row][1] == "" && data[row][2] == "" && data[row][4] == "")|| row > 200) {
        return row;
      }
    }
  }
  
  // Updating sheet sheet with row row
  function update(sheet, row) {
    // Checking whether PEP actually did obtain the mission or not
    if (row["État"] != "Etude obtenue ") {
      return;
    }
    
    // ref is gonna be the mission's reference, obtained by incrementing the most recent one
    const heads = sheet.getRange(1, 2, 1, (sheet.getLastColumn() - 1)).getValues().shift(),
        last_row = manually_getLastRow(sheet),
        ref = `'20e${Math.floor(sheet.getRange(last_row, 2, 1, 1).getValues()[0][0].toString().match(/([0-9]+)/gi)[1]) + 1}`;
    
    // Creating a new row with the destination sheet's header's informations
    let rnew = [row["Entreprise"]]; heads.forEach(function(el) {rnew.push(el == "Référence de l'étude" ? ref : (el == "État") ? "Devis accepté" : row[el]);});
    sheet.getRange((last_row + 1), 1, 1, rnew.length).setValues([rnew]);
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


function sync_onSelec() {
  // Updating sheet sheet with row row
  function update(sheet, row) {
    // Checking whether PEP actually did obtain the mission or not
    if (row["État"] != "Etude obtenue ") {
      ui.alert("Entrée invalide", `L'étude confiée par l'entreprise ${row["Entreprise"]} ne correspond pas à une étude obtenue.`, ui.ButtonSet.OK);
      return;
    }
    
    // ref is gonna be the mission's reference, obtained by incrementing the most recent one
    const heads = sheet.getRange(1, 2, 1, (sheet.getLastColumn() - 1)).getValues().shift(),
        last_row = manually_getLastRow(sheet),
        ref = `'20e${Math.floor(sheet.getRange(last_row, 2, 1, 1).getValues()[0][0].toString().match(/([0-9]+)/gi)[1]) + 1}`;
    
    // Creating a new row with the destination sheet's header's informations
    let rnew = [row["Entreprise"]]; heads.forEach(function(el) {rnew.push(el == "Référence de l'étude" ? ref : (el == "État") ? "Devis accepté" : row[el]);});
    sheet.getRange((last_row + 1), 1, 1, rnew.length).setValues([rnew]);
  }
  
  // Displaying a loading screen
  function display_LoadingScreen() {
    let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://www.demilked.com/magazine/wp-content/uploads/2016/06/gif-animations-replace-loading-screen-14.gif" alt="Loading" width="531" height="299">`)
    .setWidth(540)
    .setHeight(350);
    ui.showModelessDialog(htmlLoading, "Synchronisation des données.."); // Not clean bc ui is not defined here but who cares
  }
  
  // -- MAIN --
  const sheetscr = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
      sheetdst = SpreadsheetApp.openById("1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04").getSheetByName("Suivi"),
      ui = SpreadsheetApp.getUi(),
      htmlLoading = HtmlService
  .createHtmlOutput(`<img src="https://www.demilked.com/magazine/wp-content/uploads/2016/06/gif-animations-replace-loading-screen-14.gif" alt="Loading" width="531" height="299">`)
  .setWidth(540)
  .setHeight(350);
  
  // Confirm selection
  let confirm_selection = ui.alert("Synchronisation des données", "Vous devez préalablement sélectionner la ligne à synchroniser (par exemple en cliquant sur le numéro à gauche). \n Confirmez-vous votre sélection ?", ui.ButtonSet.YES_NO);
  if (confirm_selection == ui.Button.NO) {return;}
  
  // Loading screen
  display_LoadingScreen();
  
  // heads is the sheet's header, data a 2D-array representation of the selected values, and obj its 'array of js object' representation
  let heads = sheetscr.getRange(1, 1, 1, sheetscr.getLastColumn()).getValues().shift(),
      data = sheetscr.getActiveRange().getDisplayValues(),
      obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  
  // Checking the selection again (empty or not)
  if (obj.length == 0) {ui.alert("Pas de ligne sélectionnée, veuillez recommencer."); return;}
  
  // Synchronizing each selected row
  obj.forEach(function(row) {
    update(sheetdst, row);
  });
  
  // Confirmation
  let img_url = "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/Ddabong.png",
      sheet_url = "https://docs.google.com/spreadsheets/d/1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04/edit#gid=0",
      operation_success = HtmlService
  .createHtmlOutput(`<span style='font-size: 16pt;'> <span style="font-family: 'roboto', sans-serif;">L'opération a été effectuée avec succès, veuillez remplir manuellement le nom de l'étude dans le <a href="${sheet_url}">suivi des études</a>.<br/><br/> La bise</span></span><p style="text-align:center;"><img src="${img_url}" alt="C'est la PEP qui régale !" width="130" height="131"></p>`);
  ui.showModalDialog(operation_success, "Synchronisation des données");
}


function stats_merged(filtered) {
  // Returns true if str can be converted into a float
  function isNumeric(str) {
    if (typeof str != "string") {return false;} // Only processes strings
    return !isNaN(str) && // use type coercion to parse the _entirety_ of the string (`parseFloat` alone does not do this)...
      !isNaN(parseFloat(str)) // ...and ensure strings of whitespace fail
  }

  // Getting an array of all unique values inside of a set of data for 1 information (heads.indexOf(str) being the index of the column that contains the data inside of each row)
  function unique_val(str, data) {
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
  
  // Loading screen
  function display_LoadingScreen(msg) {
    let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="885" height="498">`)
    .setWidth(900)
    .setHeight(500);
    ui.showModelessDialog(htmlLoading, msg);
  }
  
  // Query and send mail to designated adress
  function send_Mail(htmlOutput, subject, attachments, inlineImages) {
    let query = ui.alert("Envoi des diagrammes par mail", "Souhaitez vous recevoir les diagrammes par mail ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
      let adress = ui.prompt("Envoi des diagrammes par mail", "Entrez l'adresse mail de destination :", ui.ButtonSet.OK).getResponseText(),
          msgHtml = htmlOutput.getContent(),
          msgPlain = htmlOutput.getContent().replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/ig, "");
      GmailApp.sendEmail(adress, subject, msgPlain, {htmlBody:msgHtml, attachments:attachments, inlineImages:inlineImages});
      ui.alert("Envoi des diagrammes par mail", `Les diagrammes ont été envoyés par mail à : ${adress}.`, ui.ButtonSet.OK);
    }
  }
  
  // Query and save data in designated Drive folder
  function save_onDrive(image_blobs) {
    let query = ui.alert("Enregistrement des images sur le Drive", "Souhaitez vous enregistrer les images sur le Drive ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
      try {
        let folder_id = ui.prompt("Enregistrement des images sur le Drive", "Entrez l'id du dossier de destination :", ui.ButtonSet.OK).getResponseText();
        display_LoadingScreen("Enregistrement des images sur le Drive..");
        // Folder will be dated with current date
        let today = new Date();
        today = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
        let folder = DriveApp.getFolderById(folder_id).createFolder(`Statistiques groupées ${today}`);
        image_blobs.forEach(function(f) {folder.createFile(f);});
        ui.alert("Enregistrement des images sur le Drive", `Les images ont été enregistrées à ladresse suivante : ${folder.getUrl()}`, ui.ButtonSet.OK);
      }
      catch(e) {
        Logger.log(`Erreur lors de l'enregistrement des images sur le Drive : ${e}.`);
        ui.alert("Erreur lors de l'enregistrement des images sur le Drive.");
      }
    }
  }
  
  // Creating a ColumnChart
  function create_ColumnChart(infos, data, htmlOutput, attachments, inlineImages, width, height) {
    try {
      // Creating a DataTable
      let dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
      infos.forEach(function(info) {dataTable.addColumn(Charts.ColumnType.NUMBER, info);});
      
      data.forEach(function(row) {let dataRow = [row["Mois"]];
                                  infos.forEach(function(info) {dataRow.push(row[info]);});
                                  dataTable.addRow(dataRow)});
      
      // Creating a ColumnChart with data from dataTable
      let chart = Charts.newColumnChart()
      .setDataTable(dataTable)
      .setOption('legend', {textStyle: {font: 'trebuchet ms', fontSize: 11}})
      .setTitle("Contacts prospection")
      .setDimensions(width, height)
      .set3D()
      .build();

      // Putting the image into a blob
      let cid = infos[0].replace(" ", "_"),
          imageData = chart.getAs('image/png').getBytes(),
          imgblob = Utilities.newBlob(imageData, "image/png", cid);
      
      // Adding the chart to inlineImages
      inlineImages[cid] = imgblob;
      
      // Adding the chart to the Html output
      htmlOutput.append(`<img src="cid:${cid}"/><br/>`);
      
      // Adding the chart to the attachments
      attachments.push(imgblob);
      
      return [htmlOutput, attachments, inlineImages];
    } catch(e) {Logger.log(`Could not create graph for infos : ${infos}`);}
  }
  
  // Creating a PieChart
  function create_PieChart(info, data, htmlOutput, attachments, inlineImages, width, height) {
    try {
      // Creating a DataTable
      let dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType.STRING, "Tranche de prix");
      dataTable.addColumn(Charts.ColumnType.NUMBER, info);
      
      data.forEach(function(row) {dataTable.addRow([row["Tranche de prix"], row[info]]);});
      
      // Creating a PieChart with data from dataTable
      let chart = Charts.newPieChart()
      .setDataTable(dataTable)
      .setOption('legend', {textStyle: {font: 'trebuchet ms', fontSize: 11}})
      .setTitle(`${info} par tranche de prix`)
      .setDimensions(width, height)
      .set3D()
      .build();
      
      // Putting the image into a blob
      let cid = info.replace(" ", "_"),
          imageData = chart.getAs('image/png').getBytes(),
          imgblob = Utilities.newBlob(imageData, "image/png", cid);
      
      // Adding the chart to inlineImages
      inlineImages[cid] = imgblob;
      
      // Adding the chart to the Html output
      htmlOutput.append(`<img src="cid:${cid}"/><br/>`);
      
      // Adding the chart to the attachments
      attachments.push(imgblob);
      
      return [htmlOutput, attachments, inlineImages];
    } catch(e) {Logger.log(`Could not create graph for data : ${info}`);}
  }
  
  // Creating a LineChart
  function create_LineChart(info, data, htmlOutput, attachments, inlineImages, width, height) {
    try {
      // Creating a DataTable
      let dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
      dataTable.addColumn(Charts.ColumnType.NUMBER, info);
      
      data.forEach(function(row) {dataTable.addRow([row["Mois"], row[info]]);});
      
      // Creating a LineChart with data from dataTable
      let chart = Charts.newLineChart()
      .setDataTable(dataTable)
      .setOption('legend', {textStyle: {font: 'trebuchet ms', fontSize: 11}})
      .setTitle(`${info} par mois`)
      .setDimensions(width, height)
      .set3D()
      .build();
      
      // Putting the image into a blob
      let cid = info.replace(" ", "_"),
          imageData = chart.getAs('image/png').getBytes(),
          imgblob = Utilities.newBlob(imageData, "image/png", cid);
      
      // Adding the chart to inlineImages
      inlineImages[cid] = imgblob;
      
      // Adding the chart to the Html output
      htmlOutput.append(`<img src="cid:${cid}"/><br/>`);
      
      // Adding the chart to the attachments
      attachments.push(imgblob);
      
      return [htmlOutput, attachments, inlineImages];
    } catch(e) {Logger.log(`Could not create graph for data : ${info}`);}
  }
    
  // Initialization
  const ui = SpreadsheetApp.getUi(),
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base calcul KPI"),
      heads1 = sheet.getRange(1,  1, 1, 7).getDisplayValues().shift(),
      heads2 = sheet.getRange(1, 11, 1, 3).getDisplayValues().shift();
  
  // I chose to keep the data as an array of objects (other possibility : object of objects, the keys being the months' names and the values the data in each row)
  let data1 = sheet.getRange(1,  1, sheet.getLastRow(), 7).getDisplayValues().map(r => heads1.reduce((o, k, i) => (o[k] = (r[i] != "") ? r[i] : o[k] || '', o), {})),
      data2 = sheet.getRange(1, 11, sheet.getLastRow(), 3).getDisplayValues().map(r => heads2.reduce((o, k, i) => (o[k] = (r[i] != "") ? r[i] : o[k] || '', o), {}));
  
  // Getting rid of the first column
  heads1.shift(); heads2.shift();

  display_LoadingScreen("Chargement des diagrammes..");
                                                                                                                                                                   
  // Final outputs (displaying the charts on screen & mail content)  
  let htmlOutput = HtmlService
  .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">Comme promis, voici les KPI tant attendus :<br/></span> </span> <br/>`)
  .setWidth(800)
  .setHeight(465);
  
  let htmlMail = HtmlService.createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">&nbsp; &nbsp; Bonjour, <br/><br/> Voici les diagrammes récapitulatifs des ${data.length} lignes de données.<br/> <br/>Bonne journée !</span> </span>`),
      inlineImages = {},
      attachments = [];
  
  // Creating ColumnCharts (by month) of the following informations : "Nombre de contacts", "Contact: Sans Suite", "Conversion en Étude", "Nombre de devis réalisés"
  [htmlOutput, attachments, inlineImages] = create_ColumnChart([heads[0], heads[1], heads[2], heads[3]], data1, htmlOutput, attachments, inlineImages, 750, 400);
  
  // Creating a LineChart (by month) of the foloowing information : "Taux de conversion"
  [htmlOutput, attachments, inlineImages] = create_LineChart(heads[4], data1, htmlOutput, attachments, inlineImages, 750, 400);

  // Creating ColumnCharts (by month) of the following informations : "CA sur étude potentielle", "CA réalisé"
  [htmlOutput, attachments, inlineImages] = create_ColumnChart([heads[5], heads[6]], data1, htmlOutput, attachments, inlineImages, 750, 400);
  
  // Creating PieCharts (classified by price range) of the following informations : "Pourcentage de devis acceptés", "Pourcentage de devis refusés"
  heads2.forEach(function(info) {[htmlOutput, attachments, inlineImages] = create_PieChart(info, data2, htmlOutput, attachments, inlineImages, 750, 400);});

  send_Mail(htmlMail, "KPI", attachments);
  save_onDrive(attachments);
  
  // Final display of the charts
  ui.showModalDialog(htmlOutput, "KPI");
}
