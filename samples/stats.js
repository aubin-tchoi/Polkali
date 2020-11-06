// Author : Pôle Qualité 022 (Aubin Tchoï)


// Retrieves data from all sheets found in current spreadsheet
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
  
  // Filters data with requested key
  function filter_ByKey(data, heads, key) {
    // Making sure data contains the informations about key
    if (heads.indexOf(key) == -1) {ui.alert("Filtrage", "Erreur : pas de question correspondant au critère sélectionné.", ui.ButtonSet.OK); return;}
    // Query and filter
    while (true) {
      try {
        // Converting an array into a multi-indexed-line str
        function list_ToQuery(arr) {
          let query = "";
          arr.forEach(function(el, idx) {query += `\n ${idx + 1}. ${el}`});
          return query;
        }
        
        // User input
        let key_list = unique_val(key, data),
            key_idx = parseInt(ui.prompt("Filtrage", `Entrez un numéro parmi la liste suivante : ${list_ToQuery(key_list)}`, ui.ButtonSet.OK).getResponseText(), 10);
        
        if (key_idx != "") {
          return data.filter(row => row[key] == key_list[key_idx - 1]);
        }
      } catch(e) {
        ui.alert("Filtrage", "Entrée invalide, veuillez recommencer", ui.ButtonSet.OK);
        Logger.log(`An error occured when filtering given set of data : ERROR : ${e}, KEY ${key}, DATA : ${data}`);
      }
    }
  }
  
  // Loading screen
  function display_LoadingScreen(msg) {
    let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="885" height="498">`)
    .setWidth(900)
    .setHeight(500);
    ui.showModelessDialog(htmlLoading, msg);
  }
  
  // Query and send mail
  function send_Mail(htmlOutput, adress, subject, attachments, inlineImages) {
    let query = ui.alert("Envoi des diagrammes par mail", "Souhaitez vous recevoir les diagrammes par mail ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
      let msgHtml = htmlOutput.getContent(),
          msgPlain = htmlOutput.getContent().replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/ig, "");
      GmailApp.sendEmail(adress, subject, msgPlain, {htmlBody:msgHtml, attachments:attachments, inlineImages:inlineImages});
      ui.alert("Envoi des diagrammes par mail", `Les diagrammes ont été envoyés par mail à : ${adress}`, ui.ButtonSet.OK);
    }
  }
  
  // Query and save data on the Drive
  function save_onDrive(folder_id, image_blobs) {
    let query = ui.alert("Enregistrement des images sur le Drive", "Souhaitez vous enregistrer les images sur le Drive ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
      display_LoadingScreen("Enregistrement des images sur le Drive..");
      // Folder will be dated with current date
      let today = new Date();
      today = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
      let folder = DriveApp.getFolderById(folder_id).createFolder(`Statistiques groupées ${today}`);
      image_blobs.forEach(function(f) {folder.createFile(f);});
      ui.alert("Enregistrement des images sur le Drive", `Les images ont été enregistrées à ladresse suivante : ${folder.getUrl()}`, ui.ButtonSet.OK);
    }
  }
  
  // Writing data on a new sheet
  function rewrite(data, heads, ss, newname) {
    let newdata = data.map(function(row) {newrow = []; heads.forEach(function(key) {newrow.push(row[key]);}); return newrow;});
    
    function write(newsheet, newdata, newheads) {
      newsheet.getRange(1, 1, 1, newheads.length)
      .setBackgroundRGB(255, 204, 229)
      .setValues([newheads]);
      newsheet.getRange(2, 1, newdata.length, newdata[0].length).setValues(newdata);
      newsheet.setColumnWidths(1, newsheet.getLastColumn(), 230);
    }

    try {
      let newsheet = ss.insertSheet(newname);
      write(newsheet, newdata, heads);
    } catch(e) {
      let query = ui.alert("Avertissement", `Un onglet au nom ${newname} existe déjà, son contenu sera supprimé. \n Souhaitez vous continuer ?`, ui.ButtonSet.YES_NO);
      if (query == ui.Button.YES) {
        display_LoadingScreen("Écriture des données ...");
        Logger.log(`Failed to create sheet : ${e}`);
        ss.deleteSheet(ss.getSheetByName(newname));
        let newsheet = ss.insertSheet(newname);
        write(newsheet, newdata, heads);
      }
    }
  }
      
  // Choosing how a question should be treated (PieChart, BarChart, ..), you can exclude a question by returning something else here
  function question_type(question, data) {
    // Empty column
    if (unique_val(question, data).every(element => element == "")) {return "TextResponse"};

    // Less than 10 distinct values, and all values are integers
    if (unique_val(question, data).length/data.length <= 0.25 && unique_val(question, data).every(element => isNumeric(element))) {return "ColumnChart";}
    
    // Less than 6 distinct values (not integers)
    if (unique_val(question, data).length/data.length <= 0.25) {return "PieChart";}
    
    // Too many different responses, it has to be a qualitative question
    else {return "TextResponse";}
  }
  
  // Creating a PieChart
  function create_PieChart(question, data, htmlOutput, attachments, inlineImages, width, height) {
    try {
      
      // Creating a DataTable with the proportion of responses for each unique response to question
      let dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType.STRING, question);
      dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion");
      
      unique_val(question, data).forEach(function(val) {dataTable.addRow([val, data.filter(r => r[question] == val).length/data.length]);});
      
      // Creating a PieChart with data from dataTable
      let chart = Charts.newPieChart()
      .setDataTable(dataTable)
      .setOption('legend', {textStyle: {font: 'trebuchet ms', fontSize: 11}})
      .setTitle(question)
      .setDimensions(width, height)
      .set3D()
      .build();

      // Putting the image into a blob
      let cid = question.replace(" ", "_"),
          imageData = chart.getAs('image/png').getBytes(),
          imgblob = Utilities.newBlob(imageData, "image/png", cid);
      
      // Adding the chart to inlineImages
      inlineImages[cid] = imgblob;
      
      // Adding the chart to the Html output
      htmlOutput.append(`<img src="cid:${cid}"/><br/>`);
      
      // Adding the chart to the attachments
      attachments.push(imgblob);
      
      return [htmlOutput, attachments, inlineImages];
    } catch(e) {Logger.log(`Could not create graph for question : ${question}`);}
  }
  
  // Creating a ColumnChart
  function create_ColumnChart(question, data, htmlOutput, attachments, inlineImages, width, height) {
    try {
      // Creating a DataTable with the proportion of responses for each unique response to question
      let dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType.STRING, question);
      dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion");
      
      unique_val(question, data).sort().forEach(function(val) {dataTable.addRow([val, data.filter(r => r[question] == val).length/data.length]);});
      
      // Creating a ColumnChart with data from dataTable
      let chart = Charts.newColumnChart()
      .setDataTable(dataTable)
      .setOption('legend', {textStyle: {font: 'trebuchet ms', fontSize: 11}})
      .setTitle(question)
      .setDimensions(width, height)
      .build();
      
      // Putting the image into a blob
      let cid = question.replace(" ", "_"),
          imageData = chart.getAs('image/png').getBytes(),
          imgblob = Utilities.newBlob(imageData, "image/png", cid);
      
      // Adding the chart to inlineImages
      inlineImages[cid] = imgblob;
      
      // Adding the chart to the Html output
      htmlOutput.append(`<img src="cid:${cid}"/><br/>`);
      
      // Adding the chart to the attachments
      attachments.push(imgblob);
      
      return [htmlOutput, attachments, inlineImages];
    } catch(e) {Logger.log(`Could not create graph for question : ${question}`);}
  }
  
    
  // Initialization
  const ui = SpreadsheetApp.getUi(),
      ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheets = ss.getSheets().filter(s => s.getName() != "Statistiques groupées"),
      heads = sheets[0].getRange(1, 2, 1, sheets[0].getLastColumn() - 1).getDisplayValues().shift();
  
  let data = [],
      resp_quali = [];
  
  // Retrieving the data
  sheets.forEach(function(s) {try {s.getRange(2, 2, (s.getLastRow() - 1), (s.getLastColumn() - 1)).getDisplayValues().forEach(function(r) {data.push(heads.reduce((o, k, i) => (o[k] = (r[i] != "") ? r[i] : o[k] || '', o), {}));});} catch(e) {Logger.log(`Couldn't retrieve data: ${e}`)}});
  
  // Filtering by class if requested
  if (filtered) {
    let key = ui.prompt("Filtrage des données", "Suivant quelle information souhaitez-vous filtrer les données ? (Entrez un nom de colonne)", ui.ButtonSet.OK).getResponseText();
    data = filter_ByKey(data, heads, key);
  }

  display_LoadingScreen("Chargement des diagrammes..");
                                                                                                                                                                   
  // Final outputs (displaying the charts on screen & mail content)  
  let htmlOutput = HtmlService
  .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">Voici la répartition des ${data.length} lignes de données :<br/></span> </span> <br/>`)
  .setWidth(800)
  .setHeight(465);
  
  let htmlMail = HtmlService.createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">&nbsp; &nbsp; Bonjour, <br/><br/> Voici les diagrammes récapitulatifs des ${data.length} lignes de données.<br/> <br/>Bonne journée !</span> </span>`),
      inlineImages = {},
      attachments = [];
 
  heads.forEach(function(question) {
    Logger.log(`Question : ${question}, type : ${question_type(question, data)}`);
    if (question_type(question, data) == "PieChart") {
      [htmlOutput, attachments, inlineImages] = create_PieChart(question, data, htmlOutput, attachments, inlineImages, 750, 400, question == "Quel est ton collège ? ");
    } else if (question_type(question, data) == "TextResponse") {
      resp_quali.push(question);
    } else if (question_type(question, data) == "ColumnChart") {
      [htmlOutput, attachments, inlineImages] = create_ColumnChart(question, data, htmlOutput, attachments, inlineImages, 750, 400);
    }
  });
  
  send_Mail(htmlMail, "", "Statistiques groupées", attachments, inlineImages);
  save_onDrive("", attachments);
  rewrite(data, resp_quali, ss, "Statistiques groupées");
  
  // Final display of the charts
  ui.showModalDialog(htmlOutput, "Réponses aux questionnaires");
}

function stats_merged_fetchAllSheets() {
  stats_merged(false);
}

function stats_merged_filtered() {
  stats_merged(true);
}
