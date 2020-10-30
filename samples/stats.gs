// ------ Aubin Tchoï ------
// --------< 최오빈 >---------
// ---- Stage @LTJ 2020 ----


// ------ Menu items ------

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Statistiques")
  .addItem("Statistiques groupées", "stats_merged_fetchAllSheets")
  .addItem("Vision par collège et classe", "stats_merged_filtered")
  .addItem("Archivage", "archiving")
  .addToUi();
}

function isNumeric(str) {
  if (typeof str != "string") {return false;} // Only processes strings
  return !isNaN(str) && // use type coercion to parse the _entirety_ of the string (`parseFloat` alone does not do this)...
    !isNaN(parseFloat(str)) // ...and ensure strings of whitespace fail
}

// Retrieves data from all sheets found in current spreadsheet
function stats_merged(filtered) {
  
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
  
  // Filters data with requested school and class
  function filter_ByClass(data, heads) {
    // Making sure data contains the informations about school and class
    if (heads.indexOf("Quel est ton collège ? ") == -1 || heads.indexOf("Quelle est ta classe ? ") == -1) {ui.alert("Filtrage par class", "Erreur : pas de questions correspondant au collège et à la classe.", ui.ButtonSet.OK); return;}
    // Query and filter
      while (true) {
      try {
        // Converting an array into a multi-indexed-line str
        function list_ToQuery(arr) {
          let query = "";
          arr.forEach(function(el, idx) {query += `\n ${idx + 1}. ${el}`});
          return query;
        }
        // Retrieving the list of all represented schools
        let school_list = unique_val("Quel est ton collège ? ", data),
            school_idx = parseInt(ui.prompt("Filtrage par classe", `Quel est le collège recherché ? \n Entrez un numéro parmi la liste suivante : ${list_ToQuery(school_list)}`, ui.ButtonSet.OK).getResponseText(), 10);
        // Retrieving the list of all represented school classes
        Logger.log(`School idx : ${school_idx}, school : ${school_list[school_idx-1]}`);
        let class_list = unique_val("Quelle est ta classe ? ", data.filter(row => row["Quel est ton collège ? "] == school_list[school_idx - 1])),
            class_idx = parseInt(ui.prompt("Filtrage par classe", `Quel est la classe recherchée ? \n Entrez un numéro parmi la liste suivante : ${list_ToQuery(class_list)}`, ui.ButtonSet.OK).getResponseText(), 10);

        // Filtering data with this couple of constraints
        if (class_idx != "" && school_idx != "") {
          return(data.filter(row => (row["Quel est ton collège ? "] == school_list[school_idx - 1] && row["Quelle est ta classe ? "] == class_list[class_idx - 1])));
        }
      } catch(e) {
        ui.alert("Filtrage par classe", "Entrée invalide, veuillez recommencer", ui.ButtonSet.OK);
        Logger.log(`An error occured when filtering given set of data : ERROR : ${e}, DATA : ${data}`);
      }
    }
  }
  
  // Retrieving the colors from the data found in 'Base créa Aubin' (returns a js object)
  function getColors(ss_id="1kTsC6pACEBGHt-9FVhWp-LQTCJUBLcnyeZQ0KwK-7H8", s_name="Base créa Aubin") {
    let colorsheet = SpreadsheetApp.openById(ss_id).getSheetByName(s_name),
        colors = colorsheet.getRange(1, 1, colorsheet.getLastRow(), 1).getBackgrounds().map(r => r[0]),
        schools = colorsheet.getRange(1, 1, colorsheet.getLastRow(), 1).getValues().map(r => r[0]),
        colorobj = {};
    schools.forEach(function(s, idx) {colorobj[s] = colors[idx];});
    
    Logger.log(`Liste des collèges : ${schools}`);
    Logger.log(`Liste des couleurs : ${colors}`);
    Logger.log(`Color set : ${JSON.stringify(colorobj, null, 4)}`);
  
    return colorobj;
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
  function send_Mail(htmlOutput, adress, subject, attachments) {
    let query = ui.alert("Envoi des diagrammes par mail", "Souhaitez vous recevoir les diagrammes par mail ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
      let msgHtml = htmlOutput.getContent(),
          msgPlain = htmlOutput.getContent().replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/ig, "");
      GmailApp.sendEmail(adress, subject, msgPlain, {htmlBody:msgHtml, attachments:attachments});
      ui.alert("Envoi des diagrammes par mail", `Les diagrammes ont été envoyés par mail à l'adresse ${adress}`, ui.ButtonSet.OK);
    }
  }
  
  // Query and save data on the Drive
  function save_onDrive(folder_id="16Un5b-wbKrObuRzNyo4ls8715Xq2OLS5", image_blobs) {
    let query = ui.alert("Enregistrement des images sur le Drive", "Souhaitez vous enregistrer les images sur le Drive ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
      display_LoadingScreen("Enregistrement des images sur le Drive..");
      let today = new Date();
      today = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
      let folder = DriveApp.getFolderById(folder_id).createFolder(`Statistiques groupées ${today}`);
      image_blobs.forEach(function(f) {folder.createFile(f);});
      ui.alert("Enregistrement des images sur le Drive", `Les images ont été enregistrées à ladresse suivante : ${folder.getUrl()}`, ui.ButtonSet.OK);
    }
  }
               
  // Writing data on a new sheet
  function rewrite(data, heads, ss) {
    let newdata = data.map(function(row) {newrow = []; heads.forEach(function(key) {newrow.push(row[key]);}); return newrow;});
    
    function write(newsheet, newdata, newheads) {
      newsheet.getRange(1, 1, 1, newheads.length)
      .setBackgroundRGB(255, 204, 229)
      .setValues([newheads]);
      newsheet.getRange(2, 1, newdata.length, newdata[0].length).setValues(newdata);
      newsheet.setColumnWidths(1, newsheet.getLastColumn(), 230);
    }

    try {
      let newsheet = ss.insertSheet("Réponses groupées");
      write(newsheet, newdata, heads);
    } catch(e) {
      let query = ui.alert("Avertissement", "L'onglet 'Réponses groupées' existe déjà, son contenu sera supprimé. \n Souhaitez vous continuer ?", ui.ButtonSet.YES_NO);
      if (query == ui.Button.YES) {
        display_LoadingScreen("Écriture des données ...");
        Logger.log(`Couldn't create sheet : ${e}`);
        ss.deleteSheet(ss.getSheetByName("Réponses groupées"));
        let newsheet = ss.insertSheet("Réponses groupées");
        write(newsheet, newdata, heads);
      }
    }
  }
      
  // Choosing how this question should be treated (PieChart, BarChart, ..)
  function question_type(question, data, heads) {
    // Excluding some questions
    if (unique_val(question, data).every(element => element == "")) {return "NoResponse";}
    
    // Less than 10 distinct values, and all values are integers
    if (unique_val(question, data).length <= 10 && unique_val(question, data).every(element => isNumeric(element))) {return "ColumnChart";}
    
    // Less than 6 distinct values (not integers)
    if (unique_val(question, data).length <= 6) {return "PieChart";}
    
    // Too many different responses, it has to be a qualitative question
    else {return "TextResponse";}
  }
  
  // Creating a PieChart
  function create_PieChart(question, data, heads, htmlOutput, attachments, width, height, is_College=false) {
    try {
      if (question == "Quelle est ta classe ? ") {return [htmlOutput, attachments];}
      
      // Creating a DataTable with the proportion of responses for each unique response to question
      let dataTable = Charts.newDataTable();
      dataTable.addColumn(Charts.ColumnType.STRING, question);
      dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion");
      
      // Adding colors for the question on schools
      if (is_College) {var colorlist = [];}
      
      unique_val(question, data).forEach(function(val) {dataTable.addRow([val, data.filter(r => r[question] == val).length/data.length]);
                                                        if (is_College) {colorlist.push(colorobj[val]);}
                                                       });
      
      // Creating a PieChart with data from dataTable
      let chart = Charts.newPieChart()
      .setDataTable(dataTable)
      .setOption('legend', {textStyle: {font: 'trebuchet ms', fontSize: 11}})
      .setTitle(question)
      .setDimensions(width, height)
      .set3D();
      
      if (is_College) {chart.setOption('colors', colorlist);}
      
      chart = chart.build();
      
      // Adding the chart to the Html output
      let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
          imageUrl = "data:image/png;base64," + encodeURI(imageData);
      htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");
      
      // Adding the chart to the attachments
      let imageDatamail = chart.getAs('image/png').getBytes(),
          imgblob = Utilities.newBlob(imageDatamail, "image/png", question);
      attachments.push(imgblob);
      
      return [htmlOutput, attachments];
    } catch(e) {Logger.log(`Could not create graph for question : ${question}`);}
  }
  
  // Creating a ColumnChart
  function create_ColumnChart(question, data, heads, htmlOutput, attachments, width, height) {
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
      
      // Adding the chart to the Html output
      let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
          imageUrl = "data:image/png;base64," + encodeURI(imageData);
      htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");
      
      // Adding the chart to the attachments
      let imageDatamail = chart.getAs('image/png').getBytes(),
          imgblob = Utilities.newBlob(imageDatamail, "image/png", question);
      attachments.push(imgblob);
      
      return [htmlOutput, attachments];
    } catch(e) {Logger.log(`Could not create graph for question : ${question}`);}
  }
  
    
  // Initialization
  const ui = SpreadsheetApp.getUi(),
      ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheets = ss.getSheets().filter(s => s.getName() != "Réponses groupées"),
      heads = sheets[0].getRange(1, 2, 1, sheets[0].getLastColumn() - 1).getDisplayValues().shift();
  
  // Colors (used in charts displaying the repartition in each school)
  let colorobj = getColors("1kTsC6pACEBGHt-9FVhWp-LQTCJUBLcnyeZQ0KwK-7H8", "Base créa Aubin");
  
  let data = [],
      resp_quali = [];
  
  // Retrieving the data
  sheets.forEach(function(s) {try {s.getRange(2, 2, (s.getLastRow() - 1), (s.getLastColumn() - 1)).getDisplayValues().forEach(function(r) {data.push(heads.reduce((o, k, i) => (o[k] = (r[i] != "") ? r[i] : o[k] || '', o), {}));});} catch(e) {Logger.log(`Couldn't retrieve data: ${e}`)}});
  
  // Filtering by class if requested
  if (filtered) {
    data = filter_ByClass(data, heads);
  }

  display_LoadingScreen("Chargement des diagrammes..");
                                                                                                                                                                   
  // Final outputs (displaying the charts on screen & mail content)  
  let htmlOutput = HtmlService
  .createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">Voici la répartition des ${data.length} réponses aux différentes questions :<br/></span> </span> <br/>`)
  .setWidth(800)
  .setHeight(465);
  
  let htmlMail = HtmlService.createHtmlOutput(`<span style='font-size: 12pt;'> <span style="font-family: 'trebuchet ms', sans-serif;">&nbsp; &nbsp; Bonjour Margaux, <br/><br/> Voici les diagrammes récapitulatifs des ${data.length} réponses aux différents questionnaires.<br/> <br/>Bonne journée !</span> </span>`),
      inlineImages = {},
      attachments = [];
 
  heads.forEach(function(question) {
    Logger.log(`Question : ${question}, type : ${question_type(question, data, heads)}`);
    if (question_type(question, data, heads) == "PieChart") {
      [htmlOutput, attachments] = create_PieChart(question, data, heads, htmlOutput, attachments, 750, 400, question == "Quel est ton collège ? ");
    } else if (question_type(question, data, heads) == "TextResponse") {
      resp_quali.push(question);
    } else if (question_type(question, data, heads) == "ColumnChart") {
      [htmlOutput, attachments] = create_ColumnChart(question, data, heads, htmlOutput, attachments, 750, 400);
    }
  });
  
  send_Mail(htmlMail, "mhecht@liketonjob.org", "Statistiques groupées", attachments);
  save_onDrive("16Un5b-wbKrObuRzNyo4ls8715Xq2OLS5", attachments);
  rewrite(data, resp_quali, ss);
  
  // Final display of the charts
  ui.showModalDialog(htmlOutput, "Réponses aux questionnaires");
}

function stats_merged_fetchAllSheets() {
  stats_merged(false);
}

function stats_merged_filtered() {
  stats_merged(true);
}

// Lets you make a copy of current spreadsheet in a given folder for archiving or backuping purposes
function archiving() {
  const ssOld = SpreadsheetApp.getActiveSpreadsheet(),
    ssNew = SpreadsheetApp.create(ssOld.getName()),
    ui = SpreadsheetApp.getUi();
    
  // Copying ssOld's content to ssNew
  ssOld.getSheets().forEach(function(s) {s.copyTo(ssNew);});
  
  // Retrieving destination folder ID
  let query = ui.alert("Archivage", "Souhaitez vous utiliser le dossier de destination par défaut (Dossier Archives)?", ui.ButtonSet.YES_NO);

  if (query == ui.Button.NO) {
    let drive_id = ui.prompt("Archivage", 
                             "Entrez l'ID du dossier Drive de destination.",
                              ui.ButtonSet.OK_CANCEL).getResponseText();
  }
  else if (query == ui.Button.YES) {
    let drive_id = "1LOk5WHg2hwLNVwuLG31625JY-nfUbHrW";
  }
  
  // Moving ssNew to the folder with drive_id
  var file = DriveApp.getFileById(ssNew.getId());
  DriveApp.getFolderById(drive_id).addFile(file);
  
  // Removing initial sheet in ssNew and renaming its sheets
  ssNew.deleteSheet(ssNew.getSheetByName("Feuille 1"));
  ssNew.getSheets().forEach(function(s) {s.setName(s.getName().match(/Copie de (.+)/)[1]);});
  
  // Removing ssOld if asked to
  let remove = ui.alert("Archivage", 
                        "Le fichier a été archivé dans le dossier, souhaitez-vous effacer son contenu ?",
                        ui.ButtonSet.YES_NO);
  if (remove == ui.Button.YES) {
    ssOld.getSheets().forEach(function(s) {
      if (s.getName().toLowerCase().indexOf("aubin") == -1) {
      s.getRange(2, 1, (s.getLastRow() - 1), s.getLastColumn())
      .clearContent()
      .setBackground('white');
      }
    });
  }
}
