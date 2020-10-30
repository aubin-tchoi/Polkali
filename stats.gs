// Author : Pôle Qualité 022 (Aubin Tchoï)

// Retrieves data from all sheets found in current spreadsheet
function stats_merged_fetchAllSheets() {
  
  // Getting an array of all unique values inside of a set of data for 1 information (heads.indexOf(str) being the index of the column that contains the data inside of each row)
  function unique_val(str, data, heads) {
    let val = []; for (let row = 0; row < data.length; row++) {
      if (data[row][heads.indexOf(str)] != "") {
        if (val.indexOf(data[row][heads.indexOf(str)]) == -1) {
          val.push(data[row][heads.indexOf(str)]);
        }
      }
    }
    return val;
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
    .createHtmlOutput(`<img src="https://www.demilked.com/magazine/wp-content/uploads/2016/06/gif-animations-replace-loading-screen-14.gif" alt="Loading" width="885" height="498">`)
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
  function rewrite(data, heads, newidx, ss) {
    let newheads = []; newidx.forEach(function(idx) {newheads.push(heads[idx]);});
    let newdata = data.map(function(row) {newrow = []; row.forEach(function(value, index) {if (newidx.indexOf(index) != -1) {newrow.push(row[index]);}}); return newrow;});
    
    try {
      let newsheet = ss.insertSheet("Réponses groupées");
    } catch(e) {
      let query = ui.alert("Avertissement", "L'onglet 'Réponses groupées' existe déjà, son contenu sera supprimé. \n Souhaitez vous continuer ?", ui.ButtonSet.YES_NO);
      if (query == ui.Button.YES) {
        display_LoadingScreen("Écriture des données ...");
        Logger.log(`Couldn't create sheet : ${e}`);
        ss.deleteSheet(ss.getSheetByName("Réponses groupées"));
        let newsheet = ss.insertSheet("Réponses groupées");
        newsheet.getRange(1, 1, 1, newheads.length)
        .setBackgroundRGB(255, 204, 229)
        .setValues([newheads]);
        newsheet.getRange(2, 1, newdata.length, newdata[0].length).setValues(newdata);
        newsheet.setColumnWidths(1, newsheet.getLastColumn(), 230);
      }
    }
  }
      
  // Choosing how this question should be treated (PieChart, BarChart, ..)
  function question_type(question, data, heads) {
    // Less than 10 distinct values, and all values are integers
    if (unique_val(question, data, heads).length <= 10 && unique_val(question, data, heads).every(element => isNumeric(element))) {return "ColumnChart";}
    
    // Less than 6 distinct values (not integers)
    if (unique_val(question, data, heads).length <= 6) {return "PieChart";}
    
    // Too many different responses, it has to be a qualitative question
    else {return "TextResponse";}
  }
  
  // Creating a PieChart
  function create_PieChart(question, data, heads, htmlOutput, attachments, width, height, is_College=false) {
    let dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType.STRING, question);
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion");
    
    if (is_College) {var colorlist = [];}
    
    unique_val(question, data, heads).forEach(function(val) {dataTable.addRow([val, data.filter(r => r[heads.indexOf(question)] == val).length/data.length]);
                                                       if (is_College) {colorlist.push(colorobj[val]);}
                                                      });
    
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
  }
  
  // Creating a ColumnChart
  function create_ColumnChart(question, data, heads, htmlOutput, attachments, width, height) {
    let dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType.STRING, question);
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion");
    
    unique_val(question, data, heads).sort().forEach(function(val) {dataTable.addRow([val, data.filter(r => r[heads.indexOf(question)] == val).length/data.length]);});
    
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
  
  display_LoadingScreen("Chargement des diagrammes..");
  
  // Retrieving the data
  sheets.forEach(function(s) {try {s.getRange(2, 2, s.getLastRow() - 1, s.getLastColumn() - 1).getDisplayValues().forEach(function(r) {data.push(r);});} catch(e) {Logger.log(`Couldn't retrieve data: ${e}`)}});
  
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
      resp_quali.push(heads.indexOf(question));
    } else if (question_type(question, data, heads) == "ColumnChart") {
      [htmlOutput, attachments] = create_ColumnChart(question, data, heads, htmlOutput, attachments, 750, 400);
    }
  });
  
  
  send_Mail(htmlMail, "", "Statistiques groupées", attachments);
  save_onDrive("", attachments);
  rewrite(data, heads, resp_quali, ss);
  
  // Final display of the charts
  ui.showModalDialog(htmlOutput, "Réponses aux questionnaires");
}
