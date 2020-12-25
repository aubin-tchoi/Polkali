/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

/* Must be bound to a spreadsheet that contains a sheet name "Dashboard" in which are given a list of mail adresses
  beginning in cell B5 (1 row per adress). Cells B4 to E4 should contain this header : ["Date d'envoi", "Prénom", "Nom", "Adresse mail"]
  This sheet must contain a cell "Total" next to which will be recorded the votes.
  The name of the Gmail Draft used is read from cell B1 */

/* There are only 2 variables to adapt : folderId and spreadsheetId, the folder to this Id must contain 1 folder for each group. */

const folderId = "1nHfPZR10ZCFx-Nro546WjRwAmih2_bfq",
    spreadsheetId = "11gnPSgy85927z95yxoXb_iS6B2YNoTZumkQooWndjgE";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bracket')
  .addItem("Génération du Forms", "stinson")
  .addToUi();
}

// Detects the position of an element in a set of data
function detectPos(data, element) {
  let row = (data.indexOf(data.filter(r => r.includes(element))[0]) + 1),
    column = (data[row - 1].indexOf(element) + 1);
  return [row, column];
}

function stinson() {
  // Loading screen
  function displayLoadingScreen(msg) {
    let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="442" height="249">`)
    .setWidth(450)
    .setHeight(325);
    SpreadsheetApp.getUi().showModelessDialog(htmlLoading, msg);
  }
    
  // Adds a line to the dashboard
  function formatSheet(spreadsheet, poulNum, poulSize) {
    const sheet = spreadsheet.getSheetByName("Dashboard"),
        values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues(),
        [startRow, startColumn] = detectPos(values, "Total");
    
    Logger.log(`Row offset : ${startRow}`);
    Logger.log(`Column offset : ${startColumn}`);

    // First line
    for (let poule = 1; poule <= poulNum; poule++) {
      let poulValues = sheet.getRange(startRow, (startColumn + 1 + (poule - 1) * poulSize), 1, poulSize).getValues(); poulValues[0][0] = `Poule ${poule}`;
      sheet.getRange(startRow, (startColumn + 1 + (poule - 1) * poulSize), 1, poulSize)
      .setValues(poulValues)
      .setBackground("#ffe599")
      .setHorizontalAlignments(poulValues.map(row => row.map(() => "center")))
      .merge();
      sheet.getRange((startRow + 1), (startColumn + 1 + (poule - 1) * poulSize), 1, poulSize).setValues([Array.from(Array(poulSize).keys()).map(x => x + 1)]);
      sheet.setColumnWidths((startColumn + 1), poulNum * poulSize, 23);
    }
    return [startRow, startColumn];
  }

  // Sends an email containing the link to every email adress found in the sheet
  function mailLink(spreadsheet, link) {
    // Returns a boolean function meant to be used in a filter function
    function subjectFilter_(template){
      return function(element) {
        return element.getMessage().getSubject() === template;
      }
    }
    
    // Retrieving inline images inside of the html body
    function getInlineImages(msg) {
      let msgHtml = msg.getBody(),
        rawc = msg.getRawContent(),
        imgTags = (msgHtml.match(/<img[^>]+>/g) || []),
        inlineImages = {};
      for (let i = 0; i < imgTags.length; i++) {
        let realattid = imgTags[i].match(/cid:(.*?)"/i);
        if (realattid) {
          let cid = realattid[1];
            imgTagNew = imgTags[i].replace(/src="[^\"]+\"/, "src=\"cid:" + cid + "\"");
          msgHtml = msgHtml.replace(imgTags[i], imgTagNew);
          let b64c1 = (rawc.lastIndexOf(cid) + cid.length + 3),
              b64cn = (rawc.substring(b64c1).indexOf("--") - 3),
              imgb64 = rawc.substring(b64c1, b64c1 + b64cn + 1),
              imgblob = Utilities.newBlob(Utilities.base64Decode(imgb64), "image/jpeg", cid);
          inlineImages[cid] = imgblob;
        }
      }
      return [msgHtml, inlineImages];
    }
    
    function sendMail(templateName, link, row) {
      if (row["Date d'envoi"] == "") {
          try {
          let msg = GmailApp.getDrafts().filter(subjectFilter_(templateName))[0].getMessage();
            
          // Inline images
          let [msgHtml, inlineImages] = getInlineImages(msg);

          // Customizing the model with data taken from row
          heads.forEach(function(key) {
            let regexp = new RegExp(`{${key}}|{${key.toLowerCase()}}|{${key.replace(/[éêè]/gi, "e")}}`, "gmi");
            msgHtml = msgHtml.replace(regexp, row[key]);
          });
          
          // Adding the form's link to the mail
          msgHtml = msgHtml.replace(/{link}/gmi, `<a href = ${link}>Bracket</a>`);

          // Sending the mail
          let msgPlain = msgHtml.replace(/\<br\/\>/gmi, '\n').replace(/(<([^>]+)>)/gmi, "");
          Logger.log(msgHtml); Logger.log(msgPlain);
          GmailApp.sendEmail(row["Adresse mail"], "Participez au bracket !", msgPlain, {htmlBody:msgHtml, inlineImages:inlineImages});
          return [new Date()];
        }
        catch(e) {
          Logger.log(`Issue with row ${row} : ${e}`);
          return [(e.message == "Cannot read property 'getMessage' of undefined") ? "Pas de modèle à ce nom" : e.message];
        }
      }
      else {
        return [row["Date d'envoi"]];
      }
    }

    const sheet = spreadsheet.getSheetByName("Dashboard"),
      templateName = sheet.getRange(1, 2, 1, 1).getValues()[0][0];
    let data = sheet.getRange(4, 2, (sheet.getLastRow() - 3), 4).getValues(),
      heads = data.shift(),
      output = [],
      nSent = data.filter(row => row[0] == "").length;
    data = data.map(row => heads.reduce((o, k, i) => (o[k] = (row[i] != "") ? row[i] : o[k] || '', o), {})).filter(row => row["Adresse mail"] != "");

    Logger.log(`Heads : ${heads}`);
    Logger.log(`Template name : ${templateName}`);

    // Sending a mail to each email adress
    data.forEach(function(row) {
        output.push(sendMail(templateName, link, row));
    });

    sheet.getRange(5, 2, output.length, 1).setValues(output);
    return nSent;
  }

  // Pushing a file into folder and trashing it
  function moveFileToFolder(fileId, name, folder) {
    let file = DriveApp.getFileById(fileId),
      copy = file.makeCopy(name, folder);
    file.setTrashed(true);
    return copy.getId();
  }

  // Generates a Google Forms representation of the bracket
  function genForms(folderId, spreadsheet) {
    // Adding an image+mark to the forms
    function fillSection(forms, image, poulNum, poulSize) {
      let imageItem = forms.addImageItem()
        .setImage(image)
        .setTitle(`${poulSize}${poulSize >= 2 ? "ème" : "er"} candidat de la poule ${poulNum}`)
        .setWidth(300);
      let mark = forms.addTextItem()
        .setTitle('Votre note')
        .setRequired(true);
    }    

    // First questions
    const folder = DriveApp.getFolderById(folderId);
    let forms = FormApp.create("Bracket")
      .setDescription(`Bienvenue dans le bracket, pour chaque poule vous disposez de 10 points à répartir sur l'ensemble des candidats.
      Votre vote ne sera pas pris en compte si vous dépensez plus de 10 points. Soyez avisés.`)
      .setConfirmationMessage(`Merci pour votre participation, la direction du Bracket vous recontactera sous peu pour annoncer les résultats.`),
      firstName = forms.addTextItem().setTitle("Quel est votre prénom ?").setRequired(true),
      lastName = forms.addTextItem().setTitle("Quel est votre nom ?").setRequired(true);

    let folders = folder.getFolders(),
      poulNum = 0,
      poulSize = 0;
    
    // Looping on each folder (1 section for each folder, folder <=> poule)
    while (folders.hasNext()) {
      let subFolder = folders.next(),
        files = subFolder.getFiles(),
        section = forms.addPageBreakItem()
        .setTitle(`Poule ${++poulNum}`);
      poulSize = 0;

        // Looping on each file (adding the image + a question for each file)
      while (files.hasNext()) { 
        let image = files.next().getBlob();
        fillSection(forms, image, poulNum, ++poulSize);
      }
    }

    // Setting my dashboard as a destination
    forms = FormApp.openById(moveFileToFolder(forms.getId(), "Bracket", folder));
    forms.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    return [forms.getEditUrl(), forms.getPublishedUrl(), poulNum, poulSize];
  }

  // Loading screen
  displayLoadingScreen("Génération du Forms ..")

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId),
    [editLink, publishedLink, poulNum, poulSize] = genForms(folderId, spreadsheet),
    [startRow, startColumn] = formatSheet(spreadsheet, poulNum, poulSize),
    ui = SpreadsheetApp.getUi();
  let nSent = 0;

  if (ui.alert("Bracket", "Souhaitez-vous envoyer le lien du Google Form par mail aux participants ?", ui.ButtonSet.YES_NO) == ui.Button.YES) {
    nSent = mailLink(spreadsheet, publishedLink);
  }

  // Adding a trigger to the dashboard to update it after each vote
  ScriptApp.newTrigger("updateSheet")
  .forSpreadsheet(spreadsheet)
  .onFormSubmit()
  .create();

  // Confirmation message (also ends the loading screen)
  let htmlOutput = HtmlService.createHtmlOutput(`<span style="font-family: 'trebuchet ms', sans-serif;">Voci le lien éditeur : <a href = "${editLink}">éditer le Form</a>.<br/>
                                                <br/> Voici le lien lecteur : <a href = "${publishedLink}">répondre au Form</a>.<br/>
                                                <br/> Le lien lecteur a été envoyé à ${nSent} personne${nSent >= 2 ? "s" : ""}.</span>`);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, "Succès de l'opération")
}

function updateSheet() {
  const dataSheet = SpreadsheetApp.openById(spreadsheetId).getSheets()[0],
    data = dataSheet.getRange(2, 4, (dataSheet.getLastRow() - 1), (dataSheet.getLastColumn() - 3)).getValues(),
    nVotes = data.length,
    sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Dashboard"),
    latestVote = data[data.length - 1],
    sumVotes = latestVote.reduce((sum, currentValue) => (sum += currentValue), 0),
    [startRow, startColumn] = detectPos(sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues(), "Total"),
    poulNum = Math.max(...sheet.getRange(startRow, (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0].map(el => parseInt(el.replace(/[^0-9.]/gi, ""), 10))),
    poulSize = Math.max(...sheet.getRange((startRow + 1), (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)).getValues()[0]);
  
  Logger.log(`Number of groups : ${poulNum}`);
  Logger.log(`Size of a group : ${poulSize}`);
  Logger.log(`Sum of his votes : ${sumVotes}`);

  if (sumVotes == 10 * poulNum * poulSize) {
    // Adding the last vote to the dashboard
    sheet.getRange((startRow + nVotes), (startColumn + 1), 1, poulNum * poulSize).setValues([latestVote]).setBackground("white");
    // The last row contains the sum of all vote for given candidate
    let formulas = Array(poulNum * poulSize).map(() => `=SUM(R[-${nVotes}]C[0]:R[-1]C[0])`);
    sheet.getRange((startRow + nVotes + 1), (startColumn + 1), 1, poulNum * poulSize).setFormulasR1C1([formulas]).setBackground("#d9ead3");
  }
}

// Divide in each group to normalize
