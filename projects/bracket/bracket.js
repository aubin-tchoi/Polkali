/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bracket')
  .addItem("Génération du Forms", "stinson")
  .addToUi();
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

  // Detects the position of an element in a set of data
  function detectPos(data, element) {
    let row = (data.indexOf(data.filter(r => r.includes(element))[0]) + 1),
      column = (data[row - 1].indexOf(element) + 1);
    return [row, column];
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
      templateName = sheet.getRange(1, 1, 1, 1).getValues()[0][0];
    let data = sheet.getRange(3, 2, (sheet.getLastRow() - 2), 4).getValues(),
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

    sheet.getRange(4, 2, output.length, 1).setValues(output);
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
    function fillSection(forms, image) {
      let imageItem = forms.addImageItem()
        .setImage(image);
      let mark = forms.addTextItem()
        .setTitle('Votre note')
        .setRequired(true);
    }    

    // First questions
    const folder = DriveApp.getFolderById(folderId);
    let forms = FormApp.create("Bracket")
      .setDescription("Bienvenue dans le bracket, pour chaque poule vous disposez de 10 points à répartir sur l'ensemble des candidats. Soyez avisés."),
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
      poulSize++;

        // Looping on each file (adding the image + a question for each file)
      while (files.hasNext()) { 
        let image = files.next().getBlob();
        fillSection(forms, image);
      }
    }

    // Setting my dashboard as a destination
    forms = FormApp.openById(moveFileToFolder(forms.getId(), "Bracket", folder));
    forms.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    return [forms.getEditUrl(), forms.getPublishedUrl(), poulNum, poulSize];
  }

  // Loading screen
  displayLoadingScreen("Génération du Forms ..")

  const folderId = "1nHfPZR10ZCFx-Nro546WjRwAmih2_bfq",
    spreadsheetId = "1C68xciVCymtPCM-DcHM_ZMx7o7oZ0ydi7vs5bT8I0Ns",
    spreadsheet = SpreadsheetApp.openById(spreadsheetId),
    [editLink, publishedLink, poulNum, poulSize] = genForms(folderId, spreadsheet),
    [startRow, startColumn] = formatSheet(spreadsheet, poulNum, poulSize),
    nSent = mailLink(spreadsheet, publishedLink);

  const updateSheet = () => {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
      data = dataSheet.getRange(2, 2, (dataSheet.getLastRow() - 1), (dataSheet.getLastColumn() - 1)).getValues(),
      heads = data.shift(),
      nVotes = data.length,
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard"),
      sumVotes = data[data.length - 1].reduce((sum, currentValue) => (sum += currentValue), 0);

    Logger.log(`Number of groups : ${poulNum}`);
    Logger.log(`Size of a group : ${poulSize}`);
    Logger.log(`Sum of his votes : ${sumVotes}`);

    if (sumVotes == 10 * poulNum * nVotes) {
      // Adding the last vote to the dashboard
      sheet.getRange((startRow + nVotes), (startColumn + 1), 1, poulNum * poulSize).setValues([data[data.length - 1]]).setBackground("white");
      // The last row contains the sum of all vote for given candidate
      let formulas = Array(poulNum * poulSize).map(() => `=SUM(R[-${nVotes}]C[0]:R[-1]C[0])`);
      sheet.getRange((startRow + nVotes + 1), (startColumn + 1), 1, poulNum * poulSize).setFormulasR1C1([formulas]).setBackground("#d9ead3");
    }
  }

  ScriptApp.newTrigger("updateSheet")
  .forSpreadsheet(spreadsheet)
  .onFormSubmit()
  .create();

  let htmlOutput = HtmlService.createHtmlOutput(`<span style="font-family: 'trebuchet ms', sans-serif;">Voci le lien éditeur : <a href = "${editLink}">éditer le Form</a>.<br/>
                                                <br/> Voici le lien lecteur : <a href = "${publishedLink}">répondre au Form</a>.<br/>
                                                <br/> Le lien lecteur a été envoyé à ${nSent} personne${nSent >= 2 ? "s" : ""}.</span>`);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, "Succès de l'opération")
  // Image size, description, confirmation message on Forms, date d'envoi du mail
}
