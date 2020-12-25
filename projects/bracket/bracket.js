/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bracket')
  .addItem("Génération du Forms", "stinson")
  .addToUi();
}

function stinson() {
  // Adds a line to the dashboard
  function formatSheet(spreadsheetId, poulNum, poulSize) {
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Dashboard"),
      values = sheet.getRange(1, 1, shet.getLastRow(), sheet.getLastColumn()).getValues();
    let startRow = values.indexOf(values.filter(row => row.includes("Total"))) + 1,
      startColumn = values[startRow - 1].indexOf(values[startRow - 1].filter(el => el == "Total")) + 1;
    
    Logger.log(`Row offset : ${startRow}`);
    Logger.log(`Column offset : ${startColumn}`);

    // First line
    for (let poule = 1; poule <= poulNum; poule++) {
      sheet.getRange(startRow, (startColumn + (poule - 1) * poulSize), 1, poulSize)
      .merge()
      .setValues([[`Poule ${poule}`]])
      .setBackground("#ffe599");
      sheet.getRange((startRow + 1), (startColumn + (poule - 1) * poulSize, 1, poulSize)).setValues(Array.from({length: poulSize}, (_, i) => i + 1));
    }
    return startRow, startColumn;
  }

  // Sends an email containing the link to every email adress found in the sheet
  function mailLink(spreadsheetId, link) {
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
      return msgHtml, inlineImages;
    }

    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Dashboard"),
      templateName = sheet.getRange(1, 1, 1, 1).getValues()[0][0]
      data = sheet.getRange(3, 2, sheet.getLastRow(), 3).getValues(),
      heads = data.shift();
    data = data.map(row => heads.reduce((o, k, i) => (o[k] = (r[i] != "") ? row[i] : o[k] || '', o), {}));

    Logger.log(`Data : ${data}`);

    data.forEach(function(row) {
      let msg = GmailApp.getDrafts().filter(subjectFilter_(templateName))[0].getMessage(),
      msgHtml = msg.getBody(),
      inlineImages = {};
        
    // Inline images
    msgHtml, inlineImages = getInlineImages(msgHtml);

    // Customizing the model with data taken from row
    heads.forEach(function(key) {
      let regexp = new RegExp(`{(${key}|${key.toLowerCase()}|${key.replace(/[éêè]/gi, "e")}|${key.toLowerCase().replace(/[éêè]/gi, "e")})}`, "gi");
      msgHtml.replace(regexp, row[key]);});
    
    msgHtml.replace("{link}", link);

    // Sending the mail
    let msgPlain = msgHtml.replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/gi, "");
    GmailApp.sendEmail(row["Adresse mail"], subject, msgPlain, {htmlBody:msgHtml, inlineImages:inlineImages});
    });  
  }

  // Generates a Google Forms representation of the bracket
  function genForms(folderId, spreadsheetId) {
    // Adding an image+mark to the forms
    function fillSection(forms, image) {
      let imageItem = forms.addImageItem()
        .setImage(image);
      let mark = forms.addTextItem()
        .setTitle('Votre note')
        .setRequired(true);
    }

    // Pushing a file into folder and trashing it
    function moveFileToFolder(fileId, name, folder) {
      let file = DriveApp.getFileById(fileId),
        copy = file.makeCopy(name, folder);
      file.setTrashed(true);
      return copy;
    }

    // First questions
    const folder = DriveApp.getFolderById(folderId),
      forms = FormApp.create("Bracket")
      .setDescription("Bienvenue dans le bracket, pour chaque poule vous disposez de 10 points à répartir sur l'ensemble des candidats. Soyez avisés."),
      firstName = forms.addTextItem().setTitle("Quel est votre prénom ?").setRequired(true),
      lastName = forms.addTextItem().setTitle("Quel est votre nom ?").setRequired(true);

    let folders = folder.getFolders(),
      poulNum = 1,
      poulSize = 0;
    
    // Looping on each folder (1 section for each folder, folder <=> poule)
    while (folders.hasNext()) {
      let subFolder = folders.next(),
        files = subFolder.getFiles(),
        section = forms.addPageBreakItem()
        .setTitle(`Poule ${poulNum++}`);
      poulSize++;

        // Looping on each file (adding the image + a question for each file)
      while (files.hasNext()) { 
        let image = files.next().getBlob();
        fillSection(forms, image);
      }
    }

    // Setting my dashboard as a destination
    let spreadsheet = SpreadsheetApp.getFileById(spreadsheetId);
    forms = moveFileToFolder(forms.getId(), "Bracket", folder);
    spreadsheet = moveFileToFolder(spreadsheet.getId(), "Bracket", folder);
    forms.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    return forms.getPublishedUrl(), poulNum, poulSize;
  }

  const folderId = "1nHfPZR10ZCFx-Nro546WjRwAmih2_bfq",
    spreadsheetId = "1C68xciVCymtPCM-DcHM_ZMx7o7oZ0ydi7vs5bT8I0Ns",
    [link, poulNum, poulSize] = genForms(folderId, spreadsheetId),
    [startRow, startColumn] = formatSheet(spreadsheetId, poulNum, poulSize);
  mailLink(spreadsheetId, link);

  const updateSheet = () => {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Réponses au formulaire 1"),
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
  .forSpreadsheet(SpreadsheetApp.openById(spreadsheetId))
  .onFormSubmit()
  .create();
  // Image size, description, confirmation message
}
