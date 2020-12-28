/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

// Generates a Google Forms representation of the bracket + links to our dashboard
function genForms(folderId, spreadsheet, filter) {
  // Pushing a file into folder and trashing it
  function moveFileToFolder(fileId, name, folder) {
    let file = DriveApp.getFileById(fileId),
      copy = file.makeCopy(name, folder);
    file.setTrashed(true);
    return copy.getId();
  }
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
    .setDescription(`Bienvenue dans le bracket, pour chaque poule vous disposez de 10 points à répartir sur l'ensemble des candidats. Soyez avisés.`)
    .setConfirmationMessage(`Merci pour votre participation, la direction du Bracket vous recontactera sous peu pour annoncer les résultats.`)
    .setRequireLogin(false)
    .setShowLinkToRespondAgain(false),
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
      if (filter(poulSize)) {
        let image = files.next().getBlob();
        fillSection(forms, image, poulNum, ++poulSize);
      }
    }
  }

  // Setting my dashboard as a destination
  forms = FormApp.openById(moveFileToFolder(forms.getId(), "Bracket", folder));
  forms.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  // Renaming the sheet
  let initRegex = RegExp("Réponse au formulaire"),
    targetRegex = RegExp("Round");
  spreadsheet.getSheets().filter(sheetName => initRegex.test(sheetName)).shift()
    .setName(`Round ${spreadsheet.getSheets().filter(sheetName => targetRegex.test(sheetName)).length + 1}`);

  return [forms.getEditUrl(), forms.getPublishedUrl(), poulNum, poulSize];
}

// Sends an email containing the link to every email adress found in the sheet
function mailLink(spreadsheet, link) {
  // Returns a boolean function meant to be used in a filter function
  function subjectFilter_(template) {
    return (element => element.getMessage().getSubject() === template);
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
        heads.forEach(function (key) {
          let regexp = new RegExp(`{${key}}|{${key.toLowerCase()}}|{${key.replace(/[éêè]/gi, "e")}}`, "gmi");
          msgHtml = msgHtml.replace(regexp, row[key]);
        });

        // Adding the form's link to the mail
        msgHtml = msgHtml.replace(/{link}/gmi, `<a href = ${link}>Bracket</a>`);

        // Sending the mail
        let msgPlain = msgHtml.replace(/\<br\/\>/gmi, '\n').replace(/(<([^>]+)>)/gmi, "");
        Logger.log(msgHtml);
        Logger.log(msgPlain);
        GmailApp.sendEmail(row["Adresse mail"], "Participez au bracket !", msgPlain, {
          htmlBody: msgHtml,
          inlineImages: inlineImages
        });
        return [new Date()];
      } catch (e) {
        Logger.log(`Issue with row ${row} : ${e}`);
        return [(e.message == "Cannot read property 'getMessage' of undefined") ? "Pas de modèle à ce nom" : e.message];
      }
    } else {
      return [row["Date d'envoi"]];
    }
  }

  const sheet = spreadsheet.getSheetByName("Dashboard"),
    templateName = sheet.getRange(detectPos(sheet, "Modèle Gmail")[0], (detectPos(sheet, "Modèle Gmail")[1] + 1), 1, 1).getValues()[0][0];
  let [startRow, startColumn] = detectPos(sheet, "Mailing"),
    data = sheet.getRange((startRow + 1), startColumn, (sheet.getLastRow() - startRow), 4).getValues(),
    heads = data.shift(),
    output = [],
    nSent = data.filter(row => row[0] == "" && row[1] != "").length;
  data = data.map(row => heads.reduce((o, k, i) => (o[k] = (row[i] != "") ? row[i] : o[k] || '', o), {})).filter(row => row["Adresse mail"] != "");

  Logger.log(`Heads : ${heads}`);
  Logger.log(`Template name : ${templateName}`);

  // Sending a mail to each email adress
  data.forEach(function (row) {
    output.push(sendMail(templateName, link, row));
  });
  sheet.getRange((startRow + 2), startColumn, output.length, 1).setValues(output);
  return nSent;
}

// Loading screen
function displayLoadingScreen(msg) {
  let htmlLoading = HtmlService
    .createHtmlOutput(`<img src="https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/Kkooljem.gif" alt="Loading" width="442" height="249">`)
    .setWidth(450)
    .setHeight(325);
  SpreadsheetApp.getUi().showModelessDialog(htmlLoading, msg);
}
// Detects the position of an element in a set of data
function detectPos(sheet, element) {
  let data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues(),
    row = (data.indexOf(data.filter(r => r.includes(element))[0]) + 1),
    column = (data[row - 1].indexOf(element) + 1);
  return [row, column];
}
