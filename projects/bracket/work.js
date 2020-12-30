/* For any further question, please contact Aubin Tchoï */

// Adds a line to the dashboard
function formatSheet(sheet, groupNumber, groupSize) {
    const [startRow, startColumn] = detectColor(sheet, MARKERS["groups"]);
    sheet.getRange((startRow + 1), startColumn).setValue("Numéro").setBackground(MARKERS["currentRound"]);

    // First line
    for (let group = 0; group < groupNumber; group++) {
        let firstLine = [Array.from(Array(groupSize).keys()).map(idx => (idx == 0) ? `Poule ${group + 1}` : "")],
            secondLine = [Array.from(Array(groupSize).keys()).map(x => x + 1)];
        sheet.getRange(startRow, (startColumn + 1 + group * groupSize), 1, groupSize)
            .setValues(firstLine)
            .setBackground(COLORS["groups"])
            .setHorizontalAlignment("center")
            .merge();
        sheet.getRange((startRow + 1), (startColumn + 1 + group * groupSize), 1, groupSize)
            .setValues(secondLine)
            .setBackground(COLORS["contestants"])
            .setHorizontalAlignment("center");
        sheet.setColumnWidths((startColumn + 1), groupNumber * groupSize, 33);
    }
}

// Reorganizes the folder to remove losers
function removeLosers(sheet, folderId, trashId) {
    // Sheets and folders
    const folder = DriveApp.getFolderById(folderId),
        trash = DriveApp.getFolderById(trashId),

        startColumn = detectColor(sheet, MARKERS["currentRound"])[1],

        range = sheet.getRange(sheet.getLastRow(), (startColumn + 1), 1, (sheet.getLastColumn() - startColumn)),
        losers = range.getBackgrounds()[0].map(background => background == COLORS["loser"]);

    Logger.log(`Losers' indexes : ${losers}`);

    // Discarding the losers in each subfolder
    let folders = folder.getFolders(),
        fileIdx = 0;
    while (folders.hasNext()) {
        let subFolder = folders.next(),
            files = subFolder.getFiles();
        Logger.log(`Removing losers in file : ${subFolder.getName()}.`);
        while (files.hasNext()) {
            let file = files.next();
            if (losers[fileIdx++]) {
                Logger.log(`Removing loser number : ${fileIdx - 1}.`);
                file.moveTo(trash);
            }
        }
    }

    // Merging each pair of subfolders together
    folders = folder.getFolders();
    while (folders.hasNext()) {
        let firstSubFolder = folders.next(),
            secondSubFolder = folders.next(),
            files = secondSubFolder.getFiles();
        while (files.hasNext()) {
            files.next().moveTo(firstSubFolder);
        }
        secondSubFolder.setTrashed(true);
    }
}

// New becomes old
function updateMarkers(sheet) {
    let oldPos = detectColor(sheet, MARKERS["currentRound"]),
        newPos = detectColor(sheet, MARKERS["nextRound"]);
    sheet.getRange(oldPos[0], oldPos[1]).setBackground(MARKERS["oldRound"]);
    sheet.getRange(newPos[0], newPos[1]).setBackground(MARKERS["currentRound"]);
}

// Generates a Google Forms representation of the bracket, links it to a spreadsheet and installs and onFormSubmit trigger
function generateForms(folderId, roundNumber) {
    // Adding a contestant to the Form
    function addContestant(forms, image, group, contestant) {
        forms.addImageItem()
            .setImage(image)
            .setTitle(`${contestant}${contestant >= 2 ? "ème" : "er"} candidat de la poule ${group}`)
            .setWidth(FORMS_PARAMETERS["imageSize"]);
        forms.addTextItem()
            .setTitle('Votre note')
            .setRequired(true);
    }

    displayLoadingScreen("Génération du Forms ..")

    const folder = DriveApp.getFolderById(folderId),
        roundName = `Round ${roundNumber}`,
        destination = SpreadsheetApp.create(roundName);

    // First questions & configuration
    let forms = FormApp.create(roundName)
        .setDescription(FORMS_PARAMETERS["description"])
        .setConfirmationMessage(FORMS_PARAMETERS["confirmationMessage"])
        .setRequireLogin(false)
        .setShowLinkToRespondAgain(false);
    forms.addTextItem().setTitle("Quel est votre prénom ?").setRequired(true);
    forms.addTextItem().setTitle("Quel est votre nom ?").setRequired(true);

    // Looping on each folder (1 section for each folder)
    let folders = folder.getFolders(),
        groupIdx = 0,
        groupSize = 0;

    while (folders.hasNext()) {
        let subFolder = folders.next(),
            files = subFolder.getFiles();
        forms.addPageBreakItem().setTitle(`Poule ${++groupIdx}`);
        groupSize = 0;
        // Looping on each file (adding the image + a question for each file)
        while (files.hasNext()) {
            let image = files.next().getBlob();
            addContestant(forms, image, groupIdx, ++groupSize);
        }
    }

    // Moving both the forms and its destination to folder
    DriveApp.getFileById(forms.getId()).moveTo(folder);
    DriveApp.getFileById(destination.getId()).moveTo(folder);

    // Setting destination as forms' destination
    forms.setDestination(FormApp.DestinationType.SPREADSHEET, destination.getId());

    // Setting a trigger on the destination sheet
    ScriptApp.newTrigger("updateSheet")
        .forSpreadsheet(destination)
        .onFormSubmit()
        .create();

    return [forms.getEditUrl(), forms.getPublishedUrl(), groupIdx, groupSize];
}

// Sends an email containing the link to every email adress found in the sheet
function mailLink(sheet, link) {
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

    // Sends a mail based on a template using the data stored in row
    function sendMail(data, heads, templateName, link) {
        try {
            let msg = GmailApp.getDrafts().filter(subjectFilter_(templateName))[0].getMessage();

            // Inline images
            let [msgHtml, inlineImages] = getInlineImages(msg);

            // Customizing the model with data read from data
            heads.forEach(function (key) {
                let regexp = new RegExp(`{${key}}|{${key.toLowerCase()}}|{${key.replace(/[éêè]/gi, "e")}}`, "gmi");
                msgHtml = msgHtml.replace(regexp, data[key]);
            });

            // Adding the form's link to the mail
            msgHtml = msgHtml.replace(/{link}/gmi, `<a href = ${link}>Bracket</a>`);

            Logger.log(`Contents of the mail sent : ${msgHtml}`);

            // Sending the mail
            let msgPlain = msgHtml.replace(/\<br\/\>/gmi, '\n').replace(/(<([^>]+)>)/gmi, "");
            GmailApp.sendEmail(data["Adresse mail"], "Participez au bracket !", msgPlain, {
                htmlBody: msgHtml,
                inlineImages: inlineImages
            });
            return [new Date()];
        } catch (e) {
            Logger.log(`Issue with row ${data} : ${e}`);
            return [(e.message == "Cannot read property 'getMessage' of undefined") ? "Pas de modèle à ce nom" : e.message];
        }
    }

    const ui = SpreadsheetApp.getUi();

    if (ui.alert("Bracket", "Souhaitez-vous envoyer le lien du Google Form par mail aux participants ?", ui.ButtonSet.YES_NO) == ui.Button.YES) {
        const templateName = sheet.getRange(detectColor(sheet, MARKERS["template"])[0], (detectColor(sheet, MARKERS["template"])[1] + 1)).getValue(),
            [startRow, startColumn] = detectColor(sheet, MARKERS["mail"]);

        let data = sheet.getRange((startRow + 1), startColumn, (sheet.getLastRow() - startRow), 4).getValues(),
            heads = data.shift(),
            output = [];
        data = data.map(row => heads.reduce((o, k, i) => (o[k] = (row[i] != "") ? row[i] : o[k] || '', o), {}));

        // Sending a mail to each email adress
        data.forEach(function (row) {
            output.push(sendMail(row, heads, templateName, link));
        });

        // Writing on column "Date d'envoi"
        sheet.getRange((startRow + 2), startColumn, output.length, 1).setValues(output);

        return output.length;
    }
    return 0;
}
