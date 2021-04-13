/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Functionnalities aiming at interfacing the generation of KPI with other tools from the G-Suite

// Send a mail to designated adress
function sendMail(htmlOutput, subject, attachments) {
    let adress = ui.prompt("Envoi des diagrammes par mail", "Entrez l'adresse mail de destination :", ui.ButtonSet.OK).getResponseText(),
        msgHtml = htmlOutput.getContent(),
        msgPlain = htmlOutput.getContent().replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/ig, "");
    GmailApp.sendEmail(adress, subject, msgPlain, {
        htmlBody: msgHtml,
        attachments: attachments
    });
    ui.alert("Envoi des diagrammes par mail", `Les diagrammes ont été envoyés par mail à : ${adress}.`, ui.ButtonSet.OK);
}


// Save data in designated Drive folder
function saveOnDrive(imageBlobs, folderId) {
    try {
        displayLoadingScreen("Enregistrement des images sur le Drive..");
        // Folder will be dated with current date
        let today = new Date();
        today = `KPI ${(today.getDate() < 9) ? "0" : ""}${today.getDate()}/${(today.getMonth() < 9) ? "0" : ""}${today.getMonth() + 1}/${today.getFullYear()}`;
        let folder = DriveApp.getFolderById(folderId).createFolder(today);
        imageBlobs.forEach(function (f) {
            folder.createFile(f);
        });
        let confirm = HtmlService
            .createHtmlOutput(HTML_CONTENT["saveConfirm"](folder.getUrl()))
            .setHeight(235)
            .setWidth(600);
        ui.showModelessDialog(confirm, "KPIs enregistrés !");
    } catch (e) {
        Logger.log(`Erreur lors de l'enregistrement des images sur le Drive : ${e}.`);
        ui.alert("Erreur lors de l'enregistrement des images sur le Drive.");
    }
}

// Converting a chart object into an image
function convertChart(chart, title, htmlOutput, attachments) {
    // Adding the chart to the HtmlOutput
    let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
        imageUrl = "data:image/png;base64," + encodeURI(imageData);
    htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

    // Adding the chart to the attachments
    let imageDatamail = chart.getAs('image/png').getBytes(),
        imgblob = Utilities.newBlob(imageDatamail, "image/png", title);
    attachments.push(imgblob);
}
