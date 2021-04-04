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
function saveOnDrive(imageBlobs, folderId = DRIVE["folderId"]) {
    try {
        displayLoadingScreen("Enregistrement des images sur le Drive..");
        // Folder will be dated with current date
        let today = new Date();
        today = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
        let folder = DriveApp.getFolderById(folderId).createFolder(`KPI ${today}`);
        imageBlobs.forEach(function (f) {
            folder.createFile(f);
        });
        ui.alert("Enregistrement des images sur le Drive", `Les images ont été enregistrées à ladresse suivante : ${folder.getUrl()}`, ui.ButtonSet.OK);
    } catch (e) {
        Logger.log(`Erreur lors de l'enregistrement des images sur le Drive : ${e}.`);
        ui.alert("Erreur lors de l'enregistrement des images sur le Drive.");
    }

}