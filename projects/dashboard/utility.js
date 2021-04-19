/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Functionnalities aiming at interfacing the generation of KPI with other tools from the G-Suite

// Send a mail to designated adress
function sendMail(adress, htmlOutput, subject, attachments) {
    let msgHtml = htmlOutput.getContent(),
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

// Generating a Slides file using the charts listed in the array named charts
// NB: many values are hardcoded here, do not put them at the top of the main file, these are not parameters and won't come to change
function generateSlides(template, chartImages, folderId) {
    displayLoadingScreen("Génération des slides..");
    // Today's date
    let today = new Date();
    today = `Point KPI du ${(today.getDate() < 9) ? "0" : ""}${today.getDate()}/${(today.getMonth() < 9) ? "0" : ""}${today.getMonth() + 1}/${today.getFullYear()}`;
    let presentation = SlidesApp.openById(DriveApp.getFileById(template).makeCopy(today, DriveApp.getFolderById(folderId)).getId());
    
    // First slide
    let firstSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.SECTION_HEADER);
    firstSlide.getShapes()[0].getText().setText(today);
    firstSlide.getShapes()[1].getText().setText("Aubin Tchoï, Directeur Qualité 022");
    presentation.getSlides()[0].remove();

    chartImages.forEach(img => {
        // New slide
        let slide = presentation.appendSlide();
        // White box to hide PEP's logo in the bottom right corner
        let box = slide.insertShape(SlidesApp.ShapeType.RECTANGLE).setLeft(650).setTop(100).setWidth(300).setHeight(430);
        box.getBorder().getLineFill().setSolidFill('#FFFFFF');
        box.getFill().setSolidFill('#FFFFFF');
        // Inserting the image
        slide.insertImage(img).alignOnPage(SlidesApp.AlignmentPosition.CENTER).setTop(375 - DIMS["height"] / 2);
        // Writing a title
        let text = slide.insertTextBox(img.getName()).setTop(65).setWidth(480).alignOnPage(SlidesApp.AlignmentPosition.HORIZONTAL_CENTER).getText();
        text.getTextStyle().setFontSize(28).setForegroundColor(COLORS["burgundy"]);
        text.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });
}
