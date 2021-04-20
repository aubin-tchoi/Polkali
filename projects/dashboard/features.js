/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Features aiming at interfacing the generation of KPI with other tools from the G-Suite

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
        Object.values(imageBlobs).forEach(arr => {
            arr.forEach(f => {
                folder.createFile(f);
            });
        });
        let confirm = HTML_CONTENT.saveConfirm(folder.getUrl());
        ui.showModelessDialog(confirm, "KPIs enregistrés !");
    } catch (e) {
        Logger.log(`Erreur lors de l'enregistrement des images sur le Drive : ${e}.`);
        ui.alert("Erreur lors de l'enregistrement des images sur le Drive.");
    }
}

// Generating a Slides file using the charts listed in the array named charts
// NB1: many values are hardcoded here, do not put them in the parameters file, these are not parameters and won't come to change
// NB2: many suppositions are made on the template (and on its predefined layout), this won't work on any template
function generateSlides(template, chartImages, folderId) {
    displayLoadingScreen("Génération des slides..");

    // Today's date
    let today = new Date();
    today = `Point KPI du ${(today.getDate() < 9) ? "0" : ""}${today.getDate()}/${(today.getMonth() < 9) ? "0" : ""}${today.getMonth() + 1}/${today.getFullYear()}`;
    let presentation = SlidesApp.openById(DriveApp.getFileById(template).makeCopy(today, DriveApp.getFolderById(folderId)).getId()),
        chartlayout = presentation.getLayouts()[8],
        sectionLayout = presentation.getLayouts()[6];

    // First slide
    let currentSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.SECTION_HEADER);
    currentSlide.getShapes()[0].getText().setText(today);
    currentSlide.getShapes()[1].getText().setText("Aubin Tchoï, Directeur Qualité 022");
    presentation.getSlides()[0].remove();

    let box = currentSlide.insertShape(SlidesApp.ShapeType.RECTANGLE).setLeft(650).setTop(140).setWidth(300).setHeight(380),
        text;
    box.getBorder().getLineFill().setSolidFill('#FFFFFF');
    box.getFill().setSolidFill('#FFFFFF');

    Object.keys(chartImages).forEach(key => {
        // New slide to present the section
        if (chartImages[key].length > 0) {
            currentSlide = presentation.appendSlide(sectionLayout);
            currentSlide.getShapes()[1].getText().setText(today);
            currentSlide.getShapes()[2].getText().setText("Aubin Tchoï, Directeur Qualité 022");
            currentSlide.getShapes()[0].getText().setText(TITLES[key]);
        }
        chartImages[key].forEach(img => {
            // New slide
            currentSlide = presentation.appendSlide(chartlayout);
            // White box to hide PEP's logo in the bottom right corner
            currentSlide.insertShape(box);
            // Inserting the image
            currentSlide.insertImage(img).alignOnPage(SlidesApp.AlignmentPosition.CENTER).setTop(375 - DIMS.height / 2);
            // Writing a title
            text = currentSlide.getShapes()[0].getText().setText(img.getName());
        });
    });
    box.remove();
}
