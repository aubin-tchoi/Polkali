/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Generating a Slides file using the charts listed in the array named charts
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
        slide.insertImage(img).alignOnPage(SlidesApp.AlignmentPosition.CENTER).setTop(650 - DIMS["height"]);
        // Writing a title
        let text = slide.insertTextBox(img.getName()).setTop(65).setWidth(300).alignOnPage(SlidesApp.AlignmentPosition.HORIZONTAL_CENTER).getText();
        text.getTextStyle().setFontSize(28).setForegroundColor(COLORS["burgundy"]);
        text.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });
}