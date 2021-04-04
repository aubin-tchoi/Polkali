/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Generating a Slides file using the charts listed in the array named charts
function generateSlides(template, charts, folderId) {
    let today = new Date();
    today = `Point KPI du ${(today.getDate() < 9) ? "0" : ""}${today.getDate()}/${(today.getMonth() < 9) ? "0" : ""}${today.getMonth() + 1}/${today.getFullYear()}`;
    let presentation = SlidesApp.openById(DriveApp.getFileById(template).makeCopy(today, DriveApp.getFolderById(folderId)).getId());

    charts.forEach(c => {presentation.appendSlide().insertSheetsChart(c);});
    // Htmoutput to give the link
}