/* Si vous avez des questions à propos de ce script, contactez Aubin Tchoï (Directeur Qualité 022)   */

function stinson() {
  // Generates a Google Forms representation of the bracket
  function genForms(folderId) {
    // Filling a section with image+mark
    function fillSection(forms, image) {
      let imageItem = forms.addImageItem()
        .setImage(image);
      let mark = forms.addTextItem()
        .setTitle('Votre note')
        .setRequired(true);
    }

    // Pushing a file in to folder and trashing it
    function moveFileToFolder(fileId, name, folder) {
      let file = DriveApp.getFileById(fileId),
        copy = file.makeCopy(name, folder);
      file.setTrashed(true);
    }

    const folder = DriveApp.getFolderById(folderId),
      forms = FormApp.create("Bracket")
      .setDescription("Bienvenue dans le bracket, pour chaque poule vous disposez de 10 points à répartir sur l'ensemble des candidats. Soyez avisés."),
      firstName = forms.addTextItem().setTitle("Quel est votre prénom ?").setRequired(true),
      lastName = forms.addTextItem().setTitle("Quel est votre nom ?").setRequired(true);

    let folders = folder.getFolders(),
      poulNum = 1;
    
    // Looping on each folder (1 section for each folder, folder <=> poule)
    while (folders.hasNext()) {
      let subFolder = folders.next(),
        files = subFolder.getFiles(),
        section = forms.addPageBreakItem()
        .setTitle(`Poule ${poulNum++}`);

        // Looping on each file (adding the image + a question for each file)
      while (files.hasNext()) { 
        let image = files.next().getBlob();
        fillSection(forms, image);
      }
    }
    moveFileToFolder(forms.getId(), "Bracket", folder);
    let sheet = SpreadsheetApp.create("Bracket");
    moveFileToFolder(sheet, "Bracket", folder);

    formCopy.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  }

  const folderId = "1nHfPZR10ZCFx-Nro546WjRwAmih2_bfq";
  genForms(folderId);
  // mailLink();
}
