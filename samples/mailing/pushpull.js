// Pushing and pulling drafts to share Drafts using a dedicated storage place in a shared Drive

// Name is used as a unique key to chose which draft should be pushed or pulled

const REPO_ID = {},
    TRASH_ID = {},
    ui = SpreadsheetApp.getUi();

function onOpen() {
    ui.createMenu('POUCHE-POULE')
        .addItem("Push", "pushDraft")
        .addItem("Pull", "pullDraft")
        .addSeparator()
        .addItem("Remove", "removeDraft")
        .addItem("Retrieve", "recoverDraft")
        .addToUi();
}

// Pushing a draft onto the repo
function pushDraft() {
    let drafts = GmailApp.getDrafts(),
        folder = DriveApp.getFolderById(REPO_ID),
        subFolders = folder.getFolders(),
        myDraftsNames = drafts.map(d => d.getMessage().getSubject()),
        sharedDraftsNames = [];

    // sharedDraftsNames contains the names of all Drafts stored in the repo
    while (subFolders.hasNext()) {
        let subFolder = folders.next();
        sharedDraftsNames.push(subFolder.getName());
    }

    // User input
    let pushedDraft = parseInt(ui.prompt("POUCHE-POULE", `Veuillez entrer le numéro du Draft que vous souhaitez push : ${listToQuery(drafts.map(d => d.getMessage().getSubject()))}`), 10) || 0;
    if (pushedDraft == 0 || pushedDraft > myDraftsNames.length) {
        ui.alert("Entrée non valide", "Veuillez recommencer.", ui.ButtonSet.OK);
        return;
    }

    // There is already a Draft in the repo to the same name
    if (sharedDraftsNames.includes(myDraftsNames[pushedDraft - 1])) {
        ui.alert("POUCHE-POULE", "Un Draft à ce nom existe déjà, souhaitez vous l'écraser ?", ui.ButtonSet.YES_NO);
        subFolders = folder.getFolders();
        while (subFolders.hasNext()) {
            let subFolder = folders.next();
            if (subFolder.getName() == myDraftsNames[pushedDraft - 1]) {
                subFolder.setTrashed(true);
            }
        }
    }

    // Creating a folder to store the Draft's data
    let draftFolder = folder.createFolder(myDraftsNames[pushedDraft - 1]);
    pouche(drafts[pushedDraft - 1], draftFolder);
}

// Pulling a draft from the repo
function pullDraft() {

}

// Moving a draft from the repo to the trash
function removeDraft() {

}

// Moving a draft from the trash to the repo
function recoverDraft() {

}
