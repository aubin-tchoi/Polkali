// Pushing and pulling drafts to share Drafts using a dedicated storage place in a shared Drive

// Name is used as a unique key to chose which draft should be pushed or pulled

const REPO_ID = "",
    TRASH_ID = "",
    IMAGES = {
        thumbsUp: "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/thumbsUp.png"
    },
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
        trash = DriveApp.getFolderById(TRASH_ID),
        myDraftsNames = drafts.map(d => d.getMessage().getSubject()),
        sharedDraftsNames = getSubFoldersNames(folder);

    // User input
    let pushedDraft = parseInt(ui.prompt("POUCHE-POULE", `Veuillez entrer le numéro du Draft que vous souhaitez push : ${listToQuery(myDraftsNames)}`, ui.ButtonSet.OK).getResponseText(), 10) || 0;
    if (pushedDraft == 0 || pushedDraft > myDraftsNames.length) {
        ui.alert("Entrée non valide", "Veuillez recommencer.", ui.ButtonSet.OK);
        return;
    }

    // Checking whether there is already a Draft in the repo to the same name or not
    if (sharedDraftsNames.includes(myDraftsNames[pushedDraft - 1])) {
        if (ui.alert("POUCHE-POULE", "Un Draft à ce nom existe déjà, souhaitez vous l'écraser ?", ui.ButtonSet.YES_NO) == ui.Button.YES) {
            trashFolder(folder, myDraftsNames[pushedDraft - 1], trash);
        } else {
            return;
        }
    }

    // displayLoadingScreen("Enregistrement du Draft...");

    // Creating a folder to store the Draft's data
    let draftFolder = folder.createFolder(myDraftsNames[pushedDraft - 1]);
    pushDraftIntoFolder(drafts[pushedDraft - 1], draftFolder);
}

// Pulling a draft from the repo
function pullDraft() {
    let drafts = GmailApp.getDrafts(),
        folder = DriveApp.getFolderById(REPO_ID),
        myDraftsNames = drafts.map(d => d.getMessage().getSubject()),
        sharedDraftsNames = getSubFoldersNames(folder);

    // User input
    let pulledDraft = parseInt(ui.prompt("POUCHE-POULE", `Veuillez entrer le numéro du Draft que vous souhaitez pull : ${listToQuery(sharedDraftsNames)}`, ui.ButtonSet.OK).getResponseText(), 10) || 0;
    if (pulledDraft == 0 || pulledDraft > sharedDraftsNames.length) {
        ui.alert("Entrée non valide", "Veuillez recommencer.", ui.ButtonSet.OK);
        return;
    }

    // Checking whether there is already a Draft in drafts to the same name or not
    if (myDraftsNames.includes(sharedDraftsNames[pulledDraft - 1])) {
        if (ui.alert("POUCHE-POULE", "Un Draft à ce nom existe déjà, souhaitez vous l'écraser ?", ui.ButtonSet.YES_NO) == ui.Button.YES) {
            drafts[pulledDraft - 1].deleteDraft();
        }
    }

    // displayLoadingScreen("Chargement du Draft...");

    // Creating a folder to store the Draft's data
    let dataFolder = getFolderByName(sharedDraftsNames[pulledDraft - 1]);
    pullDraftFromFolder(dataFolder);
}

// Moving a draft from the repo to the trash
function removeDraft() {
    let folder = DriveApp.getFolderById(REPO_ID),
        sharedDraftsNames = getSubFoldersNames(folder);

    // User input
    let trashedDraft = parseInt(ui.prompt("POUCHE-POULE", `Veuillez entrer le numéro du Draft que vous souhaitez mettre à la poubelle : ${listToQuery(sharedDraftsNames)}`, ui.ButtonSet.OK).getResponseText(), 10) || 0;
    if (trashedDraft == 0 || trashedDraft > sharedDraftsNames.length) {
        ui.alert("Entrée non valide", "Veuillez recommencer.", ui.ButtonSet.OK);
        return;
    }

    // displayLoadingScreen("Suppression du Draft...")

    // Trashing a folder based on its name
    if (myDraftsNames.includes(sharedDraftsNames[trashedDraft - 1])) {
        trashFolder(folder, sharedDraftsNames[trashedDraft - 1], trash);
        // Confirmation
        let operationSuccess = HtmlService
            .createHtmlOutput(`<span style='font-size: 16pt;'> <span style="font-family: 'roboto', sans-serif;">Le draft ${sharedDraftsNames[trashedDraft - 1]} a été placé dans la poubelle.</a>.<br/><br/> La bise</span></span><p style="text-align:center;"><img src="${IMAGES["thumsbUp"]}" alt="C'est la PEP qui régale !" width="130" height="131"></p>`);
        ui.showModalDialog(operationSuccess, "POUCHE-POULE");
    }
}

// Moving a draft from the trash to the repo
function recoverDraft() {
    let trash = DriveApp.getFolderById(TRASH_ID),
        trashedDraftsNames = getSubFoldersNames(trash);

    // User input
    let recoveredDraft = parseInt(ui.prompt("POUCHE-POULE", `Veuillez entrer le numéro du Draft que vous souhaitez récupérer : ${listToQuery(trashedDraftsNames)}`, ui.ButtonSet.OK).getResponseText(), 10) || 0;
    if (recoveredDraft == 0 || recoveredDraft > trashedDraftsNames.length) {
        ui.alert("Entrée non valide", "Veuillez recommencer.", ui.ButtonSet.OK);
        return;
    }

    // displayLoadingScreen("Récupération du Draft...")

    // Trashing a folder based on its name
    if (myDraftsNames.includes(trashedDraftsNames[recoveredDraft - 1])) {
        trashFolder(trash, trashedDraftsNames[recoveredDraft - 1], folder);
        // Confirmation
        let operationSuccess = HtmlService
            .createHtmlOutput(`<span style='font-size: 16pt;'> <span style="font-family: 'roboto', sans-serif;">Le draft ${trashedDraftsNames[recoveredDraft - 1]} a été récupéré.</a>.<br/><br/> La bise</span></span><p style="text-align:center;"><img src="${IMAGES["thumsbUp"]}" alt="C'est la PEP qui régale !" width="130" height="131"></p>`);
        ui.showModalDialog(operationSuccess, "POUCHE-POULE");
    }
}
