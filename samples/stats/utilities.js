// Query and send mail
function sendMail(htmlOutput, adress, subject, attachments) {
    let query = ui.alert("Envoi des diagrammes par mail", "Souhaitez vous recevoir les diagrammes par mail ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
        let msgHtml = htmlOutput.getContent(),
            msgPlain = htmlOutput.getContent().replace(/\<br\/\>/gi, '\n').replace(/(<([^>]+)>)/ig, "");
        GmailApp.sendEmail(adress, subject, msgPlain, {
            htmlBody: msgHtml,
            attachments: attachments
        });
        ui.alert("Envoi des diagrammes par mail", `Les diagrammes ont été envoyés par mail à l'adresse ${adress}`, ui.ButtonSet.OK);
    }
}

// Query and save data on the Drive
function saveOnDrive(folderId = DRIVE_FOLDERS["stats"], imageBlobs) {
    let query = ui.alert("Enregistrement des images sur le Drive", "Souhaitez vous enregistrer les images sur le Drive ?", ui.ButtonSet.YES_NO);
    if (query == ui.Button.YES) {
        displayLoadingScreen("Enregistrement des images sur le Drive..");
        let today = new Date();
        today = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
        let folder = DriveApp.getFolderById(folderId).createFolder(`Statistiques groupées ${today}`);
        imageBlobs.forEach(function (f) {
            folder.createFile(f);
        });
        ui.alert("Enregistrement des images sur le Drive", `Les images ont été enregistrées à l'adresse suivante : ${folder.getUrl()}`, ui.ButtonSet.OK);
    }
}

// Writing data on a new sheet
function rewrite(data, heads, ss) {
    let newdata = data.map(function (row) {
        newrow = [];
        heads.forEach(function (key) {
            newrow.push(row[key]);
        });
        return newrow;
    });

    function write(newsheet, newdata, newheads) {
        newsheet.getRange(1, 1, 1, newheads.length)
            .setBackgroundRGB(255, 204, 229)
            .setValues([newheads]);
        newsheet.getRange(2, 1, newdata.length, newdata[0].length).setValues(newdata);
        newsheet.setColumnWidths(1, newsheet.getLastColumn(), 230);
    }

    try {
        let newsheet = ss.insertSheet(NAME_SHEET["destination"]);
        write(newsheet, newdata, heads);
    } catch (e) {
        let query = ui.alert("Avertissement", `L'onglet ${NAME_SHEET["destination"]} existe déjà, son contenu sera supprimé. \n Souhaitez vous continuer ?`, ui.ButtonSet.YES_NO);
        if (query == ui.Button.YES) {
            displayLoadingScreen("Écriture des données ...");
            Logger.log(`Couldn't create sheet : ${e}`);
            ss.deleteSheet(ss.getSheetByName(NAME_SHEET["destination"]));
            let newsheet = ss.insertSheet(NAME_SHEET["destination"]);
            write(newsheet, newdata, heads);
        }
    }
}

// Filters data with requested key
function filterByKey(data, heads, key) {
    // Making sure data contains the informations about key
    if (heads.indexOf(key) == -1) {ui.alert("Filtrage", "Erreur : pas de question correspondant au critère sélectionné.", ui.ButtonSet.OK); return;}
    // Query and filter
    while (true) {
      try {
        // User input
        let keyList = uniqueValues(key, data),
            keyIdx = parseInt(ui.prompt("Filtrage", `Entrez un numéro parmi la liste suivante : ${listToQuery(keyList)}`, ui.ButtonSet.OK).getResponseText(), 10);
        
        if (keyIdx != "") {
          return data.filter(row => row[key] == keyList[keyIdx - 1]);
        }
      } catch(e) {
        ui.alert("Filtrage", "Entrée invalide, veuillez recommencer", ui.ButtonSet.OK);
        Logger.log(`An error occured when filtering given set of data : ERROR : ${e}, KEY ${key}, DATA : ${data}`);
      }
    }
  }