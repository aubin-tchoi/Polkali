/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Copying data corresponding to 1 mission from suivi_prospection to suivi_étude

function syncOnEdit(rowidx) {
    // Updating sheet sheet with row row
    function update(sheet, row) {
        // Checking whether PEP actually did obtain the mission or not
        if (row[HEADS["état"]] != STATES["etude"]) {
            return;
        }

        // ref is gonna be the mission's reference, obtained by incrementing the most recent one
        const heads = sheet.getRange(1, 2, 1, (sheet.getLastColumn() - 1)).getValues().shift(),
            lastRow = manuallyGetLastRow(sheet),
            ref = `'21e${Math.floor(sheet.getRange(lastRow, 2, 1, 1).getValues()[0][0].toString().match(/([0-9]+)/gi)[1]) + 1}`;

        // Creating a new row with the destination sheet's header's informations
        let rnew = [row[HEADS["entreprise"]]];
        heads.forEach(function (el) {
            rnew.push(el == "Référence de l'étude" ? ref : (el == "État") ? "Devis accepté" : row[el]);
        });
        sheet.getRange((lastRow + 1), 1, 1, rnew.length).setValues([rnew]);
    }

    // -- MAIN --
    const sheetscr = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
        sheetdst = SpreadsheetApp.openById("1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04").getSheetByName("Suivi");

    // heads is the sheet's header, data a js object of the edited row
    let heads = sheetscr.getRange(1, 1, 1, sheetscr.getLastColumn()).getValues().shift(),
        data = sheetscr.Range(rowidx, 1, 1, sheetscr.getLastColumn()).getValues().map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {}))).shift();

    // Synchronizing edited row
    update(sheetdst, data);
}

function syncOnSelec() {
    // Updating sheet sheet with row row
    function update(sheet, row) {
        // Checking whether PEP actually did obtain the mission or not
        if (row[HEADS["état"]] != STATES["etude"]) {
            ui.alert("Entrée invalide", `L'étude confiée par l'entreprise ${row[HEADS["entreprise"]]} ne correspond pas à une étude obtenue.`, ui.ButtonSet.OK);
            return;
        }

        // ref is gonna be the mission's reference, obtained by incrementing the most recent one
        const heads = sheet.getRange(1, 2, 1, (sheet.getLastColumn() - 1)).getValues().shift(),
            lastRow = manuallyGetLastRow(sheet),
            ref = `'21e${Math.floor(sheet.getRange(lastRow, 2, 1, 1).getValues()[0][0].toString().match(/([0-9]+)/gi)[1]) + 1}`;

        // Creating a new row with the destination sheet's header's informations
        let rnew = [row[HEADS["entreprise"]]];
        heads.forEach(function (el) {
            rnew.push(el == "Référence de l'étude" ? ref : (el == "État") ? "Devis accepté" : row[el]);
        });
        sheet.getRange((lastRow + 1), 1, 1, rnew.length).setValues([rnew]);
    }

    // -- MAIN --
    const sheetscr = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
        sheetdst = SpreadsheetApp.openById("1h1rObRvdb2GKxdzgTZD3Kvwde_TUTlTOLC3l8BKzOBQ").getSheetByName("Suivi"),
        ui = SpreadsheetApp.getUi();

    // Confirm selection
    let confirmSelection = ui.alert("Synchronisation des données", "Vous devez préalablement sélectionner la ligne à synchroniser (par exemple en cliquant sur le numéro à gauche). \n Confirmez-vous votre sélection ?", ui.ButtonSet.YES_NO);
    if (confirmSelection == ui.Button.NO) {
        return;
    }

    // Loading screen
    displayLoadingScreen("Synchronisation ..");

    // heads is the sheet's header, data a 2D-array representation of the selected values, and obj its 'array of js object' representation
    let heads = sheetscr.getRange(1, 1, 1, sheetscr.getLastColumn()).getValues().shift(),
        data = sheetscr.getActiveRange().getValues(),
        obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

    // Checking the selection again (empty or not)
    if (obj.length == 0) {
        ui.alert("Pas de ligne sélectionnée, veuillez recommencer.");
        return;
    }

    // Synchronizing each selected row
    obj.forEach(function (row) {
        update(sheetdst, row);
    });

    // Confirmation
    let imgUrl = "https://raw.githubusercontent.com/aubin-tchoi/Polkali/main/images/thumbsUp.png",
        sheetUrl = "https://docs.google.com/spreadsheets/d/1gmJLAKvUOYFeS32raOiSTYiE_ozr7YSk26Y0t0blm04/edit#gid=0",
        operationSuccess = HtmlService
        .createHtmlOutput(`<span style='font-size: 16pt;'> <span style="font-family: 'roboto', sans-serif;">L'opération a été effectuée avec succès, veuillez remplir manuellement le nom de l'étude dans le <a href="${sheetUrl}">suivi des études</a>.<br/><br/> La bise</span></span><p style="text-align:center;"><img src="${imgUrl}" alt="C'est la PEP qui régale !" width="130" height="131"></p>`);
    ui.showModalDialog(operationSuccess, "Synchronisation des données");
}