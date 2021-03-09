function syncOnEdit(rowidx) {
    // Updating sheet sheet with row row
    function update(sheet, row) {
        // Checking whether PEP actually did obtain the mission or not
        if (row[HEADS["état"]] != stateList[3]) {
            return;
        }

        // ref is gonna be the mission's reference, obtained by incrementing the most recent one
        const heads = sheet.getRange(1, 2, 1, (sheet.getLastColumn() - 1)).getValues().shift(),
            lastRow = manuallyGetLastRow(sheet),
            ref = `'20e${Math.floor(sheet.getRange(lastRow, 2, 1, 1).getValues()[0][0].toString().match(/([0-9]+)/gi)[1]) + 1}`;

        // Creating a new row with the destination sheet's header's informations
        let rnew = [row["Entreprise"]];
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