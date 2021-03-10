/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Extracting data from an array of Objects and returning a dataTable

function contacts(obj) {
    let dataTable = Charts.newDataTable(),
        conversionChart = [];
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    STATE_LIST.forEach(function (state) {
        dataTable.addColumn(Charts.ColumnType.NUMBER, state);
    });
    // Rows
    MONTH_LIST.forEach(function (month) {
        let dataRow = [MONTH_NAMES[month]];
        let conversionChartRow = [];
        STATE_LIST.forEach(function (state) {
            let val = obj.filter(row => parseInt(row[HEADS["premierContact"]].getMonth(), 10) == month && STATE_LIST.indexOf(row[HEADS["état"]]) >= STATE_LIST.indexOf(state)).length;
            if (state == "Devis rédigé et envoyé" || state == "En négociation") {
                val += obj.filter(row => parseInt(row[HEADS["premierContact"]].getMonth(), 10) == month && (row[HEADS["état"]] == "Sans suite" || row[HEADS["état"]] == "A relancer") && row[HEADS["devis"]]).length
            } else if (state == "Premier RDV réalisé") {
                val += obj.filter(row => parseInt(row[HEADS["premierContact"]].getMonth(), 10) == month && (row[HEADS["état"]] == "Sans suite" || row[HEADS["état"]] == "A relancer")).length
            }
            conversionChartRow.push(val);
            dataRow.push(val);
        });
        conversionChart.push(conversionChartRow);
        dataTable.addRow(dataRow);
    });
    Logger.log(`Conversion : ${conversionChart}`);
    return [dataTable, conversionChart];
}

function conversionRate(conversionChart) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${STATE_LIST[0]} et ${STATE_LIST[1]}`);
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${STATE_LIST[1]} et ${STATE_LIST[2]}`);
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${STATE_LIST[2]} et ${STATE_LIST[3]}`);
    // Rows
    for (let idx in MONTH_LIST) {
        Logger.log(`Taux de conversion : ${[prcnt(conversionChart[idx][1], conversionChart[idx][0]), prcnt(conversionChart[idx][2], conversionChart[idx][1]), prcnt(conversionChart[idx][3], conversionChart[idx][2])]}`);
    }
    MONTH_LIST.forEach(function (month, idx) {
        dataTable.addRow([MONTH_NAMES[month], prcnt(conversionChart[idx][1], conversionChart[idx][0]), prcnt(conversionChart[idx][2], conversionChart[idx][1]), prcnt(conversionChart[idx][3], conversionChart[idx][2])]);
    });
    return dataTable;
}

function turnover(obj) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA sur étude potentielle");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA sur étude signée");
    // Rows
    MONTH_LIST.forEach(function (month) {
        let ca_pot = 0,
            ca_sig = 0;
        obj.filter(row => parseInt(row[HEADS["premierContact"]].getMonth(), 10) == month && row[HEADS["état"]] != "Etude obtenue").forEach(function (row) {
            if (Object.keys(row).includes("Prix potentiel de l'étude  € (HT)")) {
                ca_pot += parseInt(row["Prix potentiel de l'étude  € (HT)"], 10) * ((parseInt(row["Pourcentage de confiance à la conversion en réelle étude"], 10) > 0) ? parseInt(row["Pourcentage de confiance à la conversion en réelle étude"], 10) / 100 : 1);
            }
        });
        obj.filter(row => parseInt(row[HEADS["premierContact"]].getMonth(), 10) == month && row[HEADS["état"]] == "Etude obtenue").forEach(function (row) {
            if (Object.keys(row).includes("Prix potentiel de l'étude  € (HT)")) {
                ca_sig += parseInt(row["Prix potentiel de l'étude  € (HT)"], 10);
            }
        });
        dataTable.addRow([MONTH_NAMES[month], ca_pot, ca_sig]);
    });
    return dataTable;
}

function conversionRateByContact(obj) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Type de contact");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre de contacts");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Taux de conversion");
    // Rows
    uniqueValues("Type de contact", obj).forEach(function (value) {
        let objFiltered = obj.filter(row => row["Type de contact"] == value);
        let conversionRate = prcnt(objFiltered.filter(row => row["État"] == "Etude obtenue").length, objFiltered.length);
        dataTable.addRow([value, objFiltered.length, conversionRate]);
    });
    return dataTable;
}