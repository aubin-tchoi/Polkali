/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Extracting data from an array of Objects and returning a dataTable

function contacts(obj) {
    let dataTable = Charts.newDataTable(),
        conversionChart = [];
    // Columns : month, all 4 states a mission can go through
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    Object.values(STATES).forEach(function (state) {
        dataTable.addColumn(Charts.ColumnType.NUMBER, state);
    });
    // Rows
    MONTH_LIST.forEach(function (month) {
        let dataRow = [MONTH_NAMES[month]],
            conversionChartRow = [];
        Object.values(STATES).forEach(function (state) {
            // Counting first every contact who find themselves in a more advanced state than state
            let val = obj.filter(row => sameMonth(row, month) && Object.values(STATES).indexOf(row[HEADS["état"]]) >= Object.values(STATES).indexOf(state)).length;
            if (state == STATES["devis"] || state == STATES["negoc"]) {
                // In both cases we forgot to count ppl to whom we sent a devis but who didn't convert it into a mission
                val += obj.filter(row => sameMonth(row, month) && Object.values(STATES_BIS).includes(row[HEADS["état"]]) && row[HEADS["devis"]]).length
            } else if (state == STATES["rdv"]) {
                // We must also count contacts who didn't lead to a mission
                val += obj.filter(row => sameMonth(row, month) && Object.values(STATES_BIS).includes(row[HEADS["état"]])).length
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
    // Columns : month, all 3 conversion rates
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${STATES["rdv"]} et ${STATES["devis"]}`);
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${STATES["devis"]} et ${STATES["negoc"]}`);
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${STATES["negoc"]} et ${STATES["etude"]}`);
    // Rows
    MONTH_LIST.forEach(function (month, idx) {
        dataTable.addRow([MONTH_NAMES[month], prcnt(conversionChart[idx][1], conversionChart[idx][0]), prcnt(conversionChart[idx][2], conversionChart[idx][1]), prcnt(conversionChart[idx][3], conversionChart[idx][2])]);
    });
    return dataTable;
}

function turnover(obj) {
    let dataTable = Charts.newDataTable();
    // Columns : month, potential turnover and contractualized turnover
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA sur étude potentielle");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA sur étude signée");
    // Rows
    MONTH_LIST.forEach(function (month) {
        let ca_pot = 0,
            ca_sig = 0;
        // Potential turnover is the sum of every expected gain
        obj.filter(row => sameMonth(row, month)).forEach(function (row) {
            // /!\ row[HEADS["confiance"]] isn't actually a percentage since we used method getValues on the range instead of getDisplayValues
            ca_pot += (parseInt(row[HEADS["caPot"]], 10) * row[HEADS["confiance"]]) || 0;
        });
        // Contractualized turnover is computed based on missions who are state STATES["study"]
        obj.filter(row => sameMonth(row, month) && row[HEADS["état"]] == STATES["etude"]).forEach(function (row) {
            ca_sig += parseInt(row[HEADS["caPot"]], 10) || 0;
        });
        dataTable.addRow([MONTH_NAMES[month], ca_pot, ca_sig]);
    });
    return dataTable;
}

function conversionRateByContact(obj) {
    let dataTable = Charts.newDataTable();
    // Columns : contactType, number of contacts and conversionRate
    dataTable.addColumn(Charts.ColumnType.STRING, "Type de contact");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre de contacts");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Taux de conversion global");
    // Rows
    uniqueValues(HEADS["typeContact"], obj).forEach(function (type) {
        let objFiltered = obj.filter(row => row[HEADS["typeContact"]] == type),
            conversionRate = prcnt(objFiltered.filter(row => row[HEADS["état"]] == STATES["etude"]).length, objFiltered.length);
        dataTable.addRow([type, objFiltered.length, conversionRate]);
    });
    return dataTable;
}