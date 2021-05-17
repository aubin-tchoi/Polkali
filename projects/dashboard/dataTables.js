/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Extracting data from an array of Objects and returning a dataTable

// Répartition des contacts par mois
function contacts(data) {
    let dataTable = Charts.newDataTable(),
        conversionChart = [];
    // Columns : month, all 4 states a mission can go through
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    Object.values(ETAT_PROSP).forEach(function (state) {
        dataTable.addColumn(Charts.ColumnType.NUMBER, state);
    });
    // Rows
    MONTH_LIST.forEach(function (month) {
        let dataRow = [`${MONTH_NAMES[month.month]} ${month.year}`],
            conversionChartRow = [];
        Object.values(ETAT_PROSP).forEach(function (state) {
            // Counting first every contact who find themselves in a more advanced state than state
            let val = data.filter(row => sameMonth(row, month) && Object.values(ETAT_PROSP).indexOf(row[HEADS.état]) >= Object.values(ETAT_PROSP).indexOf(state)).length;
            if (state == ETAT_PROSP.devis || state == ETAT_PROSP.negoc) {
                // In both cases we forgot to count ppl to whom we sent a devis but who didn't convert it into a mission
                val += data.filter(row => sameMonth(row, month) && Object.values(ETAT_PROSP_BIS).includes(row[HEADS.état]) && row[HEADS.devis]).length
            } else if (state == ETAT_PROSP.rdv) {
                // We must also count contacts who didn't lead to a mission
                val += data.filter(row => sameMonth(row, month) && Object.values(ETAT_PROSP_BIS).includes(row[HEADS.état])).length
            }
            conversionChartRow.push(val);
            dataRow.push(val);
        });
        conversionChart.push(conversionChartRow);
        dataTable.addRow(dataRow);
    });
    return [dataTable, conversionChart];
}

// Taux de conversion
function conversionRate(conversionChart) {
    let dataTable = Charts.newDataTable();
    // Columns : month, all 3 conversion rates
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion global`);
    // Rows
    MONTH_LIST.forEach(function (month, idx) {
        dataTable.addRow([`${MONTH_NAMES[month.month]} ${month.year}`, prcnt(conversionChart[idx][1], conversionChart[idx][0]) / 100 * prcnt(conversionChart[idx][2], conversionChart[idx][1]) / 100 * prcnt(conversionChart[idx][3], conversionChart[idx][2])]);
    });
    return dataTable;
}

// Taux de conversion
function conversionRateByType(conversionChart) {
    let dataTable = Charts.newDataTable();
    // Columns : month, all 3 conversion rates
    dataTable.addColumn(Charts.ColumnType.STRING, "Étape de la conversion");
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${ETAT_PROSP.rdv} et ${ETAT_PROSP.devis}`);
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${ETAT_PROSP.devis} et ${ETAT_PROSP.negoc}`);
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${ETAT_PROSP.negoc} et ${ETAT_PROSP.etude}`);
    // Rows
    dataTable.addRow([`${ETAT_PROSP.rdv} -> ${ETAT_PROSP.devis}`, prcnt(conversionChart.reduce((acc, val) => acc += val[1], 0), conversionChart.reduce((acc, val) => acc += val[0], 0))]);
    dataTable.addRow([`${ETAT_PROSP.devis} -> ${ETAT_PROSP.negoc}`, prcnt(conversionChart.reduce((acc, val) => acc += val[2], 0), conversionChart.reduce((acc, val) => acc += val[1], 0))]);
    dataTable.addRow([`${ETAT_PROSP.negoc} -> ${ETAT_PROSP.etude}`, prcnt(conversionChart.reduce((acc, val) => acc += val[3], 0), conversionChart.reduce((acc, val) => acc += val[2], 0))]);
    return dataTable;
}

// CA potentiel et CA signé
function turnover(data) {
    let dataTable = Charts.newDataTable();
    // Columns : month, potential turnover and contractualized turnover
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA sur étude potentielle");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA sur étude signée");
    // Rows
    MONTH_LIST.forEach(function (month) {
        let ca_pot = 0,
            ca_sig = 0;
        data.filter(row => sameMonth(row, month)).forEach(function (row) {
            // /!\ row[HEADS.confiance] isn't actually a percentage since we used method getValues on the range instead of getDisplayValues
            // Potential turnover is the sum of every expected gain
            ca_pot += (parseInt(row[HEADS.caPot], 10) * row[HEADS.confiance]) || 0;
            // Contractualized turnover is computed based on missions who are state ETAT_PROSP.study
            if (row[HEADS.état] == ETAT_PROSP.etude) {
                ca_sig += parseInt(row[HEADS.caPot], 10) || 0;
            }
            Logger.log(row);
        });
        dataTable.addRow([`${MONTH_NAMES[month.month]} ${month.year}`, ca_pot, ca_sig]);
    });
    return dataTable;
}

// Répartition des contacts par type
function contactType(data) {
    // Creating a DataTable with the proportion of each contact type
    let dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType.STRING, "Type de contact");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion");
    Logger.log(`Valeurs uniques : ${uniqueValues(HEADS.typeContact, data)}`);
    uniqueValues(HEADS.typeContact, data).forEach(function (type) {
        dataTable.addRow([type, data.filter(row => row[HEADS.typeContact] == type).length / data.length]);
    });
    return dataTable;
}

// Taux de conversion pour chaque type de contact
function conversionRateByContact(data) {
    let dataTable = Charts.newDataTable();
    // Columns : contactType, number of contacts and conversionRate
    dataTable.addColumn(Charts.ColumnType.STRING, "Type de contact");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre de contacts");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Taux de conversion global");
    // Rows
    uniqueValues(HEADS.typeContact, data).forEach(function (type) {
        let objFiltered = data.filter(row => row[HEADS.typeContact] == type),
            conversionRate = prcnt(objFiltered.filter(row => row[HEADS.état] == ETAT_PROSP.etude).length, objFiltered.length);
        dataTable.addRow([type, objFiltered.length, conversionRate]);
    });
    return dataTable;
}

// Répartition des prix des études
function priceRange(data, lowerBound, higherBound, nbrRanges) {
    let dataTable = Charts.newDataTable();
    // Columns : number of missions for different price ranges
    dataTable.addColumn(Charts.ColumnType.STRING, "Intervalle de prix");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre d'études");
    // Rows : creating multiple ranges of prices
    let width = parseInt((higherBound - lowerBound) / nbrRanges, 10),
        priceRanges = Array.from(Array(nbrRanges).keys()).map((_, idx) => lowerBound + idx * width);
    priceRanges.forEach(price => {
        dataTable.addRow([`${price} - ${price + width}`, data.filter(row => (price <= row[HEADS.prix] && row[HEADS.prix] < price + width)).length]);
    });
    return dataTable;
}

// Nombre d'étude par nombre de JEH
function JEHRange(data) {
    let dataTable = Charts.newDataTable();
    // Columns : number of missions for different price ranges
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre de JEH");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre d'études");
    // Rows : creating multiple ranges of prices
    let lowerBound = Math.min(...data.map(row => parseInt(row[HEADS.JEH], 10))),
        higherBound = Math.max(...data.map(row => parseInt(row[HEADS.JEH], 10)));
    let JEHNumber = Array.from(Array(higherBound - lowerBound + 1).keys()).map(number => number + lowerBound);
    JEHNumber.forEach(number => {
        dataTable.addRow([number, data.filter(row => parseInt(row[HEADS.JEH], 10) == number).length])
    });
    return [dataTable, JEHNumber];
}

// Nombre d'étude par nombre de JEH
function lengthRange(data) {
    let dataTable = Charts.newDataTable();
    // Columns : number of missions for different price ranges
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Durée (en nombre de semaines)");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre d'études");
    // Rows : creating multiple ranges of prices
    let lowerBound = Math.min(...data.map(row => parseInt(row[HEADS.durée], 10))),
        higherBound = Math.max(...data.map(row => parseInt(row[HEADS.durée], 10)));
    let lengthNumber = Array.from(Array(higherBound - lowerBound + 1).keys()).map(number => number + lowerBound);
    lengthNumber.forEach(number => {
        dataTable.addRow([number, data.filter(row => parseInt(row[HEADS.durée], 10) == number).length])
    });
    return [dataTable, lengthNumber];
}

// Proportion du CA due aux alumni
function alumniContribution(data) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Alumni/Non alumni");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion du CA");
    // Rows
    let ca = data.reduce((sum, row) => sum += parseInt(row[HEADS.prix], 10) || 0, 0);
    dataTable.addRow(["Alumni", prcnt(data.filter(row => row[HEADS.alumni] || false).reduce((sum, row) => sum += parseInt(row[HEADS.prix], 10) || 0, 0), ca)]);
    dataTable.addRow(["Non Alumni", prcnt(data.filter(row => !(row[HEADS.alumni] || false)).reduce((sum, row) => sum += parseInt(row[HEADS.prix], 10) || 0, 0), ca)]);
    return dataTable;
}

// Répartition des contacts par secteur d'activité
function contactByDomain(data) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Domaine de compétences");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion des contacts");
    // Rows
    uniqueValues(HEADS.domaine, data).forEach(currentDomain => {
        dataTable.addRow([currentDomain, data.filter(row => row[HEADS.domaine] == currentDomain).length / data.length]);
    });
    return dataTable;
}

// Proportion du CA venant de la prospection
function prospectionTurnover(data) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Prospection ou non");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion du CA");
    // Rows
    let ca = data.reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0);
    dataTable.addRow(["Prospection", prcnt(data.filter(row => !([CONTACT_TYPE.site, CONTACT_TYPE.spontané].includes(row[HEADS.typeContact]))).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0), ca)]);
    dataTable.addRow(["Hors prospection", prcnt(data.filter(row => [CONTACT_TYPE.site, CONTACT_TYPE.spontané].includes(row[HEADS.typeContact])).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0), ca)]);
    return dataTable;
}

// Proportion du CA venant de chaque secteur
function turnoverBySector(data) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Secteur");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion du CA");
    // Rows
    let ca = data.reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0);
    uniqueValues(HEADS.secteur, data).forEach(currentSector => {
        dataTable.addRow([currentSector, prcnt(data.filter(row => row[HEADS.secteur] == currentSector).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0), ca)]);
    });
    return dataTable;
}

// Performance par secteur
function performanceBySector(data) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Secteur");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre d'études");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA (en milliers d'€)");
    // Rows
    uniqueValues(HEADS.secteur, data).forEach(currentSector => {
        dataTable.addRow([currentSector, data.filter(row => row[HEADS.secteur] == currentSector).length, data.filter(row => row[HEADS.secteur] == currentSector).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0) / 1000]);
    });
    return dataTable;
}

// Performance par type d'entreprise
function performanceByCompanyType(data) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Type d'entreprise");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre d'études");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA (en milliers d'€)");
    // Rows
    uniqueValues(HEADS.typeEntreprise, data).forEach(currentType => {
        dataTable.addRow([currentType, data.filter(row => row[HEADS.typeEntreprise] == currentType).length, data.filter(row => row[HEADS.typeEntreprise] == currentType).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0) / 1000]);
    });
    return dataTable;
}

// Proportion du CA venant de chaque type d'entreprise
function turnoverByCompanyType(data) {
    let dataTable = Charts.newDataTable();
    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Type d'entreprise");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion du CA");
    // Rows
    let ca = data.reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0);
    uniqueValues(HEADS.typeEntreprise, data).forEach(currentType => {
        dataTable.addRow([currentType, prcnt(data.filter(row => row[HEADS.typeEntreprise] == currentType).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0), ca)]);
    });
    return dataTable;
}

// Nombre de contact du au site par mois
function contactBySite(data) {
    let dataTable = Charts.newDataTable();
    let dateData = data.map(row => row[row.indexOf("Horodateur")]);
    // Columns : month, all 3 conversion rates
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Nombre de contact par le site`);
    // Rows
    MONTH_LIST.forEach(function (month, idx) {
        dataTable.addRow([`${MONTH_NAMES[month.month]} ${month.year}`,dateData.filter(date => (date.getMonth() == month.month && date.getFullYear() == month.year)).length]);
    });
    return dataTable;
}
