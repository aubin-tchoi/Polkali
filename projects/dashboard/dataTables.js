/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Extracting data from an array of Objects and returning a dataTable

/**
 * Calcul de la réparition des contacts selon la colonne dont le nom est spécifié.
 * @param {string} key Nom de la colonne selon laquelle sera effectuée la répartition.
 * @param {Array} data Données d'entrée.
 * @returns {DataTable} Table des données permettant de créer le graphe correspondant.
 */
function totalDistribution(key, data, options) {
    let dataTable = Charts.newDataTable();

    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, key);
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion des contacts");

    // Rows
    uniqueValues(key, data).forEach(type => {
        dataTable.addRow([type, data.filter(row => row[key] == type).length / data.length]);
    });

    return {
        dataTable: dataTable,
        options: options
    };
}

/**
 * Calcul de la répartition du CA selon les valeurs présentes dans une colonne du fichier de suivi de la prospection (transposable au suivi d'étude en remplaçant caPot en prix).
 * @param {string} key Nom de la colonne selon laquelle sera évaluée la répartition du CA.
 * @param {Array} data Données d'entrée.
 * @returns {DataTable} Table des données permettant de créer le graphe correspondant.
 */
function turnoverDistribution(key, data, options, price = HEADS.caPot) {
    let dataTable = Charts.newDataTable();

    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, key);
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion du CA");

    // Rows
    let ca = data.reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0);
    Logger.log(`CA calculé à partir du suivi de la prospection : ${ca}`);
    uniqueValues(key, data).forEach(currentType => {
        dataTable.addRow([currentType, prcnt(data.filter(row => row[key] == currentType).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0), ca)]);
    });

    return {
        dataTable: dataTable,
        options: options
    };
}

/**
 * Calcul du taux de conversion mensuel pour chaque étape.
 * @param {Array} data Données d'entrée.
 * @returns {Array} Tableau 2D décrivant le taux de conversion entre chaque étape pour chaque mois.
 */
function conversionTable(data) {
    let conversionChart = [];
    
    // Rows
    MONTH_LIST.forEach(function (month) {
        let conversionChartRow = [];
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
        });
        conversionChart.push(conversionChartRow);
    });

    return conversionChart;
}

/**
 * Calcul du CA et du nombre d'études selon les valeurs présentes dans une colonne du fichier de suivi de la prospection (transposable au suivi d'étude en remplaçant caPot en prix).
 * @param {string} key Nom de la colonne selon laquelle sera évaluée la performance de la JE.
 * @param {Array} data Données d'entrée.
 * @returns {Array} Table des données permettant de créer le graphe correspondant
 * et tableau des valeurs à indiquer selon l'axe vertical.
 */
function performance(key, data, options, price = HEADS.caPot) {
    let dataTable = Charts.newDataTable();

    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, key);
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre d'études");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA (en milliers d'€)");

    // Rows
    let maxTick = 0;
    uniqueValues(key, data).forEach(currentType => {
        maxTick = Math.max(maxTick, data.filter(row => row[key] == currentType).length, parseInt(data.filter(row => row[key] == currentType).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0) / 1000, 10) + 1);
        dataTable.addRow([currentType, data.filter(row => row[key] == currentType).length, data.filter(row => row[key] == currentType).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0) / 1000]);
    });

    // Updating the options
    options["vticks"] = Array.from(Array(maxTick + 1).keys());

    return {
        dataTable: dataTable,
        options: options
    };
}

/**
 * Calcul du nombre d'études selon les valeurs présente dans une colonne du fichier étudié.
 * @param {string} key Nom de la colonne selon laquelle sera compté le nombre d'études.
 * @param {Array} data Données d'entrée.
 * @returns {Array} Table des données permettant de créer le graphe correspondant
 * et tableau des valeurs à indiquer selon l'axe horizontal.
 */
function numberOfMissions(key, data, options) {
    let dataTable = Charts.newDataTable();

    // Columns
    dataTable.addColumn(Charts.ColumnType.NUMBER, key);
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre d'études");

    // Rows
    let lowerBound = Math.min(...data.map(row => parseInt(row[key], 10))),
        higherBound = Math.max(...data.map(row => parseInt(row[key], 10)));
    let values = Array.from(Array(higherBound - lowerBound + 1).keys()).map(number => number + lowerBound);
    values.forEach(number => {
        dataTable.addRow([number, data.filter(row => parseInt(row[HEADS.durée], 10) == number).length])
    });

    // Updating the options
    options["hticks"] = values;

    return {
        dataTable: dataTable,
        options: options
    };
}

/**
 * Calcul du taux de conversion décliné selon les différentes catégories présentes dans une colonne spécifiée (uniquement pour le suivi de la prospection).
 * @param {string} key Nom de la colonne selon laquelle sera évalué le taux de conversion.
 * @param {Array} data Données d'entrée.
 * @returns {Array} Table des données permettant de créer le graphe correspondant.
 */
function conversionRate(key, data, options) {
    let dataTable = Charts.newDataTable();

    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, key);
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre de contacts");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Taux de conversion global");

    // Rows
    uniqueValues(key, data).forEach(function (type) {
        let objFiltered = data.filter(row => row[key] == type),
            conversionRate = prcnt(objFiltered.filter(row => row[HEADS.état] == ETAT_PROSP.etude).length, objFiltered.length);
        dataTable.addRow([type, objFiltered.length, conversionRate]);
    });

    return {
        dataTable: dataTable,
        options: options
    };
}

/**
 * Calcul de la répartition du CA selon les valeurs présentes dans une colonne du fichier de suivi de la prospection (transposable au suivi de la prospection en remplaçant prix par caPot).
 * On suppose que la colonne contient des cases à cocher (deux catégories uniquement).
 * @param {string} key Nom de la colonne selon laquelle sera évaluée la répartition du CA.
 * @param {Array} data Données d'entrée.
 * @returns {DataTable} Table des données permettant de créer le graphe correspondant.
 */
function turnoverDistributionBinary(key, data, options, price = HEADS.prix) {
    let dataTable = Charts.newDataTable();

    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, `${key}/Hors ${key}`);
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion du CA");

    // Rows
    let ca = data.reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0);
    Logger.log(`CA calculé à partir du suivi d'études : ${ca}.`);
    Logger.log(`CA alumni : ${data.filter(row => row[key] || false).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0)}.`);
    dataTable.addRow([key, prcnt(data.filter(row => row[key] || false).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0), ca)]);
    dataTable.addRow([`Hors ${key}`, prcnt(data.filter(row => !(row[key] || false)).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0), ca)]);

    return {
        dataTable: dataTable,
        options: options
    };
}

// Répartition des contacts par mois
function contacts(data, options) {
    let dataTable = Charts.newDataTable();

    // Columns : month, all 4 states a mission can go through
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    Object.values(ETAT_PROSP).forEach(function (state) {
        dataTable.addColumn(Charts.ColumnType.NUMBER, state);
    });

    // Rows
    MONTH_LIST.forEach(function (month) {
        let dataRow = [`${MONTH_NAMES[month.month]} ${month.year}`];
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
            dataRow.push(val);
        });
        dataTable.addRow(dataRow);
    });

    return {
        dataTable: dataTable,
        options: options
    };
}


// Taux de conversion global
function conversionRateOverTime(data, options) {
    let conversionChart = conversionTable(data),
        dataTable = Charts.newDataTable();
    
    // Columns : month, all 3 conversion rates
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion global`);
    
    // Rows
    MONTH_LIST.forEach(function (month, idx) {
        dataTable.addRow([`${MONTH_NAMES[month.month]} ${month.year}`, prcnt(conversionChart[idx][1], conversionChart[idx][0]) / 100 * prcnt(conversionChart[idx][2], conversionChart[idx][1]) / 100 * prcnt(conversionChart[idx][3], conversionChart[idx][2])]);
    });

    return {
        dataTable: dataTable,
        options: options
    };
}

// Taux de conversion par étape
function conversionRateByType(data, options) {
    let conversionChart = conversionTable(data),
        dataTable = Charts.newDataTable();

    // Columns : month, all 3 conversion rates
    dataTable.addColumn(Charts.ColumnType.STRING, "Étape de la conversion");
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${ETAT_PROSP.rdv} et ${ETAT_PROSP.devis}`);
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${ETAT_PROSP.devis} et ${ETAT_PROSP.negoc}`);
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Taux de conversion entre ${ETAT_PROSP.negoc} et ${ETAT_PROSP.etude}`);
    
    // Rows
    dataTable.addRow([`${ETAT_PROSP.rdv} -> ${ETAT_PROSP.devis}`, prcnt(conversionChart.reduce((acc, val) => acc += val[1], 0), conversionChart.reduce((acc, val) => acc += val[0], 0))]);
    dataTable.addRow([`${ETAT_PROSP.devis} -> ${ETAT_PROSP.negoc}`, prcnt(conversionChart.reduce((acc, val) => acc += val[2], 0), conversionChart.reduce((acc, val) => acc += val[1], 0))]);
    dataTable.addRow([`${ETAT_PROSP.negoc} -> ${ETAT_PROSP.etude}`, prcnt(conversionChart.reduce((acc, val) => acc += val[3], 0), conversionChart.reduce((acc, val) => acc += val[2], 0))]);

    return {
        dataTable: dataTable,
        options: options
    };
}

// CA potentiel et CA signé
function turnover(data, options) {
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

    return {
        dataTable: dataTable,
        options: options
    };
}

// Répartition des prix des études
function priceRange(data, options) {
    let dataTable = Charts.newDataTable();

    // Columns : number of missions for different price ranges
    dataTable.addColumn(Charts.ColumnType.STRING, "Intervalle de prix");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre d'études");

    // Rows : creating multiple ranges of prices
    let width = parseInt((options.higherBound - options.lowerBound) / options.nbrRanges, 10),
        priceRanges = Array.from(Array(options.nbrRanges).keys()).map((_, idx) => options.lowerBound + idx * width);
    priceRanges.forEach(price => {
        dataTable.addRow([`${price} - ${price + width}`, data.filter(row => (price <= row[HEADS.prix] && row[HEADS.prix] < price + width)).length]);
    });

    // Updating the options
    delete options.higherBound;
    delete options.lowerBound;
    delete options.nbrRanges;

    return {
        dataTable: dataTable,
        options: options
    };
}

// Proportion du CA venant de la prospection
function prospectionTurnover(data, options) {
    let dataTable = Charts.newDataTable();

    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, "Prospection ou non");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion du CA");

    // Rows
    let ca = data.reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0);
    dataTable.addRow(["Prospection", prcnt(data.filter(row => !([CONTACT_TYPE.site, CONTACT_TYPE.spontané].includes(row[HEADS.typeContact]))).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0), ca)]);
    dataTable.addRow(["Hors prospection", prcnt(data.filter(row => [CONTACT_TYPE.site, CONTACT_TYPE.spontané].includes(row[HEADS.typeContact])).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0), ca)]);

    return {
        dataTable: dataTable,
        options: options
    };
}

function performanceByContact(key, data, options) {
    let dataTable = Charts.newDataTable();

    // Columns
    dataTable.addColumn(Charts.ColumnType.STRING, key);
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre de devis envoyés");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "Nombre d'études signées");
    dataTable.addColumn(Charts.ColumnType.NUMBER, "CA (en milliers d'euros)");

    // Rows
    let maxTick = 0;
    uniqueValues(key, data).forEach(currentType => {
        maxTick = Math.max(
            maxTick,
            data.filter(row => !!row[HEADS.devis] && row[key] == currentType).length,
            data.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).length,
            parseInt(data.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0) / 1000, 10) + 1);
        dataTable.addRow([
            currentType,
            data.filter(row => !!row[HEADS.devis] && row[key] == currentType).length,
            data.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).length,
            data.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0) / 1000
        ]);
    });

    // Updating the options
    options["vticks"] = Array.from(Array(maxTick + 1).keys());

    return {
        dataTable: dataTable,
        options: options
    };
}

// Nombre de contact du au site par mois
function contactBySite(data, options) {
    let dataTable = Charts.newDataTable(),
        dateData = data.map(row => [row["Horodateur"], row["Entreprise"]]);

    // Columns : month, all 3 conversion rates
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.NUMBER, `Nombre de contact par le site`);

    // Rows
    MONTH_LIST.forEach(function (month, idx) {
        dataTable.addRow([`${MONTH_NAMES[month.month]} ${month.year}`, dateData.filter(row => (row[0].getMonth() == month.month && row[0].getFullYear() == month.year && row[1] != "test")).length]);
    });

    return {
        dataTable: dataTable,
        options: options
    };
}

function mailSent(obj, options) {
    let dataTable = Charts.newDataTable();

    // Columns : month, nb of quali mail, nb of quanti mail
    dataTable.addColumn(Charts.ColumnType.STRING, "Mois");
    dataTable.addColumn(Charts.ColumnType.STRING, "Mails Quali");
    dataTable.addColumn(Charts.ColumnType.STRING, "Mails Quanti");

    // Rows
    tables = []
    Object.values(obj).forEach(tableId => {
        sheet = SpreadSheet.openById(tableId).getSheetByName("BDD");
        arr = sheet.getRange(2, 2, sheet.getLastRow(), sheet.getLastColumn()).getValues();
        heads = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
        data = arr.map(row => ArrayToObj(row, heads));
        tables.push(data);
    })

    MONTH_LIST.forEach(function (month, idx) {
        dataTable.addRow([`${MONTH_NAMES[month.month]} ${month.year}`, tables.reduce(compteur, data => (data.filter(row => row["Mail Normal"] == Quali).length + compteur), 0), tables.reduce(compteur, data => (data.filter(row => row["Mail Normal"] == VRAI).length + compteur), 0)]);
    });

    return {
        dataTable: dataTable,
        options: options
    };
}
