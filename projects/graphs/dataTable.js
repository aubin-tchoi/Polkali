/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

/** Extracting data from an array of Objects and returning an Object that contains an array of data and the options that will help building the charts.
 * Some functions are using a key, these are templates for mainstream KPIs. */


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
 * Calcul de la réparition des contacts selon la colonne dont le nom est spécifié.
 * @param {string} key Nom de la colonne selon laquelle sera effectuée la répartition.
 * @param {Array} dataIn Données d'entrée.
 * @returns {Array} Table des données permettant de créer le graphe correspondant.
 */
function totalDistribution(key, dataIn, options) {
    let dataOut = [];

    uniqueValues(key, dataIn).forEach(type => {
        // [key] : setting the key of an Object using a variable (only with ES6 and Babel)
        dataOut.push({
            [key]: type,
            "Proportion des contacts": dataIn.filter(row => row[key] == type).length / dataIn.length
        });
    });

    return {
        data: dataOut,
        options: options
    };
}


/**
 * Calcul de la répartition du CA selon les valeurs présentes dans une colonne du fichier de suivi de la prospection (transposable au suivi d'étude en remplaçant caPot en prix).
 * @param {string} key Nom de la colonne selon laquelle sera évaluée la répartition du CA.
 * @param {Array} dataIn Données d'entrée.
 * @returns {Array} Table des données permettant de créer le graphe correspondant.
 */
function turnoverDistribution(key, dataIn, options, price = HEADS.caPot) {
    let dataOut = [],
        ca = dataIn.reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0);

    Logger.log(`CA calculé à partir du suivi de la prospection : ${ca}`);

    uniqueValues(key, dataIn).forEach(currentType => {
        dataOut.push({
            [key]: currentType,
            "Proportion du CA": prcnt(dataIn.filter(row => row[key] == currentType).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0), ca)
        });
    });

    return {
        data: dataOut,
        options: options
    };
}


/**
 * Calcul du CA et du nombre d'études selon les valeurs présentes dans une colonne du fichier de suivi de la prospection (transposable au suivi d'étude en remplaçant caPot en prix).
 * @param {string} key Nom de la colonne selon laquelle sera évaluée la performance de la JE.
 * @param {Array} dataIn Données d'entrée.
 * @returns {Array} Table des données permettant de créer le graphe correspondant
 * et tableau des valeurs à indiquer selon l'axe vertical.
 */
function performance(key, dataIn, options, price = HEADS.caPot) {
    let dataOut = [],
        maxTick = 0;

    uniqueValues(key, dataIn).forEach(currentType => {
        maxTick = Math.max(maxTick, dataIn.filter(row => row[key] == currentType).length, parseInt(dataIn.filter(row => row[key] == currentType).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0) / 1000, 10) + 1);
        dataOut.push({
            [key]: currentType,
            "Nombre d'études": dataIn.filter(row => row[key] == currentType).length,
            "CA (en milliers d'€)": dataIn.filter(row => row[key] == currentType).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0) / 1000
        });
    });

    options["vticks"] = Array.from(Array(maxTick + 1).keys());

    return {
        data: dataOut,
        options: options
    };
}

/**
 * Calcul du nombre d'études selon les valeurs présente dans une colonne du fichier étudié.
 * @param {string} key Nom de la colonne selon laquelle sera compté le nombre d'études.
 * @param {Array} dataIn Données d'entrée.
 * @returns {Array} Table des données permettant de créer le graphe correspondant
 * et tableau des valeurs à indiquer selon l'axe horizontal.
 */
function numberOfMissions(key, dataIn, options) {
    let dataOut = [];

    let lowerBound = Math.min(...dataIn.map(row => parseInt(row[key], 10))),
        higherBound = Math.max(...dataIn.map(row => parseInt(row[key], 10))),
        values = Array.from(Array(higherBound - lowerBound + 1).keys()).map(number => number + lowerBound);

    values.forEach(number => {
        dataOut.push({
            [key]: number,
            "Nombre d'études": dataIn.filter(row => parseInt(row[HEADS.durée], 10) == number).length
        });
    });

    options["hticks"] = values;

    return {
        data: dataOut,
        options: options
    };
}

/**
 * Calcul du taux de conversion décliné selon les différentes catégories présentes dans une colonne spécifiée (uniquement pour le suivi de la prospection).
 * @param {string} key Nom de la colonne selon laquelle sera évalué le taux de conversion.
 * @param {Array} dataIn Données d'entrée.
 * @returns {Array} Table des données permettant de créer le graphe correspondant.
 */
function conversionRate(key, dataIn, options) {
    let dataOut = [];

    uniqueValues(key, dataIn).forEach(function (type) {
        let objFiltered = dataIn.filter(row => row[key] == type),
            conversionRate = prcnt(objFiltered.filter(row => row[HEADS.état] == ETAT_PROSP.etude).length, objFiltered.length);
        dataOut.push({
            [key]: type,
            "Nombre de contacts": objFiltered.length,
            "Taux de conversion global": conversionRate
        });
    });

    return {
        data: dataOut,
        options: options
    };
}

/**
 * Calcul de la répartition du CA selon les valeurs présentes dans une colonne du fichier de suivi de la prospection (transposable au suivi de la prospection en remplaçant prix par caPot).
 * On suppose que la colonne contient des cases à cocher (deux catégories uniquement).
 * @param {string} key Nom de la colonne selon laquelle sera évaluée la répartition du CA.
 * @param {Array} dataIn Données d'entrée.
 * @returns {Array} Table des données permettant de créer le graphe correspondant.
 */
function turnoverDistributionBinary(key, dataIn, options, price = HEADS.prix) {
    let dataOut = [],
        ca = dataIn.reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0);

    Logger.log(`CA calculé à partir du suivi d'études : ${ca}.`);
    Logger.log(`CA alumni : ${dataIn.filter(row => row[key] || false).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0)}.`);

    dataOut.push({
        [`${key}/Hors ${key}`]: key,
        "Proportion du CA": prcnt(dataIn.filter(row => row[key] || false).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0), ca)
    });
    dataOut.push({
        [`${key}/Hors ${key}`]: `Hors ${key}`,
        "Proportion du CA": prcnt(dataIn.filter(row => !(row[key] || false)).reduce((sum, row) => sum += parseInt(row[price], 10) || 0, 0), ca)
    });

    return {
        data: dataOut,
        options: options
    };
}

// Répartition des contacts par mois
function contacts(dataIn, options) {
    let dataOut = [];

    MONTH_LIST.forEach(function (month) {
        let dataRow = {
            "Mois": `${MONTH_NAMES[month.month]} ${month.year}`
        };
        Object.values(ETAT_PROSP).forEach(function (state) {
            // Counting first every contact who find themselves in a more advanced state than state
            let val = dataIn.filter(row => sameMonth(row, month) && Object.values(ETAT_PROSP).indexOf(row[HEADS.état]) >= Object.values(ETAT_PROSP).indexOf(state)).length;
            if (state == ETAT_PROSP.devis || state == ETAT_PROSP.negoc) {
                // In both cases we forgot to count ppl to whom we sent a devis but who didn't convert it into a mission
                val += dataIn.filter(row => sameMonth(row, month) && Object.values(ETAT_PROSP_BIS).includes(row[HEADS.état]) && row[HEADS.devis]).length
            } else if (state == ETAT_PROSP.rdv) {
                // We must also count contacts who didn't lead to a mission
                val += dataIn.filter(row => sameMonth(row, month) && Object.values(ETAT_PROSP_BIS).includes(row[HEADS.état])).length
            }
            dataRow[state] = val;
        });
        dataOut.push(dataRow);
    });

    return {
        data: dataOut,
        options: options
    };
}


// Taux de conversion global
function conversionRateOverTime(dataIn, options) {
    let conversionChart = conversionTable(dataIn),
        dataOut = [];

    // Rows
    MONTH_LIST.forEach(function (month, idx) {
        dataOut.push({
            "Mois": `${MONTH_NAMES[month.month]} ${month.year}`,
            "Taux de conversion global": prcnt(conversionChart[idx][1], conversionChart[idx][0]) / 100 * prcnt(conversionChart[idx][2], conversionChart[idx][1]) / 100 * prcnt(conversionChart[idx][3], conversionChart[idx][2])
        });
    });

    return {
        data: dataOut,
        options: options
    };
}

// Taux de conversion par étape
function conversionRateByType(dataIn, options) {
    let conversionChart = conversionTable(dataIn),
        dataOut = [];

    dataOut.push({
        "Étape de la conversion": `${ETAT_PROSP.rdv} -> ${ETAT_PROSP.devis}`,
        [`Taux de conversion entre ${ETAT_PROSP.rdv} et ${ETAT_PROSP.devis}`]: prcnt(conversionChart.reduce((acc, val) => acc += val[1], 0), conversionChart.reduce((acc, val) => acc += val[0], 0))
    });
    dataOut.push({
        "Étape de la conversion": `${ETAT_PROSP.devis} -> ${ETAT_PROSP.negoc}`,
        [`Taux de conversion entre ${ETAT_PROSP.devis} et ${ETAT_PROSP.negoc}`]: prcnt(conversionChart.reduce((acc, val) => acc += val[2], 0), conversionChart.reduce((acc, val) => acc += val[1], 0))
    });
    dataOut.push({
        "Étape de la conversion": `${ETAT_PROSP.negoc} -> ${ETAT_PROSP.etude}`,
        [`Taux de conversion entre ${ETAT_PROSP.negoc} et ${ETAT_PROSP.etude}`]: prcnt(conversionChart.reduce((acc, val) => acc += val[3], 0), conversionChart.reduce((acc, val) => acc += val[2], 0))
    });

    return {
        data: dataOut,
        options: options
    };
}

// CA potentiel et CA signé
function turnover(dataIn, options) {
    let dataOut = [];

    MONTH_LIST.forEach(function (month) {
        let ca_pot = 0,
            ca_sig = 0;
        dataIn.filter(row => sameMonth(row, month)).forEach(function (row) {
            // /!\ row[HEADS.confiance] isn't actually a percentage since we used method getValues on the range instead of getDisplayValues
            // Potential turnover is the sum of every expected gain
            ca_pot += (parseInt(row[HEADS.caPot], 10) * row[HEADS.confiance]) || 0;
            // Contractualized turnover is computed based on missions who are state ETAT_PROSP.study
            if (row[HEADS.état] == ETAT_PROSP.etude) {
                ca_sig += parseInt(row[HEADS.caPot], 10) || 0;
            }
            Logger.log(row);
        });
        dataOut.push({
            "Mois": `${MONTH_NAMES[month.month]} ${month.year}`,
            "CA sur étude potentielle": ca_pot,
            "CA sur étude signée": ca_sig
        });
    });

    return {
        data: dataOut,
        options: options
    };
}

// Répartition des prix des études
function priceRange(dataIn, options) {
    let dataOut = [],
        width = parseInt((options.higherBound - options.lowerBound) / options.nbrRanges, 10),
        priceRanges = Array.from(Array(options.nbrRanges).keys()).map((_, idx) => options.lowerBound + idx * width);

    priceRanges.forEach(price => {
        dataOut.push({
            "Intervalle de prix": `${price} - ${price + width}`,
            "Nombre d'études": dataIn.filter(row => (price <= row[HEADS.prix] && row[HEADS.prix] < price + width)).length
        });
    });

    delete options.higherBound;
    delete options.lowerBound;
    delete options.nbrRanges;

    return {
        data: dataOut,
        options: options
    };
}

// Proportion du CA venant de la prospection
function prospectionTurnover(dataIn, options) {
    let dataOut = [],
        ca = dataIn.reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0);

    dataOut.push({
        "Prospection ou non": "Prospection",
        "Proportion du CA": prcnt(dataIn.filter(row => !([CONTACT_TYPE.site, CONTACT_TYPE.spontané].includes(row[HEADS.typeContact]))).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0), ca)
    });
    dataOut.push({
        "Prospection ou non": "Hors prospection",
        "Proportion du CA": prcnt(dataIn.filter(row => [CONTACT_TYPE.site, CONTACT_TYPE.spontané].includes(row[HEADS.typeContact])).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0), ca)
    });

    return {
        data: dataOut,
        options: options
    };
}

function performanceByContact(key, dataIn, options) {
    let dataOut = [];

    let maxTick = 0;
    uniqueValues(key, dataIn).forEach(currentType => {
        maxTick = Math.max(
            maxTick,
            dataIn.filter(row => !!row[HEADS.devis] && row[key] == currentType).length,
            dataIn.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).length,
            parseInt(dataIn.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0) / 1000, 10) + 1);

        dataOut.push({
            [key]: currentType,
            "Nombre de devis envoyés": dataIn.filter(row => !!row[HEADS.devis] && row[key] == currentType).length,
            "Nombre d'études signées": dataIn.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).length,
            "CA (en milliers d'euros)": dataIn.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).reduce((sum, row) => sum += parseInt(row[HEADS.caPot], 10) || 0, 0) / 1000
        });
    });

    // Updating the options
    options["vticks"] = Array.from(Array(maxTick + 1).keys());

    return {
        data: dataOut,
        options: options
    };
}

// Nombre de contacts dus au site par mois
function contactBySite(dataIn, options) {
    let dataOut = [],
        dateData = dataIn.map(row => [row["Horodateur"], row["Entreprise"]]);

    MONTH_LIST.forEach(function (month, idx) {
        dataOut.push({
            "Mois": `${MONTH_NAMES[month.month]} ${month.year}`,
            "Nombre de contact par le site": dateData.filter(row => (row[0].getMonth() == month.month && row[0].getFullYear() == month.year && row[1] != "test")).length
        });
    });

    return {
        data: dataOut,
        options: options
    };
}

function moyenneRate(dataIn,options,liste){
    let dataOut = [];
    liste.forEach(question => dataOut.push({
        [key]: question,
        "rate" : dataIn.map(row => row[question]).reduce( (r,x) => r = r + x,0)/dataIn.length
    }
    ));

    return {
        data: dataOut,
        options: options
    };
}

function rateOn5(question,dataIn,options){
    let dataOut = [];
    const numbers = [1,2,3,4,5];
    numbers.forEach( number => dataOut.push({
        [key]: number,
        "occurence" : dataIn.filter(row => (row[question]==key) ).length
    }));

    return {
        data: dataOut,
        options: options
    };
}

function keyNumbers(dataIn, options){
    let dataOut = [];
    let tableSignedEtude = dataIn.filter( row => !(Object.values(ETAT_ETUDE_BIS).includes(row[HEADS.état])) );
    dataOut.push({
        "etude_travaillee" : "Etudes travaillées",
        "Nombre" : dataIn.length
    });
    dataOut.push({
        "etude_signee" : "Etudes signees",
        "Nombre" : tableSignedEtude.length
    });
    dataOut.push({
        "Ca_signe" : "CA signe (en millier d'euros)",
        "Nombre" : tableSignedEtude.map(row => row[HEADS.prix]).reduce( (somme,valeur) => somme = somme + valeur,0)/1000
    });

    return {
        data : dataOut,
        options : options
    };
}

function conversionTotal(dataIn, options){
    let dataOut = [];
    let tableDevisSent = dataIn.filter(row => row[HEADS.état] == ETAT_PROSP.devis || row[HEADS.état] == ETAT_PROSP.negoc ||row[HEADS.état] == ETAT_PROSP.etude),
        tableSignedEtude = dataIn.filter( row => row[HEADS.état] == ETAT_PROSP.etudecl);
    dataOut.push({
        "taux_conversion__devis_etude": `Conversion en étude signée sur ${tableDevisSent.length} devis envoyés`,
        "Nombre" : 100*tableSignedEtude.length/tableDevisSent.length
    });
    dataOut.push({
        "taux_conversion_contact_etude": `Conversion en étude signée sur ${dataIn.length} contacts`,
        "Nombre" : 100*tableSignedEtude.length/dataIn.length
    });
    dataOut.push({
        "taux_conversion_contact_devis": `Conversion en devis envoyé sur ${dataIn.length} contacts`,
        "Nombre" : 100*tableDevisSent.length/dataIn.length
    });
    return {
        data : dataOut,
        options : options
    }
}