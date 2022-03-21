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
            conversionChartRow.push(val);
        });
        conversionChart.push(conversionChartRow);
    });

    return conversionChart;
}

// Répartition des contacts par mois
function contacts(dataIn, options) {
    let dataOut = [];

    MONTH_LIST.forEach(function (month) {
        let dataRow = {
            "Mois": `${MONTH_NAMES[month.month]} ${month.year}`
        };
        Object.values(ETAT_PROSP_TER).forEach(function (state) {
            // Counting first every contact who find themselves in a more advanced state than state
            let val = dataIn.filter(row => sameMonth(row, month) && Object.values(ETAT_PROSP).indexOf(row[HEADS.état]) >= Object.values(ETAT_PROSP).indexOf(state)).length;
            dataRow[state] = val;
        });
        dataOut.push(dataRow);
    });

    return {
        data: dataOut,
        options: options
    };
}

function conversionTotal(dataIn, options, dataAux){
    let dataOut = [];
    let tableDevisSent = dataIn.filter(row => Object.values(ETAT_PROSP).indexOf(row[HEADS.état]) >= Object.values(ETAT_PROSP).indexOf(ETAT_PROSP.devis)),
        tableSignedEtude = dataAux.filter( row => Object.values(ETAT_ETUDE_SIGNEE).includes(row[HEADS.état]));
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
    };
}

// Taux de conversion global
function conversionRateOverTime(dataIn, options, dataAux) {
    let dataOut = [];
    let dataNew = fusion(HEADS.entreprise,dataIn,dataAux,x => true);

    // Rows
    MONTH_LIST.forEach(function (month) {
        dataOut.push({
            "Mois": `${MONTH_NAMES[month.month]} ${month.year}`,
            "Taux de conversion global": prcnt(dataNew.filter(row => sameMonth(row,month) &&
            (Object.values(ETAT_ETUDE_SIGNEE).includes(row[HEADS.état]))).length, dataIn.filter( row => sameMonth(row,month) && Object.values(ETAT_PROSP).indexOf(row[HEADS.état]) >= Object.values(ETAT_PROSP).indexOf(ETAT_PROSP.devis)).length)
        });
    });

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

function keyNumbers(dataIn, options){
    let dataOut = [];
    let tableSignedEtude = dataIn.filter( row => (Object.values(ETAT_ETUDE_SIGNEE).includes(row[HEADS.état])) );
    dataOut.push({
        "etude_travaillee" : "Etudes travaillées",
        "Nombre" : dataIn.length
    });
    dataOut.push({
        "etude_signee" : "Etudes signées",
        "Nombre" : tableSignedEtude.length
    });
    dataOut.push({
        "Ca_signe" : "CA signé (en millier d'euros)",
        "Nombre" : tableSignedEtude.map(row => row[HEADS.prix]).reduce( (somme,valeur) => somme = somme + valeur,0)/1000
    });

    return {
        data : dataOut,
        options : options
    };
}

// Taux de conversion par étape
function conversionRateByType(dataIn, options) {
    let dataOut = [];
    let nbrdv = dataIn.length;
    Logger.log(nbrdv);
    let nbdevis = dataIn.filter(row => Object.values(ETAT_PROSP).indexOf(ETAT_PROSP.devis) <= Object.values(ETAT_PROSP).indexOf(row[HEADS.état])).length;
    Logger.log(nbdevis);
    let nbnegoc = dataIn.filter(row => Object.values(ETAT_PROSP).indexOf(ETAT_PROSP.negoc) <= Object.values(ETAT_PROSP).indexOf(row[HEADS.état])).length;
    Logger.log(nbnegoc);
    let nbetude = dataIn.filter(row => Object.values(ETAT_PROSP).indexOf(ETAT_PROSP.etude) <= Object.values(ETAT_PROSP).indexOf(row[HEADS.état])).length;
    Logger.log(nbetude);

    dataOut.push({
        "Étape de la conversion": `${ETAT_PROSP.rdv} -> ${ETAT_PROSP.devis}`,
        [`Taux de conversion entre ${ETAT_PROSP.rdv} et ${ETAT_PROSP.devis}`]: prcnt(nbdevis,nbrdv)
    });
    dataOut.push({
        "Étape de la conversion": `${ETAT_PROSP.devis} -> ${ETAT_PROSP.negoc}`,
        [`Taux de conversion entre ${ETAT_PROSP.devis} et ${ETAT_PROSP.negoc}`]: prcnt(nbnegoc,nbdevis)
    });
    dataOut.push({
        "Étape de la conversion": `${ETAT_PROSP.negoc} -> ${ETAT_PROSP.etude}`,
        [`Taux de conversion entre ${ETAT_PROSP.negoc} et ${ETAT_PROSP.etude}`]: prcnt(nbetude,nbnegoc)
    });

    return {
        data: dataOut,
        options: options
    };
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


function performanceByContact(key, dataIn, options, dataHelp) {
    let dataOut = [];
    let filter = etude => (etude[HEADS.état] == ETAT_PROSP.etude)
    let dataNew = fusion(HEADS.entreprise,dataIn,dataHelp,filter);
    let maxTick = 0;
    uniqueValues(key, dataIn).forEach(currentType => {
        maxTick = Math.max(
            maxTick,
            dataIn.filter(row => !!row[HEADS.devis] && row[key] == currentType).length,
            dataNew.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).length,
            parseInt(dataNew.filter(row => row[HEADS.état] == ETAT_PROSP.etude && row[key] == currentType).reduce((sum, row) => sum += parseInt(row[HEADS.prix], 10) || 0, 0) / 1000, 10) + 1);
        dataOut.push({
            [key]: currentType,
            "Nombre de devis envoyés": dataIn.filter(row => (Object.values(ETAT_ETUDE).indexOf(ETAT_ETUDE.devisAccepte) <= Object.values(ETAT_ETUDE).indexOf(row[HEADS.état]) || (Object.values(ETAT_PROSP).indexOf(ETAT_PROSP.devis) <= Object.values(ETAT_PROSP).indexOf(row[HEADS.état]) )) && row[key] == currentType).length,
            "Nombre d'études signées": dataNew.filter(row => (Object.values(ETAT_ETUDE_SIGNEE).includes(row[HEADS.état]) && row[key] == currentType)).length,
            "CA (en milliers d'euros)": dataNew.filter(row => (Object.values(ETAT_ETUDE_SIGNEE).includes(row[HEADS.état]) && row[key] == currentType)).reduce((sum, row) => sum += parseInt(row[HEADS.prix], 10) || 0, 0) / 1000
        });
    });

    // Updating the options
    options["vticks"] = Array.from(Array(maxTick + 1).keys());

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

function prospectionParSecteur(dataIn, options){
    let dataOut = [];
    total = dataIn.filter(row => (row["Mail normal"] == 'quali' || row["Mail normal"] == 'VRAI')).length
    TYPEBDD.forEach(type => dataOut.push({
        [key]: type,
        "rate" : dataIn.filter(row => row["type"] == type && (row["Mail normal"] == 'quali' || row["Mail normal"] == 'VRAI')).length/total
    }))
    return {
        data: dataOut,
        options: options
    }
}

function prospectionTotal(dataIn,options){
    let dataOut = [];
    total = dataIn.filter(row => row["Mail"] != '').length
    nbQuali = dataIn.filter(row => row["Mail normal"] == 'quali').length
    nbQuanti = dataIn.filter(row => row["Mail normal"] == 'VRAI').length
    autre = total - nbQuali - nbQuanti
    dataOut.push({
        [key]: quali,
        "rate" : nbQuali/total
    });
    dataOut.push({
        [key]: quali,
        "rate" : nbQuanti/total
    });
    dataOut.push({
        [key]: quali,
        "rate" : autre/total
    });

    return {
        data: dataOut,
        options: options
    }
}