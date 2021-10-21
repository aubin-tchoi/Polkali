/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

/** Various tools. */

// ----- Data extraction tools -----

// I chose to keep the data as an array of objects (other possibility : object of objects, the keys being the months' names and the values the data in each row)
/**
 * Extrait les données d'un sheet sous forme d'un tableau d'Objects afin d'en faciliter la lecture.
 * @param {string} spreadsheetId ID du spreadsheet.
 * @param {string} sheetName Nom de l'onglet à extraire.
 * @param {Object} pos Object contenant deux clés : data et pos pour indiquer la position du header ainsi que des données. 
 * @returns {Array} Array d'Object décrivant l'ensemble des données extraites (1 ligne correspond à une ligne du fichier et les clés sont alors les valeurs du header).
 */
function extractSheetData(spreadsheetId, sheetName, pos) {
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName),
        data = sheet.getRange(pos.data.x, pos.data.y, manuallyGetRowNumber(sheet, pos.data.x, pos.trustColumn), (sheet.getLastColumn() - pos.data.y + 1)).getValues(),
        heads = sheet.getRange(pos.header.x, pos.header.y, 1, (sheet.getLastColumn() - pos.header.y + 1)).getValues().shift();
    return data.map(row => heads.reduce((obj, key, idx) => (obj[key] = (row[idx] != "") ? row[idx] : obj[key] || '', obj), {}));
}

/**
 * Renvoie le nombre de lignes non-vide dans un Sheet (pas toujours possible avec getLastRow à cause de la validation des données).
 * @param {Sheet} sheet Onglet du sheet à considérer.
 * @param {number} start Indice de la 1ère ligne contenant des données.
 * @param {number} trustColumn Indice d'une colonne qui sera vide ssi la ligne est vide.
 * @returns {number} Nombre de lignes de données.
 */
function manuallyGetRowNumber(sheet, start, trustColumn) {
    // 100 is considered as a higher bound for the number of rows inside of this sheet
    let data = sheet.getRange(start, trustColumn, (100 - start + 1), 1).getValues();
    for (row = 0; row < 100; row++) {
        if ((data[row][0]) == "") {
            return row;
        }
    }
}


// ----- Computation tools -----

/**
 * Renvoie un tableau des valeurs uniques trouvées pour une certaine colonne des données entrées.
 * @param {string} key Clé à considérer (doit être présente dans les éléments de data).
 * @param {Array} data Données à considérer.
 * @returns {Array} Valeurs uniques se trouvant dans cette colonne.
 */
const uniqueValues = (key, data) => data.map((row) => row[key]).filter((val, idx, arr) => arr.indexOf(val) == idx);

/**
 * Renvoie a / b * 100 si b est non nul, 0 sinon.
 * @param {number} a Numérateur.
 * @param {number} b Dénominateur.
 * @returns {number} a / b * 100.
 */
const prcnt = (a, b) => (parseInt(b, 10) == 0) ? 0 : a / b * 100;

/**
 * Test vérifiant si une ligne de données correspond à un mois donné.
 * @param {Object} row Ligne de donnée à considérer.
 * @param {Object} month Object à deux éléments: month et year.
 * @returns {Boolean} true ssi le champ "Premier contact" de la ligne correspond bien au mois spécifié.
 */
const sameMonth = (row, month) => parseInt(row[HEADS.premierContact].getMonth(), 10) == month.month && parseInt(row[HEADS.premierContact].getFullYear(), 10) == month.year;

/**
 * Conversion d'un Array d'Objects en DataTable.
 * @param {Array} data Données à parser.
 * @returns {DataTable} DataTable représentant les données d'entrées.
 */
function arrayToDataTable(data) {
    let dataTable = Charts.newDataTable();

    // Columns
    Object.keys(data[0]).forEach(key => {
        dataTable.addColumn(typeof data[0][key] == "number" ? Charts.ColumnType.NUMBER : Charts.ColumnType.STRING, key);
    });

    // Rows
    data.forEach(dataRow => {
        dataTable.addRow(Object.values(dataRow));
    });

    return dataTable;
}

/**
 * Conversion d'un graphe de type Chart en une image qui sera ajoutée à un HtmlOutput ainsi qu'à une liste.
 * @param {Chart} chart Graphe.
 * @param {string} title Titre du graphe.
 * @param {HtmlOutput} htmlOutput Contenu html qui servira à afficher les graphes à l'écran.
 * @param {Array} attachments Liste des blobs des images.
 */
function convertChart(chart, title, htmlOutput, attachments) {
    // Loading the blob for this chart (this part takes the most time)
    let chartBlob = chart.getAs('image/png');
    // Adding the chart to the HtmlOutput
    htmlOutput.append(`<img border="1" src="data:image/png;base64,${encodeURI(Utilities.base64Encode(chartBlob.getBytes()))}">`);
    // Adding the chart to the attachments
    attachments.push(chartBlob.setName(title));
}


// ----- User's side features -----

/**
 * Affichage d'un écran de chargement.
 * @param {string} msg Message à afficher en titre. 
 */
function displayLoadingScreen(msg) {
    try {
        ui.showModelessDialog(HTML_CONTENT.loadingScreen, msg);
    } catch (e) {
        Logger.log(`Couldn't display loading screen with message : ${msg}, error : ${e}`);
    }
}

/**
 * Récupère une entrée de l'utilisateur en ouvrant une fenêtre.
 * Si l'argument message contient une clé 'incorrectInput' l'entrée sera testée et demandée à nouveau jusqu'à ce qu'elle soit validée.
 * @param {Object} message Object décrivant les messages à afficher: titre de la fenêtre, question à poser et message indiquant une entrée non valide (ce dernier est optionnel).
 * @param {function} test Fonction booléenne permettant de tester l'entrée.
 * @returns {string} Réponse de l'utilisateur.
 */
const userQuery = (message, test = (input) => {
    DriveApp.getFolderById(input);
}) => {
    let response = "";
    if (message.incorrectInput != undefined) {
        while (true) {
            try {
                response = ui.prompt(message.title, message.query, ui.ButtonSet.OK).getResponseText();
                // Checking whether the input is valid or not
                test(response);
                return response;
            } catch (e) {
                ui.alert(message.title, message.incorrectInput, ui.ButtonSet.OK);
            }
        }
    } else {
        return ui.prompt(message.title, message.query, ui.ButtonSet.OK).getResponseText();
    }
}

/**
 * Ajout d'une ligne (dont le contenu est hardcodé et correspond à ajouter une ligne de titre).
 * @param {string} key Clé permettant de retrouver le titre de la ligne à ajouter.
 * @param {HtmlOutput} htmlOutput Contenu html à modifier.
 */
const addHTMLLine = (key, htmlOutput) => {
    htmlOutput.append(`<br/><br/><br/><br/><span style='font-size: 16pt;'> <span style="font-family: 'roboto', sans-serif;">${CATEGORIES[key].slideTitle}<br/></span> </span> <br/>`)
};

/**
 * Log le temps écoulé depuis le checkpoint puis renvoit un nouveau checkpoint.
 * @param {Date} checkpoint Date à partir de laquelle le décompte sera fait.
 * @param {string} message Message à afficher dans le log (il s'agit seulement de la fin du message, le début est hardcodé ici).
 * @returns {Date} Nouveau checkpoint.
 */
const measureTime = (checkpoint, message) => {
    let newCheckpoint = new Date();
    Logger.log(`It took ${(newCheckpoint.getTime() - checkpoint.getTime()) / 1000} s to ${message}.`);
    return newCheckpoint;
};

// There will be a time when this decorator will be put to a good use
const logTime = (message) => ((target, name, descriptor) => {
    let initialTime = new Date();
    descriptor.value();
    measureTime(initialTime, message);
});
