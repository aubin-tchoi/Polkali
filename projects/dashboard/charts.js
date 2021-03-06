/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

/** Creating charts using a DataTable as input. */

/**
 * Creation d'un graphe au type Chart.
 * @param {Enum} chartType - Type du graphe (Pie, Line, Column).
 * @param {Array} data - Données à utiliser sous forme d'array d'objects.
 * @param {string} title - Titre du graphe.
 * @param {Object} options - Options modifiant les propriétés du graphe (dimensions, couleurs, ...).
 * @returns {Chart} - Graphe complet.
 */
function createChart(chartType, data, title, options = {}) {
    Logger.log(`Creating chart ${title} of type ${Object.keys(CHART_TYPE)[Object.values(CHART_TYPE).indexOf(chartType)]} with ${!options ? "no option" : `options : ${Object.entries(options)}`}.`);
    let dataTable = arrayToDataTable(data);

    try {
        if (chartType === CHART_TYPE.COLUMN) {
            return addOptions(Charts.newColumnChart().setDataTable(dataTable), title, options);
        } else if (chartType === CHART_TYPE.PIE) {
            return addOptions(Charts.newPieChart().setDataTable(dataTable), title, options);
        } else if (chartType === CHART_TYPE.LINE) {
            return addOptions(Charts.newLineChart().setDataTable(dataTable).setOption('curveType', 'function').setOption('pointSize', 10).setOption('pointShape', 'diamond'), title, options);
        }
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
}

/** 
 * Créer des charts incrustés aux sheets.
 * @param {Enum} chartType - Type du graphe (Pie, Line, Column).
 * @param {DataTable} dataTable - Table de données à utiliser.
 * @param {string} title - Titre du graphe.
 * @param {Object} options - Options modifiant les propriétés du graphe (dimensions, couleurs, ...).
 * @param {string} idSpreadsheet - ID du spreadsheet où on stocke les graphes.
 * @param {string} nameSheet - Nom du sheet où on place les graphes.
 */

function createEmbeddedChart(chartType, dataTable, title, options = {}, idSpreadsheet, nameSheet) {
    var sheet = SpreadsheetApp.openById().getSheetByName();
    var chart = sheet.newChart()
    try {
        if (chartType === CHART_TYPE.COLUMN) {
            chart.setChartType(Charts.ChartType.chartType)
                .asColumnChart();
        } else if (chartType === CHART_TYPE.PIE) {
            chart.setChartType(Charts.ChartType.chartType)
                .asPieChart();
        } else if (chartType === CHART_TYPE.LINE) {
            chart.setChartType(Charts.ChartType.chartType)
                .asColumnChart();
        }
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
    //Pour créer un embedded Chart il faut ensuite créer un tableau dans le sheet destination qui est synchro avec le vrai tableau pour pouvoir add la range.
    chart.sheet.addRange()
        .setOption('legend', {
            textStyle: {
                font: 'roboto',
                fontSize: 11
            }
        })
        .setOption('colors', options.colors || Object.values(COLORS))
        .setOption('vAxis.minValue', (options.percent) ? 0 : 'automatic')
        .setOption('vAxis.maxValue', (options.percent) ? 100 : 'automatic')
        .setOption('hAxis.ticks', options.hticks)
        .setOption('vAxis.ticks', options.vticks)
        .setTitle(title)
        .setDimensions(options.width, options.height)
        .build();
    sheet.insertChart(chart);
}

/** 
 * Ajouter des options à un chartBuilder (appelée par createChart).
 * @param {ChartBuilder} chartBuilder - ChartBuilder utilisé.
 * @param {string} title - Titre du graphe.
 * @param {Object} options - Options modifiant les propriétés du graphe (dimensions, couleurs, ...).
 */
function addOptions(chartBuilder, title, options) {
    // Adding default values
    options = {
        ...DEFAULT_PARAMS,
        ...options
    };

    return chartBuilder
        .setOption('legend', {
            textStyle: {
                font: 'roboto',
                fontSize: 11
            }
        })
        .setOption('colors', options.colors)
        .setOption('vAxis.minValue', (options.percent) ? 0 : 'automatic')
        .setOption('vAxis.maxValue', (options.percent) ? 100 : 'automatic')
        .setOption('hAxis.ticks', options.hticks)
        .setOption('vAxis.ticks', options.vticks)
        .setTitle(title)
        .setDimensions(options.width, options.height)
        .build();
}
