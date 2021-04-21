/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Creating charts using a data table as input

// Creating a chart using one of the different builders among PieChart, ColumnChart and LineChart
/**
 * Creation d'un graphe au type Chart.
 * @param {Enum} chartType - Type du graphe (Pie, Line, Column).
 * @param {DataTable} dataTable - Table de données à utiliser.
 * @param {string} title - Titre du graphe.
 * @param {Object} options - Options modifiant les propriétés du graphe (dimensions, couleurs, ...).
 * @returns {Chart} - Graphe complet.
 */
function createChart(chartType, dataTable, title, options = {}) {
    Logger.log(`Creating chart ${title} of type ${Object.keys(CHART_TYPE)[Object.values(CHART_TYPE).indexOf(chartType)]} with ${!options ? "no option" : `options : ${Object.entries(options)}`}.`);
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

// Add options to a chartBuilder
/** 
 * Ajouter des options à un chartBuilder (appelée par createChart).
 * @param {ChartBuilder} chartBuilder - ChartBuilder utilisé.
 * @param {string} title - Titre du graphe.
 * @param {Object} options - Options modifiant les propriétés du graphe (dimensions, couleurs, ...).
 */
function addOptions(chartBuilder, title, options) {
    return chartBuilder
        .setOption('legend', {
            textStyle: {
                font: 'roboto',
                fontSize: 11
            }
        })
        .setOption('colors', options.colors || Object.values(COLORS))
        .setOption('vAxis.minValue', (options.percent || false) ? 0 : 'automatic')
        .setOption('vAxis.maxValue', (options.percent || false) ? 100 : 'automatic')
        .setOption('hAxis.ticks', options.hticks || 'auto')
        .setTitle(title)
        .setDimensions(options.width || DIMS.width, options.height || DIMS.height)
        .build();
}
