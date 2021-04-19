/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Creating charts using a data table as input

// Creating a chart using one of the functions described below
function createChart(chartType, dataTable, title, options = {}) {
    if (chartType === CHART_TYPE.COLUMN) {
        return createColumnChart(dataTable, title, options);
    }
    else if (chartType === CHART_TYPE.PIE) {
        return createPieChart(dataTable, title, options);
    }
    else if (chartType === CHART_TYPE.LINE) {
        return createLineChart(dataTable, title, options);
    }
}

// Creating a ColumnChart
function createColumnChart(dataTable, title, options) {
    Logger.log(`Titre : ${title}, options : ${Object.entries(options)}`);
    try {
        // Creating a ColumnChart with data from dataTable
        let chart = Charts.newColumnChart()
            .setDataTable(dataTable)
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

        return chart;
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
}

// Creating a PieChart
function createPieChart(dataTable, title, options) {
    Logger.log(`Titre : ${title}`);
    try {
        // Creating a PieChart with data from dataTable
        let chart = Charts.newPieChart()
            .setDataTable(dataTable)
            .setOption('legend', {
                textStyle: {
                    font: 'roboto',
                    fontSize: 11
                }
            })
            .setOption('colors', options.colors || Object.values(COLORS))
            .setTitle(title)
            .setDimensions(options.width || DIMS.width, options.height || DIMS.height)
            .build();

        return chart;
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
}

// Creating a LineChart
function createLineChart(dataTable, title, options) {
    Logger.log(`Titre : ${title}`);
    try {
        // Creating a LineChart with data from dataTable
        let chart = Charts.newLineChart()
            .setDataTable(dataTable)
            .setOption('legend', {
                textStyle: {
                    font: 'roboto',
                    fontSize: 11
                }
            })
            .setOption('colors', options.colors || Object.values(COLORS))
            .setOption('curveType', 'function')
            .setOption('pointShape', 'square')
            .setTitle(title)
            .setDimensions(options.width || DIMS.width, options.height || DIMS.height)
            .build();

        return chart;
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title},, error : ${e}`);
    }
}
