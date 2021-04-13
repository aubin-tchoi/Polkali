/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Creating charts using a data table as input

// Creating a ColumnChart
function createColumnChart(dataTable, colors, title, dims, percent = false) {
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
            .setOption('colors', colors)
            .setOption('vAxis.minValue', percent ? 0 : 'automatic')
            .setOption('vAxis.maxValue', percent ? 100 : 'automatic')
            .setTitle(title)
            .setDimensions(dims.width, dims.height)
            .build();

        return chart;
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
}

// Creating a PieChart
function createPieChart(dataTable, colors, title, dims) {
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
            .setOption('colors', colors)
            .setTitle(title)
            .setDimensions(dims.width, dims.height)
            .build();

        return chart;
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
}

// Creating a LineChart
function createLineChart(dataTable, colors, title, dims) {
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
            .setOption('curveType', 'function')
            .setOption('pointShape', 'square')
            .setTitle(title)
            .setDimensions(dims.width, dims.height)
            .setColors(colors)
            .build();

        return chart;
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title},, error : ${e}`);
    }
}
