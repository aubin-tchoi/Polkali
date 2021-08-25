const IDSSKPI = "16EkXSEQPI7LRb45aSbjZAZFPfP4pmx0__4leyLHsDIc";

function rewrite(dataOut, sheetName, chartType) {
    const sheet = SpreadsheetApp.openById(IDSSKPI).getSheetByName(sheetName);
    let newData = [];
    dataOut.data.forEach(obj => newData.push(Object.values(obj)));
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

function convert(sheetName, chartType, options, title) {
    let sheet = SpreadsheetApp.openById(IDSSKPI).getSheetByName(sheetName);
    let chart = sheet.newChart()
        .addRange(sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()))
        .setPosition(sheet.getLastRow() + 1, sheet.getLastColumn() + 1, 1, 1);
    if (chartType === CHART_TYPE.COLUMN) {
        chart.setChartType(Charts.ChartType.COLUMN)
            .asColumnChart();
    } else if (chartType === CHART_TYPE.PIE) {
        chart.setChartType(Charts.ChartType.PIE)
            .asPieChart();
    } else if (chartType === CHART_TYPE.LINE) {
        chart.setChartType(Charts.ChartType.LINE)
            .asLineChart().setOption('curveType', 'function').setOption('pointSize', 10).setOption('pointShape', 'diamond');
    };
    options = {
        ...DEFAULT_PARAMS,
        ...options
    };
    chart.setOption('legend', {
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
        .setOption('title', title);
    try{
      chart.setOption('series', options.series)
    }catch(e){Logger.log(e)
    };
    sheet.getCharts().forEach(chart => { sheet.removeChart(chart); });
    sheet.insertChart(chart.build());
}

function actu() {
    let data = {};
    // Data extracted from each sheet as an Array of Object, each element being 1 line in the Sheet with keys that match the columns given in HEADS.
    Object.entries(DATA_LINKS).forEach(spreadsheet => {
        data[spreadsheet[0]] = Object.values(spreadsheet[1]).reduce((dataTemp, sheet) => dataTemp.concat(extractSheetData(sheet.id, sheet.sheetName, sheet.pos).filter(sheet.filter)), []);
    });
    Object.entries(CATEGORIES).forEach(
        category => {
            if (Object.keys(category[1].KPIs).length > 0) {
                // Loop on each KPI within this category.
                Object.values(category[1].KPIs).forEach(
                    KPI => {
                        rewrite(KPI.extract(data[KPI.data].filter(KPI.filter || (_ => true)), KPI.options), KPI.name, KPI.chartType);
                    })
            }
        });
}

function actu2(){
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Actu")
    .addItem("Actualiser", "actu")
    .addToUi();
}