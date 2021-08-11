const IDSSKPI = "16EkXSEQPI7LRb45aSbjZAZFPfP4pmx0__4leyLHsDIc";

function rewrite(data, sheetName, chartType) {
    const sheet = SpreadsheetApp.openById(IDSSKPI).getSheetByName(sheetName),
        lgt = Object.values(data[0]).length;
    let newData = [];
    if (chartType === "column") {
        newData.push(Object.keys(data[0]));
        data.forEach(obj => newData.push(Object.values(obj)));
    } else if (chartType === "pie" || chartType === "line") {
        data.forEach(obj => newData.push(Object.values(obj)));
        sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    }
}

function convert(sheetName, chartType) {
    let sheet = SpreadsheetApp.openById(IDSSKPI).getSheetByName(sheetName);
    let chart = sheet.newChart()
        .addRange(sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()))
        .setPosition(sheet.getLastRow() + 1, sheet.getLastColumn() + 1, 1, 1);
    if (chartType === Charts.ChartType.COLUMN) {
        chart.setChartType(Charts.ChartType.COLUMN)
            .asColumnChart();
    } else if (chartType === Charts.ChartType.PIE) {
        chart.setChartType(Charts.ChartType.PIE)
            .asPieChart();
    } else if (chartType === Charts.ChartType.LINE) {
        chart.setChartType(Charts.ChartType.LINE)
            .asLineChart();
    }
    sheet.insertChart(chart.build());
}

function actu(){
  Object.entries(CATEGORIES).forEach(
    category => {
      if (Object.keys(category[1].KPIs).length > 0) {
        // Loop on each KPI within this category.
        Object.values(category[1].KPIs).forEach(
          KPI => {
            rewrite(KPI.extract(data[KPI.data].filter(KPI.filter || (_ => true)), KPI.options),KPI.name,KPI.chartType);
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