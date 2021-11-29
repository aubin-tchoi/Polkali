function rewrite(dataOut, sheetName, spreadsheetId) {
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    let newData = [];
    dataOut.data.forEach(obj => newData.push(Object.values(obj)));
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
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
                        rewrite(KPI.extract(data[KPI.data].filter(KPI.filter || (_ => true)), KPI.options), KPI.name,category.id);
                    })
            }
        });
}

function onOpen(){
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Actu")
    .addItem("Actualiser", "actu")
    .addToUi();
}

function test(){
}