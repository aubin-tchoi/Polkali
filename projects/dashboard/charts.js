/* Si vous avez des questions à propos de ce script contactez Aubin Tchoï (Directeur Qualité 022) */

// Creating charts using a data table as input

// Creating a ColumnChart
function createColumnChart(dataTable, title, htmlOutput, attachments, width, height) {
    try {
        // Creating a ColumnChart with data from dataTable
        let chart = Charts.newColumnChart()
            .setDataTable(dataTable)
            .setOption('legend', {
                textStyle: {
                    font: 'trebuchet ms',
                    fontSize: 11
                }
            })
            .setOption('colors', Object.values(COLORS))
            .setTitle(title)
            .setDimensions(width, height)
            .build();

        // Adding the chart to the HtmlOutput
        let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
            imageUrl = "data:image/png;base64," + encodeURI(imageData);
        htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

        // Adding the chart to the attachments
        let imageDatamail = chart.getAs('image/png').getBytes(),
            imgblob = Utilities.newBlob(imageDatamail, "image/png", "Contacts prospection");
        attachments.push(imgblob);

        return [htmlOutput, attachments];
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
}

// Creating a PieChart
function createPieChart(dataTable, title, htmlOutput, attachments, width, height) {
    try {
        // Creating a PieChart with data from dataTable
        let chart = Charts.newPieChart()
            .setDataTable(dataTable)
            .setOption('legend', {
                textStyle: {
                    font: 'trebuchet ms',
                    fontSize: 11
                }
            })
            .setOption('colors', Object.values(COLORS))
            .setTitle(title)
            .setDimensions(width, height)
            .build();

        // Adding the chart to the HtmlOutput
        let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
            imageUrl = "data:image/png;base64," + encodeURI(imageData);
        htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

        // Adding the chart to the attachments
        let imageDatamail = chart.getAs('image/png').getBytes(),
            imgblob = Utilities.newBlob(imageDatamail, "image/png", title);
        attachments.push(imgblob);

        return [htmlOutput, attachments];
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title}, error : ${e}`);
    }
}

// Creating a LineChart
function createLineChart(dataTable, colors, title, htmlOutput, attachments, width, height) {
    try {
        // Creating a LineChart with data from dataTable
        let chart = Charts.newLineChart()
            .setDataTable(dataTable)
            .setOption('legend', {
                textStyle: {
                    font: 'trebuchet ms',
                    fontSize: 11
                }
            })
            .setOption('curveType', 'function')
            .setOption('pointShape', 'square')
            .setTitle(title)
            .setDimensions(width, height)
            .setColors(colors)
            .build();

        // Adding the chart to the HtmlOutput
        let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
            imageUrl = "data:image/png;base64," + encodeURI(imageData);
        htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

        // Adding the chart to the attachments
        let imageDatamail = chart.getAs('image/png').getBytes(),
            imgblob = Utilities.newBlob(imageDatamail, "image/png", `KPIs par mois`);
        attachments.push(imgblob);

        return [htmlOutput, attachments];
    } catch (e) {
        Logger.log(`Could not create graph for info : ${title},, error : ${e}`);
    }
}
