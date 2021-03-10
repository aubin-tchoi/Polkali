// Creating a PieChart
function createPieChart(question, data, htmlOutput, attachments) {
    try {
        // Creating a DataTable with the proportion of responses for each unique response to question
        let dataTable = Charts.newDataTable();
        dataTable.addColumn(Charts.ColumnType.STRING, question);
        dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion");

        uniqueValues(question, data).forEach(function (val) {
            dataTable.addRow([val.toString(), data.filter(r => r[question] == val).length / data.length]);
        });

        // Creating a PieChart with data from dataTable
        let chart = Charts.newPieChart()
            .setDataTable(dataTable)
            .setOption('legend', {
                textStyle: {
                    font: 'trebuchet ms',
                    fontSize: 11
                }
            })
            .setTitle(question)
            .setDimensions(DIMS["width"], DIMS["height"])
            .set3D();

        chart = chart.build();

        // Adding the chart to the Html output
        let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
            imageUrl = "data:image/png;base64," + encodeURI(imageData);
        htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

        // Adding the chart to the attachments
        let imageDatamail = chart.getAs('image/png').getBytes(),
            imgblob = Utilities.newBlob(imageDatamail, "image/png", question);
        attachments.push(imgblob);

        return [htmlOutput, attachments];
    } catch (e) {
        Logger.log(`Could not create graph for question : ${question}, ${e}`);
    }
}

// Creating a ColumnChart
function createColumnChart(question, data, htmlOutput, attachments) {
    try {
        // Creating a DataTable with the proportion of responses for each unique response to question
        let dataTable = Charts.newDataTable();
        dataTable.addColumn(Charts.ColumnType.STRING, question);
        dataTable.addColumn(Charts.ColumnType.NUMBER, "Proportion");

        uniqueValues(question, data).sort().forEach(function (val) {
            dataTable.addRow([val, data.filter(r => r[question] == val).length / data.length]);
        });

        // Creating a ColumnChart with data from dataTable
        let chart = Charts.newColumnChart()
            .setDataTable(dataTable)
            .setOption('legend', {
                textStyle: {
                    font: 'trebuchet ms',
                    fontSize: 11
                }
            })
            .setTitle(question)
            .setDimensions(DIMS["width"], DIMS["height"])
            .build();

        // Adding the chart to the Html output
        let imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes()),
            imageUrl = "data:image/png;base64," + encodeURI(imageData);
        htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");

        // Adding the chart to the attachments
        let imageDatamail = chart.getAs('image/png').getBytes(),
            imgblob = Utilities.newBlob(imageDatamail, "image/png", question);
        attachments.push(imgblob);

        return [htmlOutput, attachments];
    } catch (e) {
        Logger.log(`Could not create graph for question : ${question}, ${e}`);
    }
}