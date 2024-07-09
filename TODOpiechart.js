function createPieChart() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const valuesB = sheet.getRange("B2:B" + dataRange.getLastRow()).getValues().flat();
    const valuesC = sheet.getRange("C2:C" + dataRange.getLastRow()).getValues().flat();
    const valuesE = sheet.getRange("E2:E" + dataRange.getLastRow()).getValues().flat();

    const occupiedB = valuesB.filter(String).length;
    const occupiedC = valuesC.filter(String).length;
    const occupiedE = valuesE.filter(String).length;

    const chartDataRange = sheet.getRange("J1:K4");
    chartDataRange.setValues([
        ["Column", "Occupied Cells"],
        ["B", occupiedB],
        ["C", occupiedC],
        ["E", occupiedE]
    ]);

    const chart = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartDataRange)
        .setPosition(1, 10, 0, 0) // Position the chart starting at column J
        .build();

    sheet.insertChart(chart);
}
