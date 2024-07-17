// globals.js: sheet, getDataRange

function createPieChartTODO() {
    const dataRange = getDataRange();
    const valuesC = sheet.getRange("C2:C" + dataRange.getLastRow()).getValues().flat();
    const valuesD = sheet.getRange("D2:D" + dataRange.getLastRow()).getValues().flat();
    const valuesE = sheet.getRange("E2:E" + dataRange.getLastRow()).getValues().flat();

    const occupiedC = valuesC.filter(String).length;
    const occupiedD = valuesD.filter(String).length;
    const occupiedE = valuesE.filter(String).length;

    const chartDataRange = sheet.getRange("J1:K4");
    chartDataRange.setValues([
        ["Column", "Occupied Cells"],
        ["HIGH", occupiedC],
        ["MEDIUM", occupiedD],
        ["LOW", occupiedE]
    ]);

    const chart = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartDataRange)
        .setPosition(1, 10, 0, 0) // Position the chart starting at column J
        .build();

    sheet.insertChart(chart);
}
