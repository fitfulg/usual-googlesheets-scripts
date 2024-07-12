// Auto-generated file with all JS scripts

// Contents of ./Menu.js

function onOpen() {
    let ui = SpreadsheetApp.getUi();
    // Add a custom menu
    ui.createMenu('Custom Formats')
        .addItem('Apply Format', 'applyFormatToSelected')
        .addItem('Apply Format to All', 'applyFormatToAll')
        .addItem('TODOsheet: Apply Format to All', 'applyFormatToAllTODO')
        .addItem('TODOsheet: Set Ceil Background Colors', 'customCeilBGColorTODO')
        .addItem('TODOsheet: Create Pie Chart', 'createPieChartTODO')
        .addItem('Log Hello World', 'logHelloWorld')
        .addToUi();

    // Call function when the document is opened or refreshed
    // showLoading();
    createPieChartTODO();
    customCeilBGColorTODO();
    applyFormatToAllTODO();
    // hideLoading();
}

function logHelloWorld() {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Hello, World from Github to GoogleSheets!");
}
// IDEA to implement :
// function showLoading() {
//     SpreadsheetApp.getActiveSpreadsheet().toast('Loading, please wait...', 'Loading', -1);
// }

// function hideLoading() {
//     SpreadsheetApp.getUi().alert('Loading complete!');
// }

// Contents of ./shared/formatting.js

const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

function Format(range) {
    if (range) {
        // Apply the desired formats
        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        range.setHorizontalAlignment("center");
        range.setVerticalAlignment("middle");
    }
}

function applyBorders(range) {
    if (range) {
        // Apply black borders with the thinnest line
        range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    }
}

function applyThickBorders(range) {
    if (range) {
        // Apply black borders with thicker lines
        range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
}

function applyFormatToSelected() {
    // Get the active sheet and the selected range
    let range = sheet.getActiveRange();
    if (range) {
        // Apply formatting to the selected range
        Format(range);
        applyBorders(range);
    }
}

function applyFormatToAll() {
    // Get the active sheet and the entire data range
    let range = sheet.getDataRange();
    if (range) {
        // Apply formatting to the entire data range
        Format(range);
        applyBorders(range);
    }
}

function setCellStyle(cell, value, fontWeight, fontColor, backgroundColor, alignment) {
    let range = sheet.getRange(cell);
    range.setValue(value)
        .setFontWeight(fontWeight)
        .setFontColor(fontColor)
        .setHorizontalAlignment(alignment);

    if (backgroundColor) {
        range.setBackground(backgroundColor);
    }
}
// Contents of ./TODOsheet/TODOformatting.js

function exampleTextTODO(column, exampleText, fontColor = "#A9A9A9") {
    const dataRange = sheet.getDataRange();
    let values;
    if (column === "B") {
        // Exclude cells B3 and B8
        values = sheet.getRange(column + "2:" + column + "2").getValues().flat().concat(
            sheet.getRange(column + "4:" + column + "7").getValues().flat(),
            sheet.getRange(column + "9:" + column + dataRange.getLastRow()).getValues().flat()
        );
    } else {
        values = sheet.getRange(column + "2:" + column + dataRange.getLastRow()).getValues().flat();
    }

    const isEmpty = values.every(value => !value.trim());

    if (isEmpty) {
        const cell = sheet.getRange(column + "2");
        cell.setValue(exampleText)
            .setFontStyle("italic")
            .setFontColor(fontColor); // Custom font color
    }
}

function applyFormatToAllTODO() {
    // Get the active sheet and the entire data range up to row 20 and column I (9)
    let range = sheet.getRange(1, 1, 20, 9); // A1:I20
    if (range) {
        Format(range);
        applyBorders(range);
    }
    // Apply thicker borders to specific columns
    applyThickBorders(sheet.getRange(1, 3, 11, 1)); // C1:C11
    applyThickBorders(sheet.getRange(1, 4, 21, 1)); // D1:D21
    applyThickBorders(sheet.getRange(1, 5, 21, 1)); // E1:E21

    // Set the specific content and styles in the specified cells
    setCellContentAndStyleTODO();

    // Check the number of occupied cells in columns C, D, and E
    checkAndSetColumn("C", 10, "HIGH PRIORITY");
    checkAndSetColumn("D", 20, "MEDIUM PRIORITY");
    checkAndSetColumn("E", 20, "LOW PRIORITY");

    // Add example text to specific columns if empty
    exampleTextTODO("A", "Example: Do it with fear but do it.");
    exampleTextTODO("B", "Example: 45min of cardio");
    exampleTextTODO("C", "Example: Join that gym club");
    exampleTextTODO("D", "Example: Submit that pending data science task.");
    exampleTextTODO("E", "Example: Buy a new mattress.");
    exampleTextTODO("F", "Example: Santiago route.");
    exampleTextTODO("G", "Example: Change front brake pad at 44500km");
    exampleTextTODO("H", "Example: Join that Crossfit club");
}

function setColumnBackground(sheet, col, color, startRow = 2, endRow = 20) {
    let range = sheet.getRange(startRow, col, endRow - startRow + 1, 1);
    range.setBackground(color);
}

function customCeilBGColorTODO() {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Apply background colors to specific columns
    setColumnBackground(sheet, 1, '#d3d3d3'); // Column A: Light gray 3
    setColumnBackground(sheet, 6, '#fff1f1'); // Column F: Light pink
    setColumnBackground(sheet, 7, '#d3d3d3'); // Column G: Light gray 3

    // Apply white background to columns B, C, D, E, H, I starting from row 2
    let whiteColumns = [2, 3, 4, 5, 8, 9]; // Columns B, C, D, E, H, I
    for (let col of whiteColumns) {
        setColumnBackground(sheet, col, '#ffffff', 2, 20);
    }

    // Apply dark yellow background to specific cells in column B
    sheet.getRange('B3').setBackground('#b5a642'); // Dark yellow 3
    sheet.getRange('B8').setBackground('#b5a642'); // Dark yellow 3
}


function setCellContentAndStyleTODO() {
    setCellStyle("A1", "QUICK PATTERNS", "bold", "#FFFFFF", "#000000", "center");
    setCellStyle("B1", "TOMORROW", "bold", "#FFFFFF", "#b5a642", "center");
    setCellStyle("B3", "WEEK", "bold", "#FFFFFF", "#b5a642", "center");
    setCellStyle("B8", "MONTH", "bold", "#FFFFFF", "#b5a642", "center");
    setCellStyle("F1", "💡IDEAS AND PLANS", "bold", "#000000", "#FFC0CB", "center");
    setCellStyle("G1", "👀 EYES ON", "bold", "#000000", "#b7b7b7", "center");
    setCellStyle("H1", "IN QUARANTINE BEFORE BEING CANCELED", "bold", "#FF0000", null, "center");
    setCellStyle("C1", "HIGH PRIORITY", "bold", null, "#fce5cd", "center");
    setCellStyle("D1", "MEDIUM PRIORITY", "bold", null, "#fff2cc", "center");
    setCellStyle("E1", "LOW PRIORITY", "bold", null, "#d9ead3", "center");
}

// Contents of ./TODOsheet/TODOpiechart.js

function createPieChartTODO() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
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

// Contents of ./TODOsheet/TODOvalidation.js

function checkAndSetColumn(column, limit, priority) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = sheet.getRange(column + "2:" + column + dataRange.getLastRow()).getValues().flat();
    const occupied = values.filter(String).length;
    const range = sheet.getRange(column + "2:" + column + dataRange.getLastRow());

    if (occupied > limit) {
        // Set border color to red with thicker border
        range.setBorder(true, true, true, true, true, true, "#FF0000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(column + "1").setValue("⚠️CELL LIMIT REACHED⚠️");
        SpreadsheetApp.getUi().alert("⚠️CELL LIMIT REACHED⚠️ \nfor priority: " + priority);
    } else {
        // Set border color to black with thicker border
        range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(column + "1").setValue("PRIORITY " + priority);
    }
}
