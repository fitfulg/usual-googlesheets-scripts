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
        // Get values excluding B3 and B8
        values = [
            sheet.getRange(column + "2").getValue(),
            ...sheet.getRange(column + "4:" + column + "7").getValues().flat(),
            ...sheet.getRange(column + "9:" + column + dataRange.getLastRow()).getValues().flat()
        ];
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
    for (const column in exampleTexts) {
        const { text, color } = exampleTexts[column];
        exampleTextTODO(column, text, color);
    }
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
    for (const cell in cellStyles) {
        const { value, fontWeight, fontColor, backgroundColor, alignment } = cellStyles[cell];
        setCellStyle(cell, value, fontWeight, fontColor, backgroundColor, alignment);
    }
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

// Contents of ./TODOsheet/TODOtextObjects.js

const cellStyles = {
    "A1": { value: "QUICK PATTERNS", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#000000", alignment: "center" },
    "B1": { value: "TOMORROW", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#b5a642", alignment: "center" },
    "B3": { value: "WEEK", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#b5a642", alignment: "center" },
    "B8": { value: "MONTH", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#b5a642", alignment: "center" },
    "F1": { value: "💡IDEAS AND PLANS", fontWeight: "bold", fontColor: "#000000", backgroundColor: "#FFC0CB", alignment: "center" },
    "G1": { value: "👀 EYES ON", fontWeight: "bold", fontColor: "#000000", backgroundColor: "#b7b7b7", alignment: "center" },
    "H1": { value: "IN QUARANTINE BEFORE BEING CANCELED", fontWeight: "bold", fontColor: "#FF0000", backgroundColor: null, alignment: "center" },
    "C1": { value: "HIGH PRIORITY", fontWeight: "bold", fontColor: null, backgroundColor: "#fce5cd", alignment: "center" },
    "D1": { value: "MEDIUM PRIORITY", fontWeight: "bold", fontColor: null, backgroundColor: "#fff2cc", alignment: "center" },
    "E1": { value: "LOW PRIORITY", fontWeight: "bold", fontColor: null, backgroundColor: "#d9ead3", alignment: "center" }
};

const exampleTexts = {
    "A": { text: "Example: Do it with fear but do it.", color: "#FFFFFF" },
    "B": { text: "Example: 45min of cardio", color: "#A9A9A9" },
    "C": { text: "Example: Join that gym club", color: "#A9A9A9" },
    "D": { text: "Example: Submit that pending data science task.", color: "#A9A9A9" },
    "E": { text: "Example: Buy a new mattress.", color: "#A9A9A9" },
    "F": { text: "Example: Santiago route.", color: "#A9A9A9" },
    "G": { text: "Example: Change front brake pad at 44500km", color: "#FFFFFF" },
    "H": { text: "Example: Join that Crossfit club", color: "#A9A9A9" },
};

// Contents of ./TODOsheet/TODOtriggers.js

function onEdit(e) {
    const range = e.range;
    const value = e.value;

    // Check if the cell is in the specified columns and contains example text
    const columnLetter = range.getA1Notation().charAt(0);

    if (exampleTexts[columnLetter] && value === exampleTexts[columnLetter].text) {
        // Remove the example text formatting
        range.setFontStyle("normal")
            .setFontColor("#000000"); // Set font color to black or default
    }
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
