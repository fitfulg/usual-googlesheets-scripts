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

    // Check the number of occupied cells in columns C, D, and E
    checkAndSetColumn("C", 10, "ALTA");
    checkAndSetColumn("D", 20, "MEDIA");
    checkAndSetColumn("E", 20, "BAJA");

    // Set the specific content and styles in the specified cells
    setCellContentAndStyle();
}

function checkAndSetColumn(column, limit, priority) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = sheet.getRange(column + "2:" + column + dataRange.getLastRow()).getValues().flat();
    const occupied = values.filter(String).length;

    if (occupied > limit) {
        // Set border color to red
        sheet.getRange(column + "2:" + column + dataRange.getLastRow()).setBorder(true, true, true, true, true, true, "#FF0000", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(column + "1").setValue("⚠️limite de celdas alcanzadas⚠️");
        SpreadsheetApp.getUi().alert("⚠️limite de celdas alcanzadas para la prioridad: " + priority + "⚠️");
    } else {
        // Set border color to black
        sheet.getRange(column + "2:" + column + dataRange.getLastRow()).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(column + "1").setValue("PRIORIDAD " + priority);
    }
}

function setCellContentAndStyle() {
    // Set content and style for A1
    sheet.getRange("A1").setValue("QUICKPATTERNS")
        .setFontWeight("bold")
        .setFontColor("#FFFFFF")
        .setBackground("#000000")
        .setHorizontalAlignment("center");

    // Set content and style for B1
    sheet.getRange("B1").setValue("TOMORROW")
        .setFontWeight("bold")
        .setFontColor("#FFFFFF")
        .setHorizontalAlignment("center");

    // Set content and style for B3
    sheet.getRange("B3").setValue("WEEK")
        .setFontWeight("bold")
        .setFontColor("#FFFFFF")
        .setHorizontalAlignment("center");

    // Set content and style for B8
    sheet.getRange("B8").setValue("MONTH")
        .setFontWeight("bold")
        .setFontColor("#FFFFFF")
        .setHorizontalAlignment("center");

    // Set content and style for F1
    sheet.getRange("F1").setValue("💡IDEAS Y PLANES")
        .setFontWeight("bold")
        .setFontColor("#000000")
        .setBackground("#FFC0CB")
        .setHorizontalAlignment("center");

    // Set content and style for G1
    sheet.getRange("G1").setValue("👀 EYES ON")
        .setFontWeight("bold")
        .setFontColor("#000000")
        .setBackground("#D3D3D3")
        .setHorizontalAlignment("center");

    // Set content and style for H1
    sheet.getRange("H1").setValue("IN QUARANTINE BEFORE BEING CANCELED")
        .setFontWeight("bold")
        .setFontColor("#FF0000")
        .setHorizontalAlignment("center");
}

function setColumnBackground(sheet, col, color, startRow = 2) {
    let lastRow = sheet.getLastRow();
    if (lastRow > 1) { // Ensure there are more than one row
        let range = sheet.getRange(startRow, col, lastRow - startRow + 1, 1);
        range.setBackground(color);
    }
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
        setColumnBackground(sheet, col, '#ffffff');
    }

    // Apply dark yellow background to specific cells in column B
    sheet.getRange('B3').setBackground('#b5a642'); // Dark yellow 3
    sheet.getRange('B8').setBackground('#b5a642'); // Dark yellow 3
}