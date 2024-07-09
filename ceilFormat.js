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
    checkAndSetColumn("C", 10);
    checkAndSetColumn("D", 20);
    checkAndSetColumn("E", 20);
}

function checkAndSetColumn(column, limit) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = sheet.getRange(column + "2:" + column + dataRange.getLastRow()).getValues().flat();
    const occupied = values.filter(String).length;

    if (occupied > limit) {
        // Set border color to red
        sheet.getRange(column + "2:" + column + dataRange.getLastRow()).setBorder(true, true, true, true, true, true, "#FF0000", SpreadsheetApp.BorderStyle.SOLID);
        // Set cell value
        sheet.getRange(column + "1").setValue("⚠️limite de celdas alcanzadas⚠️");
    } else {
        // Set border color to black
        sheet.getRange(column + "2:" + column + dataRange.getLastRow()).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        // Set cell value
        sheet.getRange(column + "1").setValue(column === "C" ? "PRIORIDAD ALTA" : column === "D" ? "PRIORIDAD MEDIA" : "PRIORIDAD BAJA");
    }
}