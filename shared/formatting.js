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