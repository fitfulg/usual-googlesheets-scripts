const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

function Format(range) {
    // Apply the desired formats
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
}

function applyBorders(range) {
    // Apply black borders with the thinnest line
    range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
}

function applyFormatToSelected() {
    // Get the active sheet and the selected range
    let range = sheet.getActiveRange();
    // Apply formatting to the selected range
    Format(range);
    applyBorders(range)
}

function applyFormatToAll() {
    // Get the active sheet and the entire data range
    let range = sheet.getDataRange();
    // Apply formatting to the entire data range
    Format(range);
    applyBorders(range)
}
