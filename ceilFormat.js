function Format(range) {
    // Apply the desired formats
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
}

function applyFormatToSelected() {
    // Get the active sheet and the selected range
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange();

    // Apply formatting to the selected range
    Format(range);
}

function applyFormatToAll() {
    // Get the active sheet and the entire data range
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getDataRange();

    // Apply formatting to the entire data range
    Format(range);
}
