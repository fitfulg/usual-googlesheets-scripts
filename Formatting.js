function applyFormatting(range) {
    // Apply the desired formats
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
}

function applyFormat() {
    // Get the active sheet and the selected range
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange();

    // Apply formatting to the selected range
    applyFormatting(range);
}

function applyFormatToAll() {
    // Get the active sheet and the entire data range
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getDataRange();

    // Apply formatting to the entire data range
    applyFormatting(range);
}
