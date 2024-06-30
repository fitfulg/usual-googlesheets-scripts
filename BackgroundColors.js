function setColumnBackground(sheet, col, color) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) { // Ensure there are more than one row
        var range = sheet.getRange(2, col, lastRow - 1, 1);
        range.setBackground(color);
    }
}

function backgroundColorsTODO() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Apply background colors to specific columns
    setColumnBackground(sheet, 1, '#d3d3d3'); // Column A: Light gray 3
    setColumnBackground(sheet, 6, '#fff1f1'); // Column F: Light pink
    setColumnBackground(sheet, 7, '#d3d3d3'); // Column G: Light gray 3
}
