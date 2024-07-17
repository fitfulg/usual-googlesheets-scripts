// globals.js: sheet, getDataRange

function checkAndSetColumn(column, limit, priority) {
    const dataRange = getDataRange();
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
        sheet.getRange(column + "1").setValue(priority);
    }
}
