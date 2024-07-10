function checkAndSetColumn(column, limit, priority) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = sheet.getRange(column + "2:" + column + dataRange.getLastRow()).getValues().flat();
    const occupied = values.filter(String).length;

    if (occupied > limit) {
        // Set border color to red
        sheet.getRange(column + "2:" + column + dataRange.getLastRow()).setBorder(true, true, true, true, true, true, "#FF0000", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(column + "1").setValue("⚠️LIMITE DE CELDAS ALCANCADAS⚠️");
        SpreadsheetApp.getUi().alert("⚠️LIMITE DE CELDAS ALCANCADAS⚠️ \npara la prioridad: " + priority);
    } else {
        // Set border color to black
        sheet.getRange(column + "2:" + column + dataRange.getLastRow()).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(column + "1").setValue("PRIORIDAD " + priority);
    }
}
