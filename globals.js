const ui = SpreadsheetApp.getUi();
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const datePattern = /\n\d{2}\/\d{2}\/\d{2}$/; // dd/MM/yy
const dataRange = sheet.getDataRange();