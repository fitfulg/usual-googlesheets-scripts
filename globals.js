

const ui = SpreadsheetApp.getUi();
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const getDataRange = () => sheet.getDataRange();
const datePattern = /\n\d{2}\/\d{2}\/\d{2}$/; // dd/MM/yy

// state management
let isPieChartVisible = false;
let isLoaded = true;

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        ui,
        sheet,
        getDataRange,
        datePattern,
        isPieChartVisible,
        isLoaded,
    }
}