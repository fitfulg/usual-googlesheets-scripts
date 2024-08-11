

const ui = SpreadsheetApp.getUi();
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const getDataRange = () => sheet.getDataRange();
const datePattern = /\d{2}\/\d{2}\/\d{2}$/;

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