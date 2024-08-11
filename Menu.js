// globals.js: ui, isLoaded
// shared/utils.js: getSheetContentHash, shouldRunUpdates
// shared/formatting: applyFormatToSelected, applyFormatToAll
// TODOsheet/TODOformatting.js: applyFormatToAllTODO, customCeilBGColorTODO, createPieChartTODO, deleteAllChartsTODO, updateDateColorsTODO, setupDropdownTODO, pushUpEmptyCellsTODO, updateCellCommentTODO, removeMultipleDatesTODO, updateDaysLeftTODO
// TODOsheet/TODOcheckbox.js: addCheckboxToCellTODO, addCheckboxesToSelectedCellsTODO, markCheckboxSelectedCellsTODO, markAllCheckboxesSelectedCellsTODO, removeCheckboxesFromSelectedCellsTODO

/**
 * Initializes the UI menu in the spreadsheet.
 * Sets up custom menus and triggers functions when menu items are clicked.
 *
 * @customfunction
 */
function onOpen() {
    Logger.log('onOpen triggered');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const docProperties = PropertiesService.getDocumentProperties();
    const language = docProperties.getProperty('language') || 'English';

    saveSnapshotTODO() // point A

    Logger.log('Current language: ' + language);
    try {
        const docProperties = PropertiesService.getDocumentProperties();
        const lastHash = docProperties.getProperty('lastHash');
        const currentHash = getSheetContentHash();
        ss.toast(toastMessages.loading[language], 'Status:', 13);
        applyGridLoaderTODO(sheet)

        if (shouldRunUpdates(lastHash, currentHash)) {
            isLoaded = false;
            Utilities.sleep(2000);
            runAllFunctionsTODO(); // point B
            docProperties.setProperty('lastHash', currentHash);
            Logger.log('Running all update functions');
            isLoaded = true
        } else {
            Logger.log('It is not necessary to run all functions, the data has not changed significantly.');
        }

        if (isLoaded) {
            createMenusTODO();
            translateSheetTODO();
            applyFormatToAllTODO();
            customCellBGColorTODO();
            updateCellCommentTODO();
            ss.toast(toastMessages.updateComplete[language], 'Status:', 5);
        }
    } catch (e) {
        Logger.log('Error: ' + e.toString());
        ui.alert('Error during processing: ' + e.toString());
    }
}

/**
 * Applies a grid loader to the sheet.
 * Adds a red border to the first 20 rows and 8 columns.
 * Used to indicate that the sheet is loading.
 * 
 * @param {Sheet} sheet - The sheet to apply the grid loader to.
 * @customfunction
 */
function applyGridLoaderTODO(sheet) {
    const startRow = 1;
    const endRow = 21;
    const startColumn = 1;
    const endColumn = 8;

    const range = sheet.getRange(startRow, startColumn, endRow, endColumn);
    range.setBorder(true, true, true, true, true, true, '#FF0000', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Runs all functions needed to update the TODO sheet.
 * Calls multiple formatting and update functions.
 *
 * @customfunction
 */
function runAllFunctionsTODO() {
    Logger.log('runAllFunctionsTODO triggered');
    Utilities.sleep(2000);
    updateDateColorsTODO();
    setupDropdownTODO();
    removeMultipleDatesTODO();
    restoreSnapshotTODO(); // point B
    // functions that are meant to run on load
    pushUpEmptyCellsTODO();
    updateDaysLeftCounterTODO();
    Logger.log('All functions called successfully!');
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        onOpen,
        runAllFunctionsTODO,
        applyGridLoaderTODO
    }
}
