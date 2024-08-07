// globals.js: ui
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
    const docProperties = PropertiesService.getDocumentProperties();
    const language = docProperties.getProperty('language') || 'English';

    saveSnapshotTODO()
    Logger.log('Current language: ' + language);

    ss.toast(toastMessages.loading[language], 'Status:', 13);
    try {
        const docProperties = PropertiesService.getDocumentProperties();
        const lastHash = docProperties.getProperty('lastHash');
        const currentHash = getSheetContentHash();

        if (shouldRunUpdates(lastHash, currentHash)) {
            runAllFunctionsTODO();

            restoreSnapshotTODO();
            updateDaysLeftCounterTODO();
            docProperties.setProperty('lastHash', currentHash);
            Logger.log('Running all update functions');
        } else {
            Logger.log('It is not necessary to run all functions, the data has not changed significantly.');
        }

        createMenusTODO();
        translateSheetTODO();
        ss.toast(toastMessages.updateComplete[language], 'Status:', 5);
    } catch (e) {
        Logger.log('Error: ' + e.toString());
        ui.alert('Error during processing: ' + e.toString());
    }
}

/**
 * Runs all functions needed to update the TODO sheet.
 * Calls multiple formatting and update functions.
 *
 * @customfunction
 */
function runAllFunctionsTODO() {
    Logger.log('runAllFunctionsTODO triggered');
    customCellBGColorTODO();
    applyFormatToAllTODO();
    updateDateColorsTODO();
    setupDropdownTODO();
    pushUpEmptyCellsTODO();
    updateCellCommentTODO();
    removeMultipleDatesTODO();
    updateDaysLeftCounterTODO();
    Logger.log('All functions called successfully!');
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        onOpen,
        runAllFunctionsTODO,
    }
}
