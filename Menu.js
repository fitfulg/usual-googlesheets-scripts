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

    // bad practice but only way (by the moment) to not lose links from shifted up cells after reloading the page  
    saveSnapshotTODO()

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

    const ui = SpreadsheetApp.getUi();

    // Custom menu
    ui.createMenu('TODO sheet')
        .addItem('RESTORE DEFAULT TODO TEMPLATE', 'applyFormatToAllTODO')
        .addItem('RESTORE Ceil Background Colors', 'customCeilBGColorTODO')
        .addSeparator()
        .addItem('Add Checkboxes to Selected Cells', 'addCheckboxesTODO')
        .addItem('Mark Checkbox in Selected Cells', 'markCheckboxTODO')
        .addItem('Mark All Checkboxes in Selected Cells', 'markAllCheckboxesTODO')
        .addItem('Restore Checkboxes', 'restoreCheckboxesTODO')
        .addItem('Remove All Checkboxes in Selected Cells', 'removeCheckboxesTODO')
        .addSeparator()
        .addItem('Save Snapshot', 'saveSnapshot')
        .addItem('Restore Snapshot', 'restoreSnapshot')
        .addSeparator()
        .addItem('Create Pie Chart', 'createPieChartTODO')
        .addItem('Delete Pie Charts', 'deleteAllChartsTODO')
        .addSeparator()
        .addItem('Version and feature details', 'updateCellCommentTODO')
        .addSeparator()
        .addItem('Log Hello World', 'logHelloWorld')
        .addToUi();

    ui.createMenu('Custom Formats')
        .addItem('Apply Format', 'applyFormatToSelected')
        .addItem('Apply Format to All', 'applyFormatToAll')
        .addToUi();

    ui.createMenu('Language')
        .addItem('English', 'setLanguageEnglish')
        .addItem('Spanish', 'setLanguageSpanish')
        .addItem('Catalan', 'setLanguageCatalan')
        .addToUi();
    translateSheetTODO();
}

/**
 * Runs all functions needed to update the TODO sheet.
 * Calls multiple formatting and update functions.
 *
 * @customfunction
 */
function runAllFunctionsTODO() {
    customCeilBGColorTODO();
    applyFormatToAllTODO();
    updateDateColorsTODO();
    setupDropdownTODO();
    pushUpEmptyCellsTODO();
    updateCellCommentTODO();
    removeMultipleDatesTODO();
    updateDaysLeftCounterTODO();
    Logger.log('All functions called successfully!');
}

/**
 * Displays a "Hello World" message in an alert.
 *
 * @customfunction
 */
function logHelloWorld() {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Hello World!!');
    Logger.log('Hello world!!');
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        onOpen,
        runAllFunctionsTODO,
        logHelloWorld
    }
}
