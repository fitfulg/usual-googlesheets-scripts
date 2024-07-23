/* eslint-disable no-unused-vars */
// globals.js: ui
// TODOsheet/TODOformatting.js: applyFormatToAllTODO, customCeilBGColorTODO, createPieChartTODO, updateDateColorsTODO, setupDropdownTODO

function onOpen() {
    Logger.log('onOpen triggered');

    const docProperties = PropertiesService.getDocumentProperties();
    const lastHash = docProperties.getProperty('lastHash');
    const currentHash = getSheetContentHash();

    if (shouldRunUpdates(lastHash, currentHash)) {
        runAllFunctionsTODO();
        docProperties.setProperty('lastHash', currentHash);
        Logger.log('Running all update functions');
    } else {
        Logger.log('It is not necessary to run all functions, the data has not changed significantly.');
    }
    // custom menu
    let todoSubMenu = ui.createMenu('TODO sheet')
        .addItem('Apply Format to All', 'applyFormatToAllTODO')
        .addItem('Set Ceil Background Colors', 'customCeilBGColorTODO')
        .addItem('Create Pie Chart', 'createPieChartTODO')
        .addItem('Delete Pie Charts', 'deleteAllChartsTODO');

    ui.createMenu('Custom Formats')
        .addItem('Apply Format', 'applyFormatToSelected')
        .addItem('Apply Format to All', 'applyFormatToAll')
        .addSeparator()
        .addSubMenu(todoSubMenu)
        .addItem('Log Hello World', 'logHelloWorld')
        .addToUi();
}

function runAllFunctionsTODO() {
    customCeilBGColorTODO();
    applyFormatToAllTODO();
    updateDateColorsTODO();
    setupDropdownTODO();
    pushUpEmptyCellsTODO();
    updateCellCommentTODO();
    removeMultipleDatesTODO();
    Logger.log('All functions called successfully!');
}

function logHelloWorld() {
    ui.alert('Hello World from Custom Menu!!');
    Logger.log('Hello World from Custom Menu!');
}