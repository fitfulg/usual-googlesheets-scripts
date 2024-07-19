// globals.js: ui
// TODOsheet/TODOformatting.js: applyFormatToAllTODO, customCeilBGColorTODO, createPieChartTODO, updateDateColorsTODO

function onOpen() {
    // custom menu
    let todoSubMenu = ui.createMenu('TODO sheet')
        .addItem('Apply Format to All', 'applyFormatToAllTODO')
        .addItem('Set Ceil Background Colors', 'customCeilBGColorTODO')
        .addItem('Create Pie Chart', 'createPieChartTODO');

    ui.createMenu('Custom Formats')
        .addItem('Apply Format', 'applyFormatToSelected')
        .addItem('Apply Format to All', 'applyFormatToAll')
        .addSeparator()
        .addSubMenu(todoSubMenu)
        .addItem('Log Hello World', 'logHelloWorld')
        .addToUi();

    createPieChartTODO();
    customCeilBGColorTODO();
    applyFormatToAllTODO();
    updateDateColorsTODO();
}

function logHelloWorld() {
    ui.alert('Hello World from Custom Menu!!!!!!!!___!!!!');
    console.log('Hello World from Custom Menu!');
}