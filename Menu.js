function onOpen() {
    let ui = SpreadsheetApp.getUi();
    // Add a custom menu
    let todoSubMenu = ui.createMenu('TODO sheet')
        .addItem('Apply Format to All', 'applyFormatToAllTODO')
        .addItem('Set Ceil Background Colors', 'customCeilBGColorTODO')
        .addItem('Create Pie Chart', 'createPieChartTODO');

    // Add the custom menu and the TODO sub-menu
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
}