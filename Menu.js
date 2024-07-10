function onOpen() {
    let ui = SpreadsheetApp.getUi();
    // Add a custom menu
    ui.createMenu('Custom Formats')
        .addItem('Apply Format', 'applyFormatToSelected')
        .addItem('Apply Format to All', 'applyFormatToAll')
        .addItem('TODOsheet: Apply Format to All', 'applyFormatToAllTODO')
        .addItem('TODOsheet: Set Ceil Background Colors', 'customCeilBGColorTODO')
        .addItem('TODOsheet: Create Pie Chart', 'createPieChart')
        .addItem('Log Hello World', 'logHelloWorld')
        .addToUi();

    // Call func when the document is opened or F5ed
    applyFormatToAllTODO();
    createPieChartTODO();
    customCeilBGColorTODO();
}

function logHelloWorld() {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Hello, World from Github to GoogleSheets!");
}