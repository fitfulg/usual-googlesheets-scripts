function onOpen() {
    let ui = SpreadsheetApp.getUi();
    // Add a custom menu
    ui.createMenu('Custom Formats')
        .addItem('Apply Format', 'applyFormatToSelected')
        .addItem('Apply Format to All', 'applyFormatToAll')
        .addItem('Set Ceil Background Colors from TODO sheet', 'customCeilBGColorTODO')
        .addItem('Log Hello World', 'logHelloWorld')
        .addItem('Create Pie Chart', 'createPieChart')
        .addToUi();
}

function logHelloWorld() {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Hello, World from Github to GoogleSheets!");
}