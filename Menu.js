function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Add a custom menu
    ui.createMenu('Custom Formats')
        .addItem('Apply Format', 'applyFormatToSelected')
        .addItem('Apply Format to All', 'applyFormatToAll')
        .addItem('Set Ceil Background Colors from TODO sheet', 'customCeilBGColorTODO')
        .addItem('Log Hello World', 'logHelloWorld') // Añadir nueva función al menú
        .addToUi();
}

function logHelloWorld() {
    Logger.log("Hello, World from Github to GoogleSheets!");
}
