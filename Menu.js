function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Add a custom menu
    ui.createMenu('Custom Formats')
        .addItem('Apply Format', 'applyFormat')
        .addItem('Apply Format to All', 'applyFormatToAll')
        .addItem('Set Background Colors TODO', 'backgroundColorsTODO')
        .addItem('Log Hello World from github to googlesheets', 'logHelloWorld') // Añadir nueva función al menú
        .addToUi();
}

function logHelloWorld() {
    Logger.log("Hello, World!!");
}
