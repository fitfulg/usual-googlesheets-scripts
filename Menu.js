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

    // Call function when the document is opened or refreshed
    // showLoading();
    applyFormatToAllTODO();
    createPieChart();
    customCeilBGColorTODO();
    // hideLoading();
}

function logHelloWorld() {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Hello, World from Github to GoogleSheets!");
}
// IDEA to implement :
// function showLoading() {
//     SpreadsheetApp.getActiveSpreadsheet().toast('Loading, please wait...', 'Loading', -1);
// }

// function hideLoading() {
//     SpreadsheetApp.getUi().alert('Loading complete!');
// }
