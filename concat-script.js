// Auto-generated file with all JS scripts

// Contents of ./globals.js

const ui = SpreadsheetApp.getUi();
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const getDataRange = () => sheet.getDataRange();
const datePattern = /\n\d{2}\/\d{2}\/\d{2}$/; // dd/MM/yy

// state management
let isPieChartVisible = false;

// Contents of ./jest.config.js

module.exports = {
    testEnvironment: 'node',
    testPathIgnorePatterns: ['/node_modules/'],
    testMatch: ["**/tests/**/*.test.js"],
    verbose: true,
};

// Contents of ./Menu.js

// globals.js: ui
// TODOsheet/TODOformatting.js: applyFormatToAllTODO, customCeilBGColorTODO, createPieChartTODO, updateDateColorsTODO, setupDropdownTODO

function onOpen() {
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

    customCeilBGColorTODO();
    applyFormatToAllTODO();
    updateDateColorsTODO();
    setupDropdownTODO();
    pushUpEmptyCellsTODO();
    updateCellCommentTODO()
}

function logHelloWorld() {
    ui.alert('Hello World from Custom Menu!!!');
    console.log('Hello World from Custom Menu!');
}
// Contents of ./shared/formatting.js

// globals.js: sheet, getDataRange

// Higher-order fn to apply formatting to a range only if it is valid
const withValidRange = (fn) => (range, ...args) => range && fn(range, ...args);

const Format = withValidRange((range) => {
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
});

const applyBordersWithStyle = withValidRange((range, borderStyle) => range.setBorder(true, true, true, true, true, true, "#000000", borderStyle));
const applyBorders = range => applyBordersWithStyle(range, SpreadsheetApp.BorderStyle.SOLID);
const applyThickBorders = range => applyBordersWithStyle(range, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

function applyFormatToSelected() {
    let range = sheet.getActiveRange();
    Format(range);
    applyBorders(range);
}

function applyFormatToAll() {
    let range = getDataRange();
    Format(range);
    applyBorders(range);
}

function setCellStyle(cell, value, fontWeight, fontColor, backgroundColor, alignment) {
    let range = sheet.getRange(cell);
    range.setValue(value)
        .setFontWeight(fontWeight)
        .setFontColor(fontColor)
        .setHorizontalAlignment(alignment);

    if (backgroundColor) {
        range.setBackground(backgroundColor);
    }
}

// Append DATE to cell
function appendDateWithStyle(cellValue, dateFormatted, column, config) {
    const newText = cellValue.endsWith('\n' + dateFormatted) ? cellValue : cellValue.trim() + '\n' + dateFormatted;
    return createRichTextValue(newText, dateFormatted, column, config);
}

// Update DATE in cell if it already exists
function updateDateWithStyle(cellValue, dateFormatted, column, config) {
    const newText = cellValue.replace(datePattern, '\n' + dateFormatted).trim();
    return createRichTextValue(newText, dateFormatted, column, config);
}

// Create rich text value with italic date
function createRichTextValue(text, dateFormatted, column, config) {
    const columnConfig = config[column];
    const color = columnConfig.defaultColor || '#A9A9A9'; // Default color (dark gray)

    return SpreadsheetApp.newRichTextValue()
        .setText(text)
        .setTextStyle(text.length - dateFormatted.length, text.length, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor(color).build())
        .build();
}

// Reset the text style of a cell
function resetTextStyle(range) {
    const richTextValue = SpreadsheetApp.newRichTextValue()
        .setText(range.getValue())
        .setTextStyle(SpreadsheetApp.newTextStyle().build())
        .build();

    range.setRichTextValue(richTextValue);
}

function clearTextFormatting(range) {
    const values = range.getValues();
    const richTextValues = values.map(row => row.map(value =>
        SpreadsheetApp.newRichTextValue()
            .setText(value)
            .setTextStyle(SpreadsheetApp.newTextStyle().build())
            .build()
    ));
    range.setRichTextValues(richTextValues);
}

// for testing 
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        // applyFormatToSelected,
        // applyFormatToAll,
        setCellStyle,
        // appendDateWithStyle,
        // updateDateWithStyle,
        // createRichTextValue,
        // resetTextStyle,
        // clearTextFormatting
    };
}
// Contents of ./shared/utils.js

function extractUrls(richTextValue) {
    const urls = [];
    const text = richTextValue.getText();
    for (let i = 0; i < text.length; i++) {
        const url = richTextValue.getLinkUrl(i, i + 1);
        if (url) {
            urls.push(url);
        }
    }
    return urls;
}

function arraysEqual(arr1, arr2) {
    if (arr1.length !== arr2.length) return false;
    for (let i = 0; i < arr1.length; i++) {
        if (arr1[i] !== arr2[i]) return false;
    }
    return true;
}
// Contents of ./TODOsheet/TODOformatting.js

// globals.js: sheet, getDataRange, datePattern
// shared/formatting.js: Format, applyBorders, applyThickBorders, setCellStyle
// TODOsheet/TODOlibrary.js: dateColorConfig

function updateCellCommentTODO() {
    const cell = sheet.getRange("I2");
    const version = "v1.1";
    const emoji = "💡";
    const changes = `
        - There is an indicative limit of cells for each priority. In the end the objective of a TODO is none other than to complete the tasks and that they do not accumulate. Once this limit is reached, a warning is activated for the entire column.
        This feature does not block cells, that is, you can continue occupying cells even if you have the warning.\n
        - You can apply some custom formats that do not require to refresh the page from the "Custom Formats" menu.\n
        - Writing or modifying a cell causes the current date to be added, which over time changes color from gray to orange and from orange to red.\n
        - The date color change times are different for each column, with HIGH PRIORITY being the fastest to change and LOW PRIORITY being the slowest.\n
        - The Piechart can be shown or hidden directly using its dropdown cell.\n
        - Empty cells that are deleted are occupied by their immediately lower cell.\n
        - Empty cells that remain empty are occupied by the cell immediately below them by opening or refreshing the page.\n
    `;

    const comment = `Versión: ${version}\nFEATURES:\n${changes}`;
    cell.setComment(comment);
    cell.setBackground("#efefef");
    cell.setBorder(true, true, true, true, true, true, '#D3D3D3', SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Crear RichTextValue con diferentes tamaños de fuente
    const richText = SpreadsheetApp.newRichTextValue()
        .setText(`${version}\n${emoji}`)
        .setTextStyle(0, version.length, SpreadsheetApp.newTextStyle().setFontSize(8).build())
        .setTextStyle(version.length + 1, version.length + 2, SpreadsheetApp.newTextStyle().setFontSize(20).build())
        .setTextStyle(version.length + 2, version.length + 3, SpreadsheetApp.newTextStyle().setFontSize(20).build())
        .build();

    cell.setRichTextValue(richText);
    Format(cell);
}

function exampleTextTODO(column, exampleText) {
    const dataRange = getDataRange();
    let values;

    if (column === "B") {
        // Get values excluding B3 and B8
        values = [
            sheet.getRange(column + "2").getValue(),
            ...sheet.getRange(column + "4:" + column + "7").getValues().flat(),
            ...sheet.getRange(column + "9:" + column + dataRange.getLastRow()).getValues().flat()
        ];
    } else {
        values = sheet.getRange(column + "2:" + column + dataRange.getLastRow()).getValues().flat();
    }

    const isEmpty = values.every(value => !value.toString().trim());

    if (isEmpty) {
        const cell = sheet.getRange(column + "2");
        cell.setValue(exampleText);
    }
}

function applyFormatToAllTODO() {
    const totalRows = sheet.getMaxRows();

    // Get the range for all the columns A to H up to the last row
    let range = sheet.getRange(1, 1, totalRows, 8); // A1:H(last row)
    if (range) {
        Format(range);
        applyBorders(range);
    }

    // Apply thicker borders to specific columns C, D, and E for defined rows
    applyThickBorders(sheet.getRange(1, 3, 11, 1)); // C1:C11
    applyThickBorders(sheet.getRange(1, 4, 21, 1)); // D1:D21
    applyThickBorders(sheet.getRange(1, 5, 21, 1)); // E1:E21

    // Set the specific content and styles in the specified cells
    setCellContentAndStyleTODO();

    // Check the number of occupied cells in columns C, D, and E
    checkAndSetColumnTODO("C", 9, "HIGH PRIORITY");
    checkAndSetColumnTODO("D", 19, "MEDIUM PRIORITY");
    checkAndSetColumnTODO("E", 19, "LOW PRIORITY");

    // Add example text to specific columns if empty
    for (const column in exampleTexts) {
        const { text } = exampleTexts[column];
        exampleTextTODO(column, text);
    }
}

function checkAndSetColumnTODO(column, limit, priority) {
    const dataRange = getDataRange();
    const values = sheet.getRange(column + "2:" + column + dataRange.getLastRow()).getValues().flat();
    const occupied = values.filter(String).length;
    const range = sheet.getRange(column + "2:" + column + dataRange.getLastRow());

    if (occupied > limit) {
        // red with thicker border
        range.setBorder(true, true, true, true, true, true, "#FF0000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(column + "1").setValue("⚠️CELL LIMIT REACHED⚠️");
        SpreadsheetApp.getUi().alert("⚠️CELL LIMIT REACHED⚠️ \nfor priority: " + priority);
    } else {
        // black with thicker border
        range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(column + "1").setValue(priority);
    }
}

function setColumnBackground(sheet, col, color, startRow = 2) {
    let totalRows = sheet.getMaxRows();
    let range = sheet.getRange(startRow, col, totalRows - startRow + 1, 1);
    range.setBackground(color);
}
function customCeilBGColorTODO() {
    // Apply background colors to specific columns
    setColumnBackground(sheet, 1, '#d3d3d3', 2); // Column A: Light gray 3
    setColumnBackground(sheet, 6, '#fff1f1', 2); // Column F: Light pink
    setColumnBackground(sheet, 7, '#d3d3d3', 2); // Column G: Light gray 3

    // Apply white background to columns B, C, D, E, H, I starting from row 2
    let whiteColumns = [2, 3, 4, 5, 8, 9]; // Columns B, C, D, E, H, I
    for (let col of whiteColumns) {
        setColumnBackground(sheet, col, '#ffffff', 2);
    }

    // Apply dark yellow background to specific cells in column B
    sheet.getRange('B3').setBackground('#b5a642'); // Dark yellow 3
    sheet.getRange('B8').setBackground('#b5a642'); // Dark yellow 3
}

function setCellContentAndStyleTODO() {
    for (const cell in cellStyles) {
        const { value, fontWeight, fontColor, backgroundColor, alignment } = cellStyles[cell];
        setCellStyle(cell, value, fontWeight, fontColor, backgroundColor, alignment);
    }
}

// update date colors based on time passed
function updateDateColorsTODO() {
    const columns = ['C', 'D', 'E', 'F', 'G', 'H'];
    const dataRange = getDataRange();
    const lastRow = dataRange.getLastRow();

    for (const column of columns) {
        const config = dateColorConfig[column];
        for (let row = 2; row <= lastRow; row++) {
            const cell = sheet.getRange(`${column}${row}`);
            const cellValue = cell.getValue();
            if (datePattern.test(cellValue)) {
                const dateText = cellValue.match(datePattern)[0].trim();
                const cellDate = new Date(dateText.split('/').reverse().join('/'));
                const today = new Date();
                const diffDays = Math.floor((today - cellDate) / (1000 * 60 * 60 * 24));

                let color = config.defaultColor || '#A9A9A9'; // Default color (dark gray)
                if (diffDays >= config.danger) {
                    color = config.dangerColor;
                } else if (diffDays >= config.warning) {
                    color = config.warningColor;
                }

                const richTextValue = SpreadsheetApp.newRichTextValue()
                    .setText(cellValue)
                    .setTextStyle(cellValue.length - dateText.length, cellValue.length, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor(color).build())
                    .build();

                cell.setRichTextValue(richTextValue);
            }
        }
    }
}

function setupDropdownTODO() {
    // Setup dropdown in I1
    const buttonCell = sheet.getRange("I1");
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Piechart', 'Show Piechart', 'Hide Piechart'], true).build();
    buttonCell.setDataValidation(rule);
    buttonCell.setValue('Piechart');
    buttonCell.setFontWeight('bold');
    buttonCell.setFontSize(12);
    buttonCell.setHorizontalAlignment("center");
    buttonCell.setVerticalAlignment("middle");
}

// Shift cells up if empty 
function shiftCellsUpTODO(column, startRow, endRow) {
    Logger.log(`shiftCellsUpTODO called for column: ${column}, from row ${startRow} to ${endRow}`);
    const range = sheet.getRange(startRow, column, endRow - startRow + 1, 1);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();
    const newValues = [];
    const newRichTextValues = [];

    for (let i = 0; i < values.length; i++) {
        Logger.log(`Value at row ${i + startRow}: ${values[i][0]}`);
        if (values[i][0].toString().trim() !== '') {
            newValues.push([values[i][0]]);
            newRichTextValues.push([richTextValues[i][0]]);
        }
    }

    while (newValues.length < values.length) {
        newValues.push(['']);
        newRichTextValues.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
    }

    if (newValues.length > 0) {
        Logger.log('Setting new values and rich text values');
        range.setValues(newValues);
        range.setRichTextValues(newRichTextValues);
    }

    if (values.length > newValues.length) {
        const emptyRange = sheet.getRange(startRow + newValues.length, column, values.length - newValues.length, 1);
        clearTextFormatting(emptyRange);
    }

    Logger.log('shiftCellsUpTODO completed');
}

// Force push up empty cells in columns A, C, D, E, F, G, H
function pushUpEmptyCellsTODO() {
    const dataRange = getDataRange();
    const totalRows = dataRange.getLastRow();
    const columns = [1, 3, 4, 5, 6, 7, 8]; // A, C, D, E, F, G, H

    columns.forEach(column => {
        for (let row = 2; row <= totalRows; row++) {
            const cell = sheet.getRange(row, column);
            const cellValue = cell.getValue().toString().trim();
            if (cellValue === '') {
                Logger.log(`Empty cell found at ${cell.getA1Notation()}, shifting cells up`);
                shiftCellsUpTODO(column, 2, totalRows);
                break; // Reset the loop for the same column
            }
        }
    });

    Logger.log('pushUpEmptyCells completed');
}

function updateRichTextTODO(range, originalValue, newValue, columnLetter, row, e) {
    const cellValue = newValue;
    Logger.log(`Cell value after edit: ${cellValue}`);

    // Get rich text value of the edited cell, or use the plain cell value
    const richTextValue = range.getRichTextValue();
    const text = richTextValue ? richTextValue.getText() : cellValue;

    // Retrieve original rich text value before edit, or create new rich text value if not available
    const originalRichText = e.oldRichTextValue || SpreadsheetApp.newRichTextValue().setText(originalValue).build();
    const originalText = originalRichText.getText();

    const originalUrls = extractUrls(originalRichText);
    const newUrls = extractUrls(richTextValue);

    Logger.log(`Original URLs: ${JSON.stringify(originalUrls)}, New URLs: ${JSON.stringify(newUrls)}`);

    if (originalText === text && arraysEqual(originalUrls, newUrls)) {
        Logger.log('No change in cell value or links, skipping update');
        return;
    }

    if (text.trim() === "") return resetTextStyle(range);

    const dateFormatted = ` ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy")}`;

    // Append or update the date in the text based on whether the date pattern exists
    const newRichTextValue = datePattern.test(text)
        ? updateDateWithStyle(text, dateFormatted, columnLetter, dateColorConfig)
        : appendDateWithStyle(text, dateFormatted, columnLetter, dateColorConfig);

    Logger.log(`Setting rich text value for cell ${columnLetter}${row}`);
    range.setRichTextValue(newRichTextValue);

    preserveUrlsTODO(range, richTextValue, newRichTextValue);
}

function preserveUrlsTODO(range, richTextValue, newRichTextValue) {
    const updatedRichTextValue = range.getRichTextValue();
    const updatedText = updatedRichTextValue.getText();
    const finalRichTextValue = SpreadsheetApp.newRichTextValue().setText(updatedText);

    for (let i = 0; i < updatedText.length; i++) {
        const url = richTextValue.getLinkUrl(i, i + 1);
        if (url) {
            finalRichTextValue.setLinkUrl(i, i + 1, url);
        }
    }
    range.setRichTextValue(finalRichTextValue.build());
}

// Contents of ./TODOsheet/TODOlibrary.js

const cellStyles = {
    "A1": { value: "QUICK PATTERNS", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#000000", alignment: "center" },
    "B1": { value: "TOMORROW", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#b5a642", alignment: "center" },
    "B3": { value: "WEEK", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#b5a642", alignment: "center" },
    "B8": { value: "MONTH", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#b5a642", alignment: "center" },
    "F1": { value: "💡IDEAS AND PLANS", fontWeight: "bold", fontColor: "#000000", backgroundColor: "#FFC0CB", alignment: "center" },
    "G1": { value: "👀 EYES ON", fontWeight: "bold", fontColor: "#000000", backgroundColor: "#b7b7b7", alignment: "center" },
    "H1": { value: "IN QUARANTINE BEFORE BEING CANCELED", fontWeight: "bold", fontColor: "#FF0000", backgroundColor: null, alignment: "center" },
    "C1": { value: "HIGH PRIORITY", fontWeight: "bold", fontColor: null, backgroundColor: "#fce5cd", alignment: "center" },
    "D1": { value: "MEDIUM PRIORITY", fontWeight: "bold", fontColor: null, backgroundColor: "#fff2cc", alignment: "center" },
    "E1": { value: "LOW PRIORITY", fontWeight: "bold", fontColor: null, backgroundColor: "#d9ead3", alignment: "center" }
};

const exampleTexts = {
    "A": { text: "Example: Do it with fear but do it.", color: "#FFFFFF" },
    "B": { text: "Example: 45min of cardio", color: "#A9A9A9" },
    "C": { text: "Example: Join that gym club", color: "#A9A9A9" },
    "D": { text: "Example: Submit that pending data science task.", color: "#A9A9A9" },
    "E": { text: "Example: Buy a new mattress.", color: "#A9A9A9" },
    "F": { text: "Example: Santiago route.", color: "#A9A9A9" },
    "G": { text: "Example: Change front brake pad at 44500km", color: "#FFFFFF" },
    "H": { text: "Example: Join that Crossfit club", color: "#A9A9A9" },
};

const dateColorConfig = {
    C: { warning: 7, danger: 30, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' }, // 1 week, 1 month
    D: { warning: 90, danger: 180, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
    E: { warning: 180, danger: 365, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
    F: { warning: 180, danger: 365, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
    G: { warning: 0, danger: 0, warningColor: '#A9A9A9', dangerColor: '#A9A9A9', defaultColor: '#A9A9A9' }, // Always default
    H: { warning: 0, danger: 0, warningColor: '#FF0000', dangerColor: '#FF0000', defaultColor: '#FF0000' } // Always red
};
// Contents of ./TODOsheet/TODOpiechart.js

// globals.js: sheet, getDataRange

function createPieChartTODO() {
    Logger.log('Creating piechart');
    const dataRange = getDataRange();
    const valuesC = sheet.getRange("C2:C" + dataRange.getLastRow()).getValues().flat();
    const valuesD = sheet.getRange("D2:D" + dataRange.getLastRow()).getValues().flat();
    const valuesE = sheet.getRange("E2:E" + dataRange.getLastRow()).getValues().flat();

    const occupiedC = valuesC.filter(String).length;
    const occupiedD = valuesD.filter(String).length;
    const occupiedE = valuesE.filter(String).length;

    const chartDataRange = sheet.getRange("J1:K4");
    chartDataRange.setValues([
        ["Column", "Occupied Cells"],
        ["HIGH", occupiedC],
        ["MEDIUM", occupiedD],
        ["LOW", occupiedE]
    ]);

    const chart = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartDataRange)
        .setPosition(1, 10, 0, 0) // Position the chart starting at column J
        .setOption('title', 'Pie Chart')
        .build();

    sheet.insertChart(chart);
    Logger.log('Piechart created');
    isPieChartVisible = true;
}

function deleteAllChartsTODO() {
    Logger.log('Deleting all charts');
    const charts = sheet.getCharts();

    charts.forEach(chart => {
        sheet.removeChart(chart);
    });

    sheet.getRange("J1:K4").clearContent();
    Logger.log(`Deleted ${charts.length} charts`);
    isPieChartVisible = false;
}

// Contents of ./TODOsheet/TODOtoggleFn.js

// TODOsheet/TODOtoggleFn.js: createPieChartTODO, deleteAllChartsTODO

function togglePieChartTODO(action) {
    Logger.log(`togglePieChartTODO called with action: ${action}`);
    if (action === 'Hide Piechart') {
        deleteAllChartsTODO();
        isPieChartVisible = false;
        Logger.log('Piechart hidden');
    } else if (action === 'Show Piechart') {
        createPieChartTODO();
        isPieChartVisible = true;
        Logger.log('Piechart shown');
    } else {
        Logger.log('Invalid action selected');
    }
}

function handlePieChartToggleTODO(range) {
    const action = range.getValue().toString().trim();
    Logger.log(`Action selected: ${action}`);
    if (action === 'Show Piechart' || action === 'Hide Piechart') {
        togglePieChartTODO(action);
    } else {
        Logger.log('Invalid action selected');
    }
    sheet.getRange("I1").setValue("Piechart");
}



// Contents of ./TODOsheet/TODOtriggers.js

// globals.js: sheet, datePattern, getDataRange
// shared/formatting.js: resetTextStyle, appendDateWithStyle, updateDateWithStyle
// shared/utils.js: extractUrls, arraysEqual

// Track changes in specified columns and add the date
function onEdit(e) {
    try {
        if (!e || !e.range) {
            Logger.log('Edit event is undefined or does not have range property');
            return;
        }

        const { range } = e;
        const column = range.getColumn();
        const row = range.getRow();
        const columnLetter = String.fromCharCode(64 + column);
        const totalRows = sheet.getMaxRows();

        Logger.log(`onEdit triggered: column ${column}, row ${row}`);

        // Check if the edited cell is for toggling the pie chart (cell I1)
        if (column === 9 && row === 1) {
            handlePieChartToggle(range);
            return;
        }

        const originalValue = e.oldValue || '';
        const newValue = range.getValue().toString();

        Logger.log(`Original value: ${originalValue}, New value: ${newValue}`);

        // Shift cells up if the edited cell is in columns A, C, D, E, F, G, H and is now empty
        if ((column === 1 || (column >= 3 && column <= 8)) && row >= 2 && newValue.trim() === '') {
            Logger.log(`Shifting cells up for column ${column}`);
            shiftCellsUpTODO(column, 2, totalRows);
        }

        // Check if the edit is in columns C, D, E, F, G, H and from row 2 onwards
        if (column >= 3 && column <= 8 && row >= 2) {
            updateRichTextTODO(range, originalValue, newValue, columnLetter, row, e);
        }
    } catch (error) {
        Logger.log(`Error in onEdit: ${error.message}`);
    }
}














