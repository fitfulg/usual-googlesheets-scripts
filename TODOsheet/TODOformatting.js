// globals.js: sheet, getDataRange, datePattern
// shared/formatting.js: Format, applyBorders, applyThickBorders, setCellStyle
// TODOsheet/TODOlibrary.js: dateColorConfig

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

    const isEmpty = values.every(value => !value.trim());

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
    checkAndSetColumn("C", 9, "HIGH PRIORITY");
    checkAndSetColumn("D", 19, "MEDIUM PRIORITY");
    checkAndSetColumn("E", 19, "LOW PRIORITY");

    // Add example text to specific columns if empty
    for (const column in exampleTexts) {
        const { text } = exampleTexts[column];
        exampleTextTODO(column, text);
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
            if (datePattern.test(cellValue)) {// test() is a
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

function shiftCellsUpTODO(column, startRow, endRow) {
    Logger.log(`shiftCellsUpTODO called for column: ${column}, from row ${startRow} to ${endRow}`);
    const range = sheet.getRange(startRow, column, endRow - startRow + 1, 1);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();

    const newValues = [];
    const newRichTextValues = [];

    // Filter out empty values
    for (let i = 0; i < values.length; i++) {
        Logger.log(`Value at row ${i + startRow}: ${values[i][0]}`);
        if (values[i][0].trim() !== '') {
            newValues.push([values[i][0]]);
            newRichTextValues.push([richTextValues[i][0]]);
        }
    }
    // Add empty values to match the original range size
    while (newValues.length < values.length) {
        newValues.push(['']);
        newRichTextValues.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
    }
    Logger.log('Setting new values and rich text values');
    range.setValues(newValues);
    range.setRichTextValues(newRichTextValues);

    // Clear text formatting for the new empty cells
    const emptyRange = sheet.getRange(startRow + newValues.length, column, values.length - newValues.length, 1);
    clearTextFormatting(emptyRange);

    Logger.log('shiftCellsUpTODO completed');
}




