/* eslint-disable no-unused-vars */
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

function removeMultipleDatesTODO() {
    const dataRange = getDataRange();
    const lastRow = dataRange.getLastRow();
    const columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];

    Logger.log('Init removeMultipleDatesTODO');

    for (const column of columns) {
        for (let row = 2; row <= lastRow; row++) {
            const cell = sheet.getRange(`${column}${row}`);
            const cellValue = cell.getValue();
            const richTextValue = cell.getRichTextValue();
            const text = richTextValue ? richTextValue.getText() : cellValue;

            Logger.log(`Checking cell ${column}${row}: ${text}`);

            const dateMatches = text.match(/\d{2}\/\d{2}\/\d{2}/g);
            if (dateMatches && dateMatches.length > 1) {
                Logger.log(`Found dates in ${column}${row}: ${dateMatches.join(', ')}`);

                const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy");
                Logger.log(`Today is: ${today}`);

                // filter and keep only the last occurrence of today's date
                const datesToKeep = [today];
                for (let date of dateMatches) {
                    if (date !== today) {
                        datesToKeep.push(date);
                    }
                }

                // create updated text with only the last occurrence of today's date
                let updatedText = text;
                for (let date of datesToKeep) {
                    let lastOccurrence = updatedText.lastIndexOf(date);
                    if (lastOccurrence !== -1) {
                        updatedText = updatedText.substring(0, lastOccurrence) + updatedText.substring(lastOccurrence).replace(new RegExp(date, 'g'), '');
                    }
                }

                updatedText = updatedText.replace(new RegExp(`\\b(${dateMatches.join('|')})\\b`, 'g'), '').trim() + `\n${today}`;
                Logger.log(`Updated text for ${column}${row}: ${updatedText}`);

                // build new rich text value with updated text
                let builder = SpreadsheetApp.newRichTextValue().setText(updatedText);
                let currentPos = 0;

                for (let part of updatedText.split('\n')) {
                    let startPos = currentPos;
                    let endPos = startPos + part.length;
                    if (datePattern.test(part)) {
                        builder.setTextStyle(startPos, endPos, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor('#A9A9A9').build());
                    } else {
                        let style = richTextValue.getTextStyle(startPos, endPos);
                        builder.setTextStyle(startPos, endPos, style);
                    }
                    currentPos += part.length + 1; // +1 for the newline character
                }

                const richTextResult = builder.build();
                cell.setRichTextValue(richTextResult);
                Logger.log(`Cell ${column}${row} updated with value: ${richTextResult.getText()}`);
            }
        }
    }
    Logger.log('End removeMultipleDatesTODO');
}


// for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        updateCellCommentTODO,
        exampleTextTODO,
        applyFormatToAllTODO,
        checkAndSetColumnTODO,
        setColumnBackground,
        customCeilBGColorTODO,
        setCellContentAndStyleTODO,
        updateDateColorsTODO,
        setupDropdownTODO,
        pushUpEmptyCellsTODO,
        updateRichTextTODO,
        preserveUrlsTODO,
        removeMultipleDatesTODO
    }
}