// Auto-generated file with all JS scripts

// Contents of ./globals.js

 

const ui = SpreadsheetApp.getUi();
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const getDataRange = () => sheet.getDataRange();
const datePattern = /\n\d{2}\/\d{2}\/\d{2}$/; // dd/MM/yy

// state management
let isPieChartVisible = false;


// Contents of ./Menu.js

// globals.js: ui
// shared/utils.js: getSheetContentHash, shouldRunUpdates
// shared/formatting: applyFormatToSelected, applyFormatToAll
// TODOsheet/TODOformatting.js: applyFormatToAllTODO, customCeilBGColorTODO, createPieChartTODO, deleteAllChartsTODO, updateDateColorsTODO, setupDropdownTODO, pushUpEmptyCellsTODO, updateCellCommentTODO, removeMultipleDatesTODO, updateDaysLeftTODO
// TODOsheet/TODOcheckbox.js: addCheckboxToCellTODO, addCheckboxesToSelectedCellsTODO, markCheckboxSelectedCellsTODO, markAllCheckboxesSelectedCellsTODO, removeCheckboxesFromSelectedCellsTODO

/**
 * Initializes the UI menu in the spreadsheet.
 * Sets up custom menus and triggers functions when menu items are clicked.
 *
 * @customfunction
 */
function onOpen() {
    Logger.log('onOpen triggered');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docProperties = PropertiesService.getDocumentProperties();
    const language = docProperties.getProperty('language') || 'English';

    saveSnapshotTODO()
    Logger.log('Current language: ' + language);

    ss.toast(toastMessages.loading[language], 'Status:', 13);
    try {
        const docProperties = PropertiesService.getDocumentProperties();
        const lastHash = docProperties.getProperty('lastHash');
        const currentHash = getSheetContentHash();

        if (shouldRunUpdates(lastHash, currentHash)) {
            runAllFunctionsTODO();

            restoreSnapshotTODO();
            updateDaysLeftCounterTODO();
            docProperties.setProperty('lastHash', currentHash);
            Logger.log('Running all update functions');
        } else {
            Logger.log('It is not necessary to run all functions, the data has not changed significantly.');
        }

        createMenusTODO();
        translateSheetTODO();
        ss.toast(toastMessages.updateComplete[language], 'Status:', 5);
    } catch (e) {
        Logger.log('Error: ' + e.toString());
        ui.alert('Error during processing: ' + e.toString());
    }
}

/**
 * Runs all functions needed to update the TODO sheet.
 * Calls multiple formatting and update functions.
 *
 * @customfunction
 */
function runAllFunctionsTODO() {
    Logger.log('runAllFunctionsTODO triggered');
    customCellBGColorTODO();
    applyFormatToAllTODO();
    updateDateColorsTODO();
    setupDropdownTODO();
    pushUpEmptyCellsTODO();
    updateCellCommentTODO();
    removeMultipleDatesTODO();
    updateDaysLeftCounterTODO();
    Logger.log('All functions called successfully!');
}


// Contents of ./shared/formatting.js


// globals.js: sheet, getDataRange

/**
 * Higher-order function to apply formatting to a range only if it is valid.
 *
 * @param {Function} fn - The formatting function to apply.
 * @return {Function} A function that applies formatting if the range is valid.
 */
const withValidRange = (fn) => (range, ...args) => range && fn(range, ...args);

/**
 * Applies wrap strategy and alignment to a range.
 *
 * @param {Range} range - The range to format.
 */
const Format = withValidRange((range) => {
    Logger.log('Format triggered');
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
});

/**
 * Applies border style to a range.
 *
 * @param {Range} range - The range to format.
 * @param {BorderStyle} borderStyle - The border style to apply.
 */
const applyBordersWithStyle = withValidRange((range, borderStyle) => range.setBorder(true, true, true, true, true, true, "#000000", borderStyle));

/**
 * Applies solid borders to a range.
 *
 * @param {Range} range - The range to apply borders to.
 */
const applyBorders = range => applyBordersWithStyle(range, SpreadsheetApp.BorderStyle.SOLID);

/**
 * Applies thick borders to a range.
 *
 * @param {Range} range - The range to apply thick borders to.
 */
const applyThickBorders = range => applyBordersWithStyle(range, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

/**
 * Applies formatting to the selected range.
 *
 * @customfunction
 */
function applyFormatToSelected() {
    Logger.log('applyFormatToSelected triggered');
    let range = sheet.getActiveRange();
    Format(range);
    applyBorders(range);
}

/**
 * Applies formatting to all data in the sheet.
 *
 * @customfunction
 */
function applyFormatToAll() {
    Logger.log('applyFormatToAll triggered');
    let range = getDataRange();
    Format(range);
    applyBorders(range);
}

/**
 * Sets the content and style of a specific cell.
 *
 * @param {string} cell - The cell to set.
 * @param {string} value - The value to set.
 * @param {string} fontWeight - The font weight to set.
 * @param {string} fontColor - The font color to set.
 * @param {string} backgroundColor - The background color to set.
 * @param {string} alignment - The alignment to set.
 */
function setCellStyle(cell, value, fontWeight, fontColor, backgroundColor, alignment) {
    Logger.log('setCellStyle triggered');
    let range = sheet.getRange(cell);
    Logger.log('setCellStyle: setting value');
    range.setValue(value)
        .setFontWeight(fontWeight)
        .setFontColor(fontColor)
        .setHorizontalAlignment(alignment);
    Logger.log('setCellStyle: setting background color');
    if (backgroundColor) {
        range.setBackground(backgroundColor);
    }
}

/**
 * Appends a formatted date to a cell value.
 *
 * @param {string} cellValue - The current cell value.
 * @param {string} dateFormatted - The formatted date to append.
 * @param {string} column - The column of the cell.
 * @param {Object} config - The configuration for formatting.
 * @return {RichTextValue} The new rich text value.
 */
function appendDateWithStyle(cellValue, dateFormatted, column, config) {
    Logger.log('appendDateWithStyle triggered');
    const newText = cellValue.endsWith('\n' + dateFormatted) ? cellValue : cellValue.trim() + '\n' + dateFormatted;
    Logger.log('returning createRichTextValue');
    return createRichTextValue(newText, dateFormatted, column, config);
}

/**
 * Updates a formatted date in a cell value.
 *
 * @param {string} cellValue - The current cell value.
 * @param {string} dateFormatted - The formatted date to update.
 * @param {string} column - The column of the cell.
 * @param {Object} config - The configuration for formatting.
 * @return {RichTextValue} The new rich text value.
 */
function updateDateWithStyle(cellValue, dateFormatted, column, config) {
    const newText = cellValue.replace(datePattern, '\n' + dateFormatted).trim();
    Logger.log('returning createRichTextValue');
    return createRichTextValue(newText, dateFormatted, column, config);
}

/**
 * Creates a rich text value with an italic date.
 *
 * @param {string} text - The text to format.
 * @param {string} dateFormatted - The formatted date.
 * @param {string} column - The column of the cell.
 * @param {Object} config - The configuration for formatting.
 * @return {RichTextValue} The new rich text value.
 */
function createRichTextValue(text, dateFormatted, column, config) {
    Logger.log('createRichTextValue triggered');
    const columnConfig = config[column];
    const color = columnConfig.defaultColor || '#A9A9A9'; // Default color (dark gray)
    Logger.log('returning SpreadsheetApp.newRichTextValue');
    return SpreadsheetApp.newRichTextValue()
        .setText(text)
        .setTextStyle(text.length - dateFormatted.length, text.length, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor(color).build())
        .build();
}

/**
 * Helper function to create or update RichTextValue with the given parameters.
 * 
 * @param {string} text - The full text for the RichTextValue.
 * @param {string} subText - The part of the text to be styled.
 * @param {string} color - The color to apply to the subText.
 * @param {RichTextValue} originalRichTextValue - The original RichTextValue to preserve links.
 * @return {RichTextValue} The updated RichTextValue.
 */
function buildRichTextValue(text, subText, color, originalRichTextValue) {
    Logger.log(`buildRichTextValue(): Building rich text for "${subText}" with color ${color}`);

    const newRichTextValueBuilder = SpreadsheetApp.newRichTextValue()
        .setText(text)
        .setTextStyle(0, text.length, SpreadsheetApp.newTextStyle().build())
        .setTextStyle(text.length - subText.length, text.length, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor(color).build());

    if (originalRichTextValue) {
        const originalTextLength = originalRichTextValue.getText().length;
        for (let i = 0; i < Math.min(text.length, originalTextLength); i++) {
            const url = originalRichTextValue.getLinkUrl(i, i + 1);
            if (url) {
                newRichTextValueBuilder.setLinkUrl(i, i + 1, url);
                Logger.log(`buildRichTextValue(): Set link for character ${i}: ${url}`);
            }
        }
    }

    return newRichTextValueBuilder.build();
}

/**
 * Resets the text style of a cell.
 *
 * @param {Range} range - The range to reset.
 */
function resetTextStyle(range) {
    Logger.log('resetTextStyle triggered');
    const richTextValue = SpreadsheetApp.newRichTextValue()
        .setText(range.getValue())
        .setTextStyle(SpreadsheetApp.newTextStyle().build())
        .build();
    Logger.log('resetTextStyle: setting rich text value');
    range.setRichTextValue(richTextValue);
}

/**
 * Clears the text formatting of a range.
 *
 * @param {Range} range - The range to clear.
 */
function clearTextFormatting(range) {
    Logger.log('clearTextFormatting triggered');
    const values = range.getValues();
    const richTextValues = values.map(row => row.map(value =>
        SpreadsheetApp.newRichTextValue()
            .setText(value)
            .setTextStyle(SpreadsheetApp.newTextStyle().build())
            .build()
    ));
    Logger.log('clearTextFormatting: setting rich text values');
    range.setRichTextValues(richTextValues);
}

// for testing 

// Contents of ./shared/utils.js



/**
 * Extracts URLs from a rich text value.
 *
 * @param {RichTextValue} richTextValue - The rich text value to extract URLs from.
 * @return {Array} The extracted URLs with their start and end positions.
 */
function extractUrls(richTextValue) {
    Logger.log('extractUrls triggered');
    const urls = [];
    const text = richTextValue.getText();
    for (let i = 0; i < text.length; i++) {
        const url = richTextValue.getLinkUrl(i, i + 1);
        if (url) {
            urls.push({ url, start: i, end: i + 1 });
        }
    }
    Logger.log('returning urls');
    return urls;
}

/**
 * Checks if two arrays are equal.
 *
 * @param {Array} arr1 - The first array.
 * @param {Array} arr2 - The second array.
 * @return {boolean} True if the arrays are equal, false otherwise.
 */
function arraysEqual(arr1, arr2) {
    Logger.log('arraysEqual triggered');
    if (arr1.length !== arr2.length) return false;
    for (let i = 0; i < arr1.length; i++) {
        if (arr1[i] !== arr2[i]) return false;
    }
    Logger.log('arrays are equal');
    return true;
}

/**
 * Generates a SHA-256 hash for the given content.
 *
 * @param {string} content - The content to hash.
 * @return {string} The generated hash in base64 encoding.
 */
function generateHash(content) {
    Logger.log('generateHash triggered');
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, content));
    Logger.log('hash generated');
    return hash;
}

/**
 * Checks if the hash of the content of the sheet has changed.
 *
 * @param {string} lastHash - The previous hash value.
 * @param {string} currentHash - The current hash value.
 * @return {boolean} True if the hash has changed, false otherwise.
 */
function shouldRunUpdates(lastHash, currentHash) {
    Logger.log('shouldRunUpdates triggered');
    const hasChanged = lastHash !== currentHash;
    Logger.log(`hash has changed: ${hasChanged}`);
    return hasChanged;
}

/**
 * Gets the content of the sheet and generates a hash for it.
 *
 * @return {string} The generated hash of the sheet content.
 */
function getSheetContentHash() {
    Logger.log('getSheetContentHash triggered');
    const range = getDataRange();
    const values = range.getValues().flat().join(",");
    Logger.log('getSheetContentHash: returning generateHash');
    const hash = generateHash(values);
    Logger.log(`generated hash: ${hash}`);
    return hash;
}

/**
 * Saves a snapshot of the current state of the active sheet.
 * The snapshot includes the text content and links of each cell.
 * You can specify cells to ignore by passing an array of cell references.
 * 
 * @param {Array<string>} cellsToIgnore - (e.g., ["R1C3", "R1C4", "R1C5"] for C1, D1, E1).
 * @return {object} The snapshot object.
 */
function saveSnapshot(cellsToIgnore = []) {
    Logger.log('saveSnapshot triggered');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getDataRange();
    const richTextValues = range.getRichTextValues();
    const snapshot = {};

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellKey = `R${row + 1}C${col + 1}`;
            if (cellsToIgnore.includes(cellKey)) {
                Logger.log(`Ignoring cell ${cellKey} from snapshot.`);
                continue;
            }

            const cellValue = richTextValues[row][col];
            if (cellValue) {
                const urls = extractUrls(cellValue);
                snapshot[cellKey] = {
                    text: cellValue.getText(),
                    links: urls
                };
                Logger.log(`Snapshot saved for cell ${cellKey}.`);
            }
        }
    }

    // Save snapshot to script properties
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('sheetSnapshot', JSON.stringify(snapshot));
    Logger.log("Snapshot saved.");
    return snapshot;
}

/**
 * Restores the sheet to a previously saved snapshot state.
 * This includes restoring text content, links, and optional custom formatting.
 *
 * @param {function} formatCallback - Optional callback function to apply custom formatting.
 * @return {void}
 */
function restoreSnapshot(formatCallback) {
    Logger.log('restoreSnapshot triggered');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getDataRange();
    const properties = PropertiesService.getScriptProperties();
    const snapshotJson = properties.getProperty('sheetSnapshot');

    if (!snapshotJson) {
        Logger.log("No snapshot found.");
        return;
    }

    const snapshot = JSON.parse(snapshotJson);
    const richTextValues = range.getRichTextValues();

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellKey = `R${row + 1}C${col + 1}`;
            if (snapshot[cellKey]) {
                const cellData = snapshot[cellKey];
                const builder = SpreadsheetApp.newRichTextValue()
                    .setText(cellData.text);
                Logger.log(`Restoring snapshot for cell ${cellKey}.`);
                // Restore links
                for (const link of cellData.links) {
                    Logger.log(`Restoring link: ${link.url} at ${link.start}-${link.end}.`);
                    builder.setLinkUrl(link.start, link.end, link.url);
                }
                Logger.log(`Restored links: ${cellData.links.length}.`);
                // Apply custom formatting if a callback is provided
                if (formatCallback) {
                    Logger.log(`restoreSnapshot()/formatCallback(): Applying custom formatting for cell ${cellKey}.`);
                    formatCallback(builder, cellData.text);
                }
                richTextValues[row][col] = builder.build();
            }
        }
    }

    range.setRichTextValues(richTextValues);
    Logger.log("Snapshot restored.");
}

/**
 * Iterates over each cell in the selected range and applies the specified function to it.
 * @param {function(Range, RichTextValue): void} cellFunction - The function to apply to each cell.
 */
function processCells(cellFunction) {
    Logger.log('processCells triggered');
    const range = sheet.getActiveRange();
    const richTextValues = range.getRichTextValues();

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellValue = richTextValues[row][col];
            if (cellValue) {
                const cellRange = range.getCell(row + 1, col + 1);
                cellFunction(cellRange, cellValue);
            }
        }
    }
    Logger.log('processCells completed');
}

/**
 * Preserves existing text styles and links from the original text to the new text.
 * @param {RichTextValue} originalTextValue - The original rich text value.
 * @param {RichTextValueBuilder} newTextValueBuilder - The rich text value builder for the new text.
 * @param {number} offset - The offset to apply to the new text positions.
 */
function preserveStylesAndLinks(originalTextValue, newTextValueBuilder, offset) {
    Logger.log('preserveStylesAndLinks triggered');
    const originalText = originalTextValue.getText();
    const newText = newTextValueBuilder.build().getText();
    const minLength = Math.min(originalText.length, newText.length - offset);

    for (let i = 0; i < minLength; i++) {
        const style = originalTextValue.getTextStyle(i, i + 1);
        newTextValueBuilder.setTextStyle(i + offset, i + offset + 1, style);

        const url = originalTextValue.getLinkUrl(i, i + 1);
        if (url) {
            newTextValueBuilder.setLinkUrl(i + offset, i + offset + 1, url);
        }
    }
    Logger.log('preserveStylesAndLinks completed');
}



// Contents of ./TODOsheet/TODOcheckbox.js



// globals.js: sheet
// shared/utils.js: processCells, preserveStylesAndLinks

/**
 * Adds by default a checkbox to a cell while preserving existing rich text styles and links.
 * @param {Range} range - The range of the cell to which the checkbox is added.
 * @customfunction
 * @returns {void}
 */
function addCheckboxToCellTODO(range) {
    Logger.log('addCheckboxToCellTODO triggered');
    const cellValue = range.getValue().toString();
    const richTextValue = range.getRichTextValue() || SpreadsheetApp.newRichTextValue().setText(cellValue).build();

    if (cellValue.startsWith('☑️') || cellValue.startsWith('✅')) {
        Logger.log(`Checkbox already present at the start of cell ${range.getA1Notation()}`);
        return;
    }

    const newRichTextValueBuilder = SpreadsheetApp.newRichTextValue().setText('☑️' + cellValue);
    newRichTextValueBuilder.setTextStyle(0, 2, SpreadsheetApp.newTextStyle().setBold(true).build());

    preserveStylesAndLinks(richTextValue, newRichTextValueBuilder, 2);

    range.setRichTextValue(newRichTextValueBuilder.build());
    Logger.log(`Checkbox added to the start of cell ${range.getA1Notation()}`);
}

/**
 * Adds a checkbox to all selected cells, preserving existing rich text styles and links.
 * @customfunction
 * @returns {void}
 */
function addCheckboxesTODO() {
    Logger.log('addCheckboxesTODO triggered');
    processCells((cellRange, cellValue) => {
        const originalText = cellValue.getText();
        Logger.log(`Original cell text: "${originalText}"`);

        const onlyDefaultCheckbox = originalText === '☑️';
        let newText = onlyDefaultCheckbox ? '☑️☑️' : `☑️${originalText}`;
        Logger.log(`New text: "${newText}"`);

        const builder = SpreadsheetApp.newRichTextValue().setText(newText);
        preserveStylesAndLinks(cellValue, builder, onlyDefaultCheckbox ? 1 : 2);

        cellRange.setRichTextValue(builder.build());
    });
    Logger.log("Checkboxes added to selected cells.");
}

/**
 * Changes the first checkbox in each selected cell to a green checkbox.
 * @customfunction
 * @returns {void}
 */
function markCheckboxTODO() {
    Logger.log('markCheckboxTODO triggered');
    processCells((cellRange, cellValue) => {
        let newText = cellValue.getText();
        const firstCheckboxIndex = newText.indexOf('☑️');
        if (firstCheckboxIndex !== -1) {
            newText = newText.substring(0, firstCheckboxIndex) + '✅' + newText.substring(firstCheckboxIndex + 2);

            const builder = SpreadsheetApp.newRichTextValue().setText(newText);
            preserveStylesAndLinks(cellValue, builder, 0);

            cellRange.setRichTextValue(builder.build());
            Logger.log(`First checkbox in cell ${cellRange.getA1Notation()} changed to green.`);
        }
    });
    Logger.log("One checkbox changed to green in selected cells.");
}

/**
 * Changes all checkboxes in each selected cell to green checkboxes.
 * @customfunction
 * @returns {void}
 */
function markAllCheckboxesTODO() {
    Logger.log('markAllCheckboxesTODO triggered');
    processCells((cellRange, cellValue) => {
        let newText = cellValue.getText().replace(/☑️/g, '✅');
        Logger.log(`New text after changing all checkboxes to green: "${newText}"`);

        const builder = SpreadsheetApp.newRichTextValue().setText(newText);
        preserveStylesAndLinks(cellValue, builder, 0);

        cellRange.setRichTextValue(builder.build());
    });
    Logger.log("All checkboxes changed to green in selected cells.");
}

/**
 * Restores all checkboxes in selected cells to their default state.
 * Changes green checkboxes back to default checkboxes while preserving styles and links.
 * @customfunction
 * @returns {void}
 */
function restoreCheckboxesTODO() {
    Logger.log("restoreCheckboxesTODO triggered");
    processCells((cellRange, cellValue) => {
        let newText = cellValue.getText().replace(/✅/g, '☑️');
        Logger.log(`Updated text for cell ${cellRange.getA1Notation()}: ${newText}`);

        const builder = SpreadsheetApp.newRichTextValue().setText(newText);
        preserveStylesAndLinks(cellValue, builder, 0);

        cellRange.setRichTextValue(builder.build());
    });
    Logger.log("All checkboxes restored to default in selected cells.");
}

/**
 * Removes all checkboxes from the selected cells while preserving existing rich text styles and links.
 * @customfunction
 * @returns {void}
 */
function removeCheckboxesTODO() {
    Logger.log("removeCheckboxesTODO triggered");
    processCells((cellRange, cellValue) => {
        let newText = cellValue.getText().replace(/☑️|✅/g, '');
        Logger.log(`Text after checkbox removal for cell ${cellRange.getA1Notation()}: "${newText}"`);

        const builder = SpreadsheetApp.newRichTextValue().setText(newText);
        preserveStylesAndLinks(cellValue, builder, 0);

        cellRange.setRichTextValue(builder.build());
    });
    Logger.log("Checkboxes removed from selected cells.");
}

// for testing

// Contents of ./TODOsheet/TODOformatting.js

// globals.js: sheet, getDataRange, datePattern
// shared/formatting.js: Format, applyBorders, applyThickBorders, setCellStyle, appendDateWithStyle, updateDateWithStyle, resetTextStyle, clearTextFormatting
// TODOsheet/TODOtimeHandle.js: updateDaysLeftCellTODO
// TODOsheet/TODOlibrary.js: dateColorConfig

/**
 * Updates the comment for a specific cell with version and feature details.
 * 
 * @customfunction
 */
function updateCellCommentTODO() {
    Logger.log('updateCellCommentTODO called');
    const cell = sheet.getRange("I2");
    const version = "v1.2";
    const emoji = "💡";
    const changes = `\n
        - A checkbox is added by default from the 3rd to the 8th column when a cell is written or modified.\n
        - You can add, mark, restore and delete checkboxes in cells by selecting them and using the "Custom Formats" menu.\n
        - The "days left" counter is updated daily in the 8th column. When the counter reaches zero, the cell is cleared.\n
        - A snapshot of the sheet can be saved and restored from the "Custom Formats" menu.\n
        - Snapshots are automatically saved and restored when the sheet is reloaded so that the last state is always preserved.\n

        OLD FEATURES: \n
        - There is an indicative limit of cells for each priority. In the end the objective of a TODO is none other than to complete the tasks and that they do not accumulate. Once this limit is reached, a warning is activated for the entire column.
        This feature does not block cells, that is, you can continue occupying cells even if you have the warning.\n
        - You can apply some custom formats that do not require to refresh the page from the "Custom Formats" menu.\n
        - The date color change times are different for each column, with HIGH PRIORITY being the fastest to change and LOW PRIORITY being the slowest.\n
        - The Piechart can be shown or hidden directly using its dropdown cell.\n
        - Empty cells that are deleted are occupied by their immediately lower cell.\n
    `;

    const comment = `Versión: ${version}\n NEW FEATURES:\n${changes}`;
    cell.setComment(comment);
    cell.setBackground("#efefef");
    cell.setBorder(true, true, true, true, true, true, '#D3D3D3', SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Create RichTextValue with different font sizes
    const richText = SpreadsheetApp.newRichTextValue()
        .setText(`${version}\n${emoji}`)
        .setTextStyle(0, version.length, SpreadsheetApp.newTextStyle().setFontSize(8).build())
        .setTextStyle(version.length + 1, version.length + 2, SpreadsheetApp.newTextStyle().setFontSize(20).build())
        .setTextStyle(version.length + 2, version.length + 3, SpreadsheetApp.newTextStyle().setFontSize(20).build())
        .build();

    cell.setRichTextValue(richText);
    Format(cell);
}

/**
 * Sets example text for a specific column if the cells are empty.
 * 
 * @customfunction
 * @param {string} column - The column to check for empty cells.
 * @param {string} exampleText - The example text to set if cells are empty.
 */
function exampleTextTODO(column, exampleText) {
    Logger.log(`exampleTextTODO called for column: ${column}, example text: ${exampleText}`);
    const dataRange = getDataRange();
    const lastRow = dataRange.getLastRow();  // Get the last row with data
    let values;

    if (column === "B") {
        // Get values excluding B3 and B8
        const firstPart = sheet.getRange(column + "2").getValues().flat();  // Get value from B2
        const middlePart = sheet.getRange(column + "4:" + column + "7").getValues().flat();  // Get values from B4-B7
        const lastPart = sheet.getRange(column + "9:" + column + lastRow).getValues().flat();  // Get values from B9 to the last row

        values = [...firstPart, ...middlePart, ...lastPart];
    } else {
        values = sheet.getRange(column + "2:" + column + lastRow).getValues().flat();  // Get values from the column's 2nd row to the last row
    }

    Logger.log(`Values in column ${column}: ${values}`);

    // Check if the first cell of the column is empty
    const firstCellEmpty = values[0].toString().trim() === '';

    if (firstCellEmpty) {
        const cell = sheet.getRange(column + "2");
        cell.setValue(exampleText);  // Set example text if the first cell is empty
        Logger.log(`Example text set for column ${column} at ${column}2: ${exampleText}`);
    } else {
        Logger.log(`Column ${column} is not empty at ${column}2, skipping setting example text.`);
    }
}

/**
 * Applies formatting to the entire sheet and sets example text.
 * 
 * @customfunction
 */
function applyFormatToAllTODO() {
    Logger.log('applyFormatToAllTODO called');
    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';
    const totalRows = sheet.getMaxRows();  // Get the total number of rows
    let range = sheet.getRange(1, 1, totalRows, 8);  // Define the range covering all rows and 8 columns
    if (range) {
        Format(range);  // Apply formatting to the range
        applyBorders(range);  // Apply borders to the range
    }

    Logger.log('applyFormatToAllTODO()/applyThickBorders(): applying thick borders');
    applyThickBorders(sheet.getRange(1, 3, 11, 1));  // Apply thick borders to a specific range
    applyThickBorders(sheet.getRange(1, 4, 21, 1));  // Apply thick borders to another range
    applyThickBorders(sheet.getRange(1, 5, 21, 1));  // Apply thick borders to yet another range

    Logger.log('applyFormatToAllTODO()/setCellContentAndStyle(): setting cell content and style');
    setCellContentAndStyleTODO();  // Set cell content and styles

    Logger.log('applyFormatToAllTODO()/checkAndSetColumnTODO(): checking and setting columns');
    for (const column in cellStyles) {
        const { limit, priority, value } = cellStyles[column];

        // Validate if the limit and priority are available in the selected language
        const translatedLimit = limit?.[language];
        const translatedPriority = priority?.[language];

        if (translatedLimit !== undefined && translatedPriority !== undefined) {
            checkAndSetColumnTODO(column.charAt(0), translatedLimit, translatedPriority);  // Apply column-specific settings
            Logger.log(`applyFormatToAllTODO(): translatedText set for column ${column} - limit: ${translatedLimit}, priority: ${translatedPriority}`);
        } else {
            Logger.log(`applyFormatToAllTODO(): limit or priority not found for column ${column} and language ${language}`);
        }
    }

    Logger.log('applyFormatToAllTODO()/exampleTextTODO(): setting example text');
    for (const column in exampleTexts) {
        const { text } = exampleTexts[column];
        const translatedText = text[language];  // Get the example text based on the selected language
        exampleTextTODO(column, translatedText);  // Set example text for the column
        Logger.log(`applyFormatToAllTODO(): example text set for column ${column} - translatedText: ${translatedText}`);
    }
}

/**
 * Checks and sets the column based on the limit of occupied cells.
 * 
 * @customfunction
 * @param {string} column - The column to check.
 * @param {number} limit - The limit of occupied cells.
 * @param {string} priority - The priority level.
 */
function checkAndSetColumnTODO(column, limit, priority) {
    Logger.log(`checkAndSetColumnTODO called for column: ${column}, limit: ${limit}, priority: ${priority}`);
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

/**
 * Sets the background color of a specific column.
 * 
 * @customfunction
 * @param {Sheet} sheet - The sheet object.
 * @param {number} col - The column number.
 * @param {string} color - The background color to set.
 * @param {number} [startRow=2] - The starting row number.
 */
function setColumnBackground(sheet, col, color, startRow = 2) {
    Logger.log(`setColumnBackground called for column: ${col}, color: ${color}, startRow: ${startRow}`);
    let totalRows = sheet.getMaxRows();
    let range = sheet.getRange(startRow, col, totalRows - startRow + 1, 1);
    range.setBackground(color);
}

/**
 * Customizes the background colors of specific columns and cells.
 * 
 * @customfunction
 */
function customCellBGColorTODO() {
    Logger.log('customCellBGColorTODO called');
    // Apply background colors to specific columns
    setColumnBackground(sheet, 1, '#d3d3d3', 2); // Column A: Light gray 3
    setColumnBackground(sheet, 6, '#fff1f1', 2); // Column F: Light pink
    setColumnBackground(sheet, 7, '#d3d3d3', 2); // Column G: Light gray 3

    // Apply white background to columns B, C, D, E, H, I starting from row 2
    let whiteColumns = [2, 3, 4, 5, 8, 9]; // Makes cell I2 momentarily white(column 8) while loading rest of the sheet. Useful for testing. Then turns dark gray(updateCellCommentTODO)
    for (let col of whiteColumns) {
        setColumnBackground(sheet, col, '#ffffff', 2);
    }

    // Apply dark yellow background to specific cells in column B
    sheet.getRange('B3').setBackground('#b5a642'); // Dark yellow 3
    sheet.getRange('B8').setBackground('#b5a642'); // Dark yellow 3
}

/**
 * Sets content and style for specific cells based on predefined configurations.
 * 
 * @customfunction
 */
function setCellContentAndStyleTODO() {
    Logger.log('setCellContentAndStyleTODO called');
    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';
    for (const cell in cellStyles) {
        const { value, fontWeight, fontColor, backgroundColor, alignment } = cellStyles[cell];
        const translatedValue = value[language];
        setCellStyle(cell, translatedValue, fontWeight, fontColor, backgroundColor, alignment);
    }
}


/**
 * Sets up a dropdown menu in cell I1 with options to show or hide the pie chart.
 *
 * @customfunction
 */
function setupDropdownTODO() {
    Logger.log('setupDropdownTODO called');
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

/**
 * Shifts cells up in a column if they are empty, filling with the values below.
 *
 * @customfunction
 * @param {number} column - The column to shift cells up (1-indexed).
 * @param {number} startRow - The starting row number.
 * @param {number} endRow - The ending row number.
 */
function shiftCellsUpTODO(column, startRow, endRow) {
    Logger.log(`shiftCellsUpTODO called for column: ${column}, from row ${startRow} to ${endRow}`);

    const range = sheet.getRange(startRow, column, endRow - startRow + 1, 1);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();

    let hasChanges = false;

    for (let i = 0; i < values.length - 1; i++) {
        if (values[i][0] === '' && values[i + 1][0] !== '') {
            Logger.log(`Empty cell found at row ${i + startRow}, shifting cells up`);

            // Preserve the original rich text, including links
            values[i][0] = values[i + 1][0];
            richTextValues[i][0] = richTextValues[i + 1][0];

            values[i + 1][0] = '';
            richTextValues[i + 1][0] = SpreadsheetApp.newRichTextValue().setText('').build();

            hasChanges = true;
            Logger.log(`After shifting: Row ${i + startRow}, New Value: ${values[i][0]}, New RichText: ${richTextValues[i][0].getText()}`);
        }
    }

    if (hasChanges) {
        Logger.log(`Setting values for range: ${startRow} to ${endRow}, column: ${column}`);
        range.setValues(values);
        range.setRichTextValues(richTextValues);
    }
    Logger.log(`shiftCellsUpTODO completed for column: ${column}`);
}

/**
 * Forces empty cells to shift up in specified columns.
 *
 * @customfunction
 */
function pushUpEmptyCellsTODO() {
    Logger.log('pushUpEmptyCellsTODO called');
    const range = sheet.getDataRange();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();

    for (let col = 1; col <= numCols; col++) {
        let startRow = null;
        for (let row = 2; row <= numRows; row++) {
            if (sheet.getRange(row, col).getValue() === '' && startRow === null) {
                startRow = row;
            } else if (sheet.getRange(row, col).getValue() !== '' && startRow !== null) {
                shiftCellsUpTODO(col, startRow, row - 1);
                startRow = null;
            }
        }
        // Handle the case where the last rows are empty
        if (startRow !== null) {
            shiftCellsUpTODO(col, startRow, numRows);
        }
    }
    Logger.log('pushUpEmptyCells completed');
}

/**
 * Updates rich text content of a cell based on original and new values.
 *
 * @customfunction
 * @param {Range} range - The cell range to update.
 * @param {string} originalValue - The original value of the cell.
 * @param {string} newValue - The new value of the cell.
 * @param {string} columnLetter - The column letter of the cell.
 * @param {number} row - The row number of the cell.
 * @param {Event} e - The edit event object.
 */
function updateRichTextTODO(range, originalValue, newValue, columnLetter, row, e) {
    Logger.log(`updateRichTextTODO called for column: ${columnLetter}, row: ${row}, original value: "${originalValue}", new value: "${newValue}"`);

    let updatedText = newValue.toString().trim();
    const dateFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy");
    Logger.log(`Today's date: ${dateFormatted}`);

    const originalRichTextValue = range.getRichTextValue() || SpreadsheetApp.newRichTextValue().setText(originalValue).build();

    // Reemplazar el patrón (n) days left por la fecha actual
    const daysLeftPattern = /\(\d+\) days left/;
    if (daysLeftPattern.test(updatedText)) {
        updatedText = updatedText.replace(daysLeftPattern, dateFormatted);
        Logger.log(`Replaced (n) days left with today's date: "${updatedText}"`);
    } else if (!/\d{2}\/\d{2}\/\d{2}/.test(updatedText)) {
        updatedText = updatedText + '\n' + dateFormatted;
        Logger.log(`No date found, updated text with new date: "${updatedText}"`);
    } else {
        updatedText = updatedText.replace(/\d{2}\/\d{2}\/\d{2}/, dateFormatted);
        Logger.log(`Replaced date with new date: "${updatedText}"`);
    }

    Logger.log(`Updated text: "${updatedText}"`);

    const newRichTextValueBuilder = SpreadsheetApp.newRichTextValue()
        .setText(updatedText)
        .setTextStyle(0, updatedText.length, SpreadsheetApp.newTextStyle().build());

    const lastLineIndex = updatedText.lastIndexOf('\n');
    Logger.log(`Last line index: ${lastLineIndex}`);
    if (lastLineIndex !== -1) {
        const color = columnLetter === 'H' ? '#FF0000' : '#A9A9A9';
        newRichTextValueBuilder.setTextStyle(
            lastLineIndex + 1,
            updatedText.length,
            SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor(color).build()
        );
        Logger.log(`Applied style to last line: ${lastLineIndex + 1} to ${updatedText.length}`);
    }

    const originalText = originalRichTextValue.getText();
    Logger.log(`Preserving links from original text: ${originalText}`);
    for (let i = 0; i < Math.min(lastLineIndex !== -1 ? lastLineIndex : updatedText.length, originalText.length); i++) {
        const url = originalRichTextValue.getLinkUrl(i, i + 1);
        if (url) {
            newRichTextValueBuilder.setLinkUrl(i, i + 1, url);
        }
        Logger.log(`Preserved link for index: ${i}`);
    }

    range.setRichTextValue(newRichTextValueBuilder.build());
    Logger.log(`Set new rich text value for cell ${columnLetter}${row}`);
}

/**
 * Handles the editing of a cell based on its column.
 * 
 * @param {Range} range - The cell range that was edited.
 * @param {string} originalValue - The original value of the cell before editing.
 * @param {string} newValue - The new value of the cell after editing.
 * @param {string} columnLetter - The letter of the column that was edited.
 * @param {number} row - The row number of the edited cell.
 * @param {Event} e - The edit event object.
 */
function handleColumnEditTODO(range, originalValue, newValue, columnLetter, row, e) {
    Logger.log(`handleColumnEditTODO called for column: ${columnLetter}, row: ${row}, original value: "${originalValue}", new value: "${newValue}"`);
    if (columnLetter === 'H') {
        let daysLeft = parseDaysLeftTODO(newValue);
        updateDaysLeftCellTODO(range, daysLeft);
    } else {
        updateRichTextTODO(range, originalValue, newValue, columnLetter, row, e);
        removeMultipleDatesTODO();
    }
}

// for testing

// Contents of ./TODOsheet/TODOlibrary.js

const cellStyles = {
    "A1": {
        value: {
            "English": "BEHAVIOR PATTERNS",
            "Spanish": "PATRONES DE CONDUCTA",
            "Catalan": "PATRONS DE CONDUCTA"
        },
        fontWeight: "bold",
        fontColor: "#FFFFFF",
        backgroundColor: "#000000",
        alignment: "center"
    },
    "B1": {
        value: {
            "English": "TOMORROW",
            "Spanish": "MAÑANA",
            "Catalan": "DEMÀ"
        },
        fontWeight: "bold",
        fontColor: "#FFFFFF",
        backgroundColor: "#b5a642",
        alignment: "center"
    },
    "B3": {
        value: {
            "English": "WEEK",
            "Spanish": "SEMANA",
            "Catalan": "SETMANA"
        },
        fontWeight: "bold",
        fontColor: "#FFFFFF",
        backgroundColor: "#b5a642",
        alignment: "center"
    },
    "B8": {
        value: {
            "English": "MONTH",
            "Spanish": "MES",
            "Catalan": "MES"
        },
        fontWeight: "bold",
        fontColor: "#FFFFFF",
        backgroundColor: "#b5a642",
        alignment: "center"
    },
    "F1": {
        value: {
            "English": "IDEAS AND PLANS",
            "Spanish": "IDEAS Y PLANES",
            "Catalan": "IDEES I PLANS"
        },
        fontWeight: "bold",
        fontColor: "#000000",
        backgroundColor: "#FFC0CB",
        alignment: "center"
    },
    "G1": {
        value: {
            "English": "EYES ON",
            "Spanish": "ATENTO A",
            "Catalan": "ATENT A"
        },
        fontWeight: "bold",
        fontColor: "#000000",
        backgroundColor: "#b7b7b7",
        alignment: "center"
    },
    "H1": {
        value: {
            "English": "IN QUARANTINE",
            "Spanish": "EN CUARENTENA",
            "Catalan": "EN QUARANTENA"
        },
        fontWeight: "bold",
        fontColor: "#FF0000",
        backgroundColor: null,
        alignment: "center"
    },
    "C1": {
        value: {
            "English": "HIGH PRIORITY",
            "Spanish": "PRIORIDAD ALTA",
            "Catalan": "PRIORITAT ALTA"
        },
        limit: {
            "English": 10,
            "Spanish": 10,
            "Catalan": 10
        },
        priority: {
            "English": "HIGH PRIORITY",
            "Spanish": "PRIORIDAD ALTA",
            "Catalan": "PRIORITAT ALTA"
        },
        fontWeight: "bold",
        fontColor: null,
        backgroundColor: "#fce5cd",
        alignment: "center"
    },
    "D1": {
        value: {
            "English": "MEDIUM PRIORITY",
            "Spanish": "PRIORIDAD MEDIA",
            "Catalan": "PRIORITAT MITJANA"
        },
        limit: {
            "English": 20,
            "Spanish": 20,
            "Catalan": 20
        },
        priority: {
            "English": "MEDIUM PRIORITY",
            "Spanish": "PRIORIDAD MEDIA",
            "Catalan": "PRIORITAT MITJANA"
        },
        fontWeight: "bold",
        fontColor: null,
        backgroundColor: "#fff2cc",
        alignment: "center"
    },
    "E1": {
        value: {
            "English": "LOW PRIORITY",
            "Spanish": "BAJA PRIORIDAD",
            "Catalan": "BAIXA PRIORITAT"
        },
        limit: {
            "English": 20,
            "Spanish": 20,
            "Catalan": 20
        },
        priority: {
            "English": "LOW PRIORITY",
            "Spanish": "BAJA PRIORIDAD",
            "Catalan": "BAIXA PRIORITAT"
        },
        fontWeight: "bold",
        fontColor: null,
        backgroundColor: "#d9ead3",
        alignment: "center"
    }
};


const exampleTexts = {
    "A": {
        text: {
            "English": "Example: Do it with fear but do it.",
            "Spanish": "Ejemplo: Hazlo con miedo pero hazlo.",
            "Catalan": "Exemple: Fes-ho si cal amb por però fes-ho."
        },
        color: "#FFFFFF"
    },
    "B": {
        text: {
            "English": "Example: 45min of cardio",
            "Spanish": "Ejemplo: 45min de cardio",
            "Catalan": "Exemple: 45min de cardio"
        },
        color: "#A9A9A9"
    },
    "C": {
        text: {
            "English": "Example: Join that gym club",
            "Spanish": "Ejemplo: Apuntate al gym",
            "Catalan": "Exemple: Apunta't al gym"
        },
        color: "#A9A9A9"
    },
    "D": {
        text: {
            "English": "Example: Submit that pending data science task.",
            "Spanish": "Ejemplo: Entrega esa tarea pendiente de ciencia de datos.",
            "Catalan": "Exemple: Lliura aquella tasca pendent de ciència de dades."
        },
        color: "#A9A9A9"
    },
    "E": {
        text: {
            "English": "Example: Buy a new mattress.",
            "Spanish": "Ejemplo: Compra un nuevo colchón.",
            "Catalan": "Exemple: Compra un nou matalàs."
        },
        color: "#A9A9A9"
    },
    "F": {
        text: {
            "English": "Example: Santiago route.",
            "Spanish": "Ejemplo: Ruta de Santiago.",
            "Catalan": "Exemple: Ruta de Santiago."
        },
        color: "#A9A9A9"
    },
    "G": {
        text: {
            "English": "Example: Change front brake pad at 44500km",
            "Spanish": "Ejemplo: Cambia la pastilla de freno delantera a los 44500km",
            "Catalan": "Exemple: Canvia la pastilla de fren davanter als 44500km"
        },
        color: "#FFFFFF"
    },
    "H": {
        text: {
            "English": "Example: Join that Crossfit club",
            "Spanish": "Ejemplo: Únete al club de Crossfit",
            "Catalan": "Exemple: Uneix-te al club de Crossfit"
        },
        color: "#A9A9A9"
    }
};


const dateColorConfig = {
    C: { warning: 7, danger: 30, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' }, // 1 week, 1 month
    D: { warning: 90, danger: 180, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
    E: { warning: 180, danger: 365, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
    F: { warning: 180, danger: 365, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
    G: { warning: 0, danger: 0, warningColor: '#A9A9A9', dangerColor: '#A9A9A9', defaultColor: '#A9A9A9' }, // Always default
    H: { warning: 0, danger: 0, warningColor: '#FF0000', dangerColor: '#FF0000', defaultColor: '#FF0000' } // Always red
};

const languages = {
    English: 'English',
    Spanish: 'Spanish',
    Catalan: 'Catalan'
};

const menuLanguage = [
    {
        title: {
            English: 'Language',
            Spanish: 'Idioma',
            Catalan: 'Idioma'
        },
        items: {
            setLanguageEnglish: {
                English: 'English',
                Spanish: 'Inglés',
                Catalan: 'Anglès'
            },
            setLanguageSpanish: {
                English: 'Spanish',
                Spanish: 'Español',
                Catalan: 'Espanyol'
            },
            setLanguageCatalan: {
                English: 'Catalan',
                Spanish: 'Catalán',
                Catalan: 'Català'
            }
        }
    }
]
const menuTodoSheet = [
    {
        title: {
            English: 'TODO sheet',
            Spanish: 'Hoja TODO',
            Catalan: 'Full de TODO'
        },
        items: {
            restoreDefaultTodoTemplate: {
                English: 'RESTORE DEFAULT TODO TEMPLATE',
                Spanish: 'RESTAURAR PLANTILLA POR DEFECTO',
                Catalan: 'RESTAURAR PLANTILLA PER DEFECTE'
            },
            restoreCellBackgroundColors: {
                English: 'RESTORE Cell Background Colors',
                Spanish: 'RESTAURAR Colores de Fondo de Celda',
                Catalan: 'RESTAURAR Colors de Fons de Cel·la'
            },
            addCheckboxesToSelectedCells: {
                English: 'Add Checkboxes to Selected Cells',
                Spanish: 'Añadir Casillas a las Celdas Seleccionadas',
                Catalan: 'Afegir Caselles a les Cel·les Seleccionades'
            },
            markCheckboxInSelectedCells: {
                English: 'Mark Checkbox in Selected Cells',
                Spanish: 'Marcar Casilla en las Celdas Seleccionadas',
                Catalan: 'Marcar Casella a les Cel·les Seleccionades'
            },
            markAllCheckboxesInSelectedCells: {
                English: 'Mark All Checkboxes in Selected Cells',
                Spanish: 'Marcar Todas las Casillas en las Celdas Seleccionadas',
                Catalan: 'Marcar Totes les Caselles a les Cel·les Seleccionades'
            },
            restoreCheckboxes: {
                English: 'Restore Checkboxes',
                Spanish: 'Restaurar Casillas',
                Catalan: 'Restaurar Caselles'
            },
            removeAllCheckboxesInSelectedCells: {
                English: 'Remove All Checkboxes in Selected Cells',
                Spanish: 'Eliminar Todas las Casillas en las Celdas Seleccionadas',
                Catalan: 'Eliminar Totes les Caselles a les Cel·les Seleccionades'
            },
            saveSnapshot: {
                English: 'Save Snapshot',
                Spanish: 'Guardar Instantánea',
                Catalan: 'Guardar Instantània'
            },
            restoreSnapshot: {
                English: 'Restore Snapshot',
                Spanish: 'Restaurar Instantánea',
                Catalan: 'Restaurar Instantània'
            },
            createPieChart: {
                English: 'Create Pie Chart',
                Spanish: 'Crear Gráfico Circular',
                Catalan: 'Crear Gràfic Circular'
            },
            deletePieCharts: {
                English: 'Delete Pie Charts',
                Spanish: 'Eliminar Gráficos Circulares',
                Catalan: 'Eliminar Gràfics Circulars'
            },
            versionAndFeatureDetails: {
                English: 'Version and feature details',
                Spanish: 'Detalles de Versión y Funcionalidades',
                Catalan: 'Detalls de Versió i Funcionalitats'
            },
            logHelloWorld: {
                English: 'Log Hello World',
                Spanish: 'Registrar Hola Mundo',
                Catalan: 'Registrar Hola Món'
            }
        }
    }]

const menuCustomFormats = [
    {
        title: {
            English: 'Custom Formats',
            Spanish: 'Formatos Personalizados',
            Catalan: 'Formats Personalitzats'
        },
        items: {
            applyFormat: {
                English: 'Apply Format',
                Spanish: 'Aplicar Formato',
                Catalan: 'Aplicar Format'
            },
            applyFormatToAll: {
                English: 'Apply Format to All',
                Spanish: 'Aplicar Formato a Todo',
                Catalan: 'Aplicar Format a Tot'
            }
        }
    }]

const menus = [
    {
        config: menuTodoSheet,
        items: [
            { key: 'restoreDefaultTodoTemplate', separatorAfter: false },
            { key: 'restoreCellBackgroundColors', separatorAfter: true },
            { key: 'addCheckboxesToSelectedCells', separatorAfter: false },
            { key: 'markCheckboxInSelectedCells', separatorAfter: false },
            { key: 'markAllCheckboxesInSelectedCells', separatorAfter: false },
            { key: 'restoreCheckboxes', separatorAfter: false },
            { key: 'removeAllCheckboxesInSelectedCells', separatorAfter: true },
            { key: 'saveSnapshot', separatorAfter: false },
            { key: 'restoreSnapshot', separatorAfter: true },
            { key: 'createPieChart', separatorAfter: false },
            { key: 'deletePieCharts', separatorAfter: true },
            { key: 'versionAndFeatureDetails', separatorAfter: false },
            { key: 'logHelloWorld', separatorAfter: false }
        ],
        suffix: ''
    },
    {
        config: menuCustomFormats,
        items: [
            { key: 'applyFormat', separatorAfter: false },
            { key: 'applyFormatToAll', separatorAfter: false }
        ],
        suffix: ''
    },
    {
        config: menuLanguage,
        items: [
            { key: 'setLanguageEnglish', separatorAfter: false },
            { key: 'setLanguageSpanish', separatorAfter: false },
            { key: 'setLanguageCatalan', separatorAfter: false }
        ],
        suffix: ''
    }
];

const toastMessages = {
    loading: {
        English: 'Data is loading...\n Please wait.',
        Spanish: 'Cargando datos...\n Por favor espera.',
        Catalan: "S'estan carregant les dades...\n Si us plau, espera."
    },
    updateComplete: {
        English: 'Update Complete!',
        Spanish: 'Actualización completada!',
        Catalan: 'Actualització completada!'
    }
};



// Contents of ./TODOsheet/TODOmenu.js

// globals.js: ui

/**
 * Gets the current language from document properties or returns 'English' as default.
 * @returns {string} The current language
 */
const getCurrentLanguageTODO = () => PropertiesService.getDocumentProperties().getProperty('language') || 'English';

/**
 * Creates custom menus in the spreadsheet.
 * Adds menu items to the UI and assigns functions to them.
 * 
 * @customfunction
 */
function createMenusTODO() {
    Logger.log('createMenusTODO triggered');
    const currentLanguage = getCurrentLanguageTODO();

    const functionNameMap = {
        'restoreDefaultTodoTemplate': 'applyFormatToAllTODO',
        'restoreCellBackgroundColors': 'customCellBGColorTODO',
        'addCheckboxesToSelectedCells': 'addCheckboxesTODO',
        'markCheckboxInSelectedCells': 'markCheckboxTODO',
        'markAllCheckboxesInSelectedCells': 'markAllCheckboxesTODO',
        'restoreCheckboxes': 'restoreCheckboxesTODO',
        'removeAllCheckboxesInSelectedCells': 'removeCheckboxesTODO',
        'applyFormat': 'applyFormatToSelected',
        'applyFormatToAll': 'applyFormatToAll',
        'createPieChart': 'createPieChartTODO',
        'deletePieCharts': 'deleteAllChartsTODO',
        'logHelloWorld': 'logHelloWorld',
        'versionAndFeatureDetails': 'updateCellCommentTODO',
        'saveSnapshot': 'saveSnapshotTODO',
        'restoreSnapshot': 'restoreSnapshotTODO'
    };

    for (const { config, items } of menus) {
        const menuTitle = config[0].title[currentLanguage];
        const menu = ui.createMenu(menuTitle);

        for (const { key, separatorAfter } of items) {
            const itemTitle = config[0].items[key][currentLanguage];
            const functionName = functionNameMap[key] || key;

            menu.addItem(itemTitle, functionName);
            if (separatorAfter) {
                menu.addSeparator();
            }
        }

        menu.addToUi();
    }
}

/**
 * Displays a "Hello World" message in an alert.
 *
 * @customfunction
 */
function logHelloWorld() {
    ui.alert('Hello World!!');
    Logger.log('hello world test');
}


// Contents of ./TODOsheet/TODOpiechart.js

// globals.js: sheet, getDataRange, isPieChartVisible

/**
 * Creates a pie chart in the sheet, displaying the occupied cells in columns C, D, and E.
 * @customfunction
 */
function createPieChartTODO() {
    Logger.log('createPieChartTODO triggered');
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

/**
 * Deletes all charts in the sheet and clears the content in the range J1:K4.
 * @customfunction
 */
function deleteAllChartsTODO() {
    Logger.log('deleteAllChartsTODO triggered');
    const charts = sheet.getCharts();

    charts.forEach(chart => {
        sheet.removeChart(chart);
    });

    sheet.getRange("J1:K4").clearContent();
    Logger.log(`Deleted ${charts.length} charts`);
    isPieChartVisible = false;
}


// Contents of ./TODOsheet/TODOsnapshot.js

// shared/util.js: saveSnapshot, restoreSnapshot

/**
 * Saves a snapshot of the current state of the active sheet while ignoring specific cells.
 * Ignores cells C1, D1, and E1 so we retain the changed column titles when cell max limit is reached.
 * 
 * @return {void}
 */
function saveSnapshotTODO() {
    Logger.log('saveSnapshotTODO triggered');
    const cellsToIgnore = ["R1C1", "R1C2", "R1C3", "R1C4", "R1C5", "R1C6", "R1C7", "R1C8"]
    Logger.log(`Ignoring cells ${cellsToIgnore.join(', ')} from snapshot.`);
    const snapshot = saveSnapshot(cellsToIgnore);

    // Save filtered snapshot to script properties
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('sheetSnapshot', JSON.stringify(snapshot));
    Logger.log("Snapshot saved, excluding specified cells.");
}

/**
 * Restores the sheet snapshot and applies custom formatting for dates and "days left".
 *
 * @return {void}
 */
function restoreSnapshotTODO() {
    Logger.log('restoreSnapshotTODO triggered');
    restoreSnapshot((builder, text) => {
        // Reapply formatting for dates and "days left"
        const dateMatches = text.match(/\d{2}\/\d{2}\/\d{2}/g);
        const daysLeftPattern = /\((\d+)\) days left/;
        const daysLeftMatch = text.match(daysLeftPattern);

        if (dateMatches) {
            Logger.log('restoreSnapshotTODO)(): dateMatches :', dateMatches);
            for (const date of dateMatches) {
                const start = text.lastIndexOf(date);
                const end = start + date.length;
                builder.setTextStyle(start, end, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor('#A9A9A9').build());
                Logger.log('restoreSnapshotTODO() date to be formatted :', date);
            }
        }

        if (daysLeftMatch) {
            Logger.log('restoreSnapshotTODO() daysLeftMatch :', daysLeftMatch);
            const start = text.lastIndexOf(daysLeftMatch[0]);
            const end = start + daysLeftMatch[0].length;
            builder.setTextStyle(start, end, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor('#FF0000').build());
            Logger.log('restoreSnapshotTODO() days left to be formatted :', daysLeftMatch[0]);
        }
    });
}

// for testing

// Contents of ./TODOsheet/TODOtimeHandle.js

// globals.js: sheet, getDataRange, datePattern
// TODOsheet/TODOlibrary.js: dateColorConfig

/**
 * Updates the colors of dates in specific columns based on the time passed.
 *
 * @customfunction
 */
function updateDateColorsTODO() {
    Logger.log('updateDateColorsTODO called');
    const columns = ['C', 'D', 'E', 'F', 'G'];
    const today = new Date();

    for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
        const column = columns[columnIndex];
        const config = dateColorConfig[column];

        processCells((cellRange, cellValue) => {
            const text = cellValue.getText();
            Logger.log(`updateDateColorsTODO(): Checking if cell ${text} contains a date that matches the pattern ${datePattern}`);

            if (datePattern.test(text)) {
                const dateText = text.match(datePattern)[0].trim();
                const cellDate = new Date(dateText.split('/').reverse().join('/'));
                const diffDays = Math.floor((today - cellDate) / (1000 * 60 * 60 * 24));

                const color = diffDays >= config.danger ? config.dangerColor :
                    diffDays >= config.warning ? config.warningColor :
                        config.defaultColor || '#A9A9A9';

                const richTextValue = buildRichTextValue(text, dateText, color, cellValue);
                cellRange.setRichTextValue(richTextValue);
            }
        });

        Logger.log(`updateDateColorsTODO(): Updated date colors for column ${column}`);
    }
}

/**
 * Updates the days left counter for each cell in column H.
 * If the counter reaches zero, the cell is cleared.
 * 
 * @customfunction
 * @return {void}
 */
function updateDaysLeftCounterTODO() {
    Logger.log("updateDaysLeftCounterTODO called");
    const properties = PropertiesService.getDocumentProperties();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    const range = sheet.getRange('H2:H' + sheet.getLastRow());
    const values = range.getValues();
    let updated = false;

    Logger.log("Starting to update days left for each cell.");

    for (let i = 0; i < values.length; i++) {
        const cellValue = values[i][0].toString();
        const match = cellValue.match(/\((\d+)\) days left/);

        if (match) {
            const daysLeft = Math.max(0, parseInt(match[1]) - 1);
            Logger.log(`Row ${i + 2}: original days left = ${match[1]}, new days left = ${daysLeft}`);

            if (daysLeft === 0) {
                sheet.getRange(i + 2, 8).clearContent();
                Logger.log(`Row ${i + 2}: Days left counter reached zero, clearing cell.`);
                values[i][0] = '';
            } else {
                values[i][0] = cellValue.replace(/\(\d+\) days left/, `(${daysLeft}) days left`);
                updated = true;
            }
        }
    }

    if (updated) {
        range.setValues(values);
        properties.setProperty('lastUpdateDate', today);
        Logger.log("Days left counter updated for all applicable cells.");
    } else {
        Logger.log("No updates were necessary.");
    }
}

/**
 * Updates the cell with the number of days left, preserving any existing links.
 * 
 * @param {Range} range - The cell range to update.
 * @param {number} daysLeft - The number of days left to display.
 */
function updateDaysLeftCellTODO(range, daysLeft) {
    Logger.log(`updateDaysLeftCellTODO called`);
    const originalText = range.getValue().toString().split('\n')[0];
    const daysLeftText = `(${daysLeft}) days left`;
    const newText = `${originalText}\n${daysLeftText}`;

    const originalRichTextValue = range.getRichTextValue();
    const richTextValue = buildRichTextValue(newText, daysLeftText, '#FF0000', originalRichTextValue);

    range.setRichTextValue(richTextValue);
    Logger.log(`updateDaysLeftCellTODO(): updated cell with value: ${richTextValue.getText()}`);

    const now = new Date();
    PropertiesService.getDocumentProperties().setProperty(range.getA1Notation(), now.toISOString());
    Logger.log(`updateDaysLeftCellTODO(): set custom property for cell ${range.getA1Notation()}: ${now.toISOString()}`);
}

/**
 * Parses the number of days left from a given value.
 * 
 * @param {string} value - The value to parse for days left.
 * @returns {number} The number of days left, or 60 if not parseable.
 */
function parseDaysLeftTODO(value) {
    Logger.log(`parseDaysLeftTODO called with value: ${value}`);
    const daysLeftMatch = value.match(/\((\d+)\) days left/);
    const daysLeft = daysLeftMatch ? parseInt(daysLeftMatch[1]) : (/^\d+$/.test(value.trim()) ? parseInt(value.trim()) : 60);
    Logger.log(`parseDaysLeftTODO(): parsed days left: ${daysLeft}`);
    return daysLeft;
}

/**
 * Removes multiple dates from cells, keeping only the last occurrence of today's date.
 * 
 * @customfunction
 */
function removeMultipleDatesTODO() {
    Logger.log('removeMultipleDatesTODO called');
    const columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy");

    for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
        const column = columns[columnIndex];

        processCells((cellRange, cellValue) => {
            const text = cellValue.getText();
            Logger.log(`removeMultipleDatesTODO(): Checking cell ${column}${cellRange.getRow()}: ${text}`);

            const dateMatches = text.match(/\d{2}\/\d{2}\/\d{2}/g);
            if (dateMatches && dateMatches.length > 1) {
                Logger.log(`removeMultipleDatesTODO(): Found dates in ${column}${cellRange.getRow()}: ${dateMatches.join(', ')}`);

                let updatedText = text;
                for (let i = 0; i < dateMatches.length; i++) {
                    const date = dateMatches[i];
                    const lastOccurrence = updatedText.lastIndexOf(date);
                    if (lastOccurrence !== -1) {
                        updatedText = updatedText.substring(0, lastOccurrence) + updatedText.substring(lastOccurrence).replace(new RegExp(date, 'g'), '');
                    }
                }

                updatedText = updatedText.replace(new RegExp(`\\b(${dateMatches.join('|')})\\b`, 'g'), '').trim() + `\n${today}`;
                Logger.log(`removeMultipleDatesTODO(): Updated text for ${column}${cellRange.getRow()}: ${updatedText}`);

                const richTextValue = buildRichTextValue(updatedText, today, '#A9A9A9', cellValue);
                cellRange.setRichTextValue(richTextValue);
                Logger.log(`Cell ${column}${cellRange.getRow()} updated with value: ${richTextValue.getText()}`);
            }
        });
    }

    Logger.log('End removeMultipleDatesTODO');
}

// for testing

// Contents of ./TODOsheet/TODOtoggleFn.js

// globals.js: sheet, isPieChartVisible
// TODOsheet/TODOpiechart.js: createPieChartTODO, deleteAllChartsTODO

/**
 * Toggles the visibility of the piechart
 * @param {string} action - The action to be performed
 * @returns {void}
 */
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

/**
 * Handles the Piechart toggle action
 * @param {Range} range - The range containing the action
 * @returns {void}
 */
function handlePieChartToggleTODO(range) {
    Logger.log('handlePieChartToggleTODO called');
    const action = range.getValue().toString().trim();
    Logger.log(`Action selected: ${action}`);
    if (action === 'Show Piechart' || action === 'Hide Piechart') {
        togglePieChartTODO(action);
    } else {
        Logger.log('Invalid action selected');
    }
    sheet.getRange("I1").setValue("Piechart");
}


// Contents of ./TODOsheet/TODOtranslate.js


// globals.js: sheet
// TODOsheet/TODOlibrary.js: languages

const setLanguageEnglish = () => setLanguage('English');
const setLanguageSpanish = () => setLanguage('Spanish');
const setLanguageCatalan = () => setLanguage('Catalan');

function setLanguage(language) {
    Logger.log('setLanguage called with language: ' + language);
    if (languages[language]) {
        PropertiesService.getDocumentProperties().setProperty('language', language);
        translateSheetTODO();
        const ui = SpreadsheetApp.getUi();
        const message = {
            'English': 'Language changed.\n Please reload the sheet to update menus.',
            'Spanish': 'Idioma cambiado.\n Por favor, recargue la hoja para actualizar los menús.',
            'Catalan': 'Idioma canviat.\n Si us plau, recarregui el full per actualitzar els menús.'
        };
        ui.alert(message[language]);
    } else {
        Logger.log('Language not supported: ' + language);
    }
}

/**
 * Translates the sheet to the selected language
 * @returns {void}
 * @customfunction
 */
function translateSheetTODO() {
    Logger.log('translateSheetTODO called');
    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';

    // Update with the corresponding styles
    for (const cell in cellStyles) {
        const cellData = cellStyles[cell];
        if (cellData.value[language]) {
            let range = sheet.getRange(cell);
            range.setValue(cellData.value[language])
                .setFontWeight(cellData.fontWeight)
                .setFontColor(cellData.fontColor)
                .setHorizontalAlignment(cellData.alignment);

            if (cellData.backgroundColor) {
                range.setBackground(cellData.backgroundColor);
            }
        }
    }

    // Update the example texts with the corresponding language
    const range = sheet.getDataRange();
    const values = range.getValues();
    for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
            for (const exampleKey in exampleTexts) {
                if (typeof values[i][j] === 'string' && values[i][j].startsWith("Example:")) {
                    const exampleData = exampleTexts[exampleKey];
                    if (exampleData.text[language]) {
                        sheet.getRange(i + 1, j + 1).setValue(exampleData.text[language]);
                    }
                }
            }
        }
    }
}

// for testing

// Contents of ./TODOsheet/TODOtriggers.js

// globals.js: sheet
// TODOsheet/TODOtoggleFn.js: handlePieChartToggleTODO
// TODOsheet/TODOformatting.js: shiftCellsUpTODO, handleColumnEditTODO, addCheckboxToCellTODO

/**
 * Track changes in specified columns and add the date.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object for the edit trigger.
 * @customfunction
 */
function onEdit(e) {
    Logger.log('onEdit triggered');
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

        if (column === 9 && row === 1) {
            handlePieChartToggleTODO(range);
            return;
        }

        const originalValue = e.oldValue || '';
        const newValue = range.getValue().toString();

        Logger.log(`onEdit(): Original value: "${originalValue}", New value: "${newValue}"`);

        if ((column === 1 || (column >= 3 && column <= 8)) && row >= 2 && newValue.trim() === '') {
            Logger.log(`onEdit(): Shifting cells up for column ${column}`);
            shiftCellsUpTODO(column, 2, totalRows);
            return;
        }

        if (row >= 2 && column >= 3 && column <= 8) {
            Logger.log(`onEdit()/handleColumnEditTODO(): Handling column edit for column ${column}`);
            handleColumnEditTODO(range, originalValue, newValue, columnLetter, row, e);
            if (newValue && !newValue.includes('☑️')) {
                Logger.log(`onEdit(): Adding default checkbox to cell ${columnLetter}${row}`);
                addCheckboxToCellTODO(range);
            }
        }
    } catch (error) {
        Logger.log(`Error in onEdit: ${error.message}`);
        Logger.log(`Error stack: ${error.stack}`);
    }
}

