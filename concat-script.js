// Auto-generated file with all JS scripts

// Contents of ./globals.js



const ui = SpreadsheetApp.getUi();
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const getDataRange = () => sheet.getDataRange();
const datePattern = /\n\d{2}\/\d{2}\/\d{2}$/; // dd/MM/yy

// state management
let isPieChartVisible = false;
let isLoaded = false;


// Contents of ./shared/documentProperties.js

/**
 * Logs all the document properties to the console.
 * @returns {void}
 * @customfunction
 */
function logAllDocumentProperties() {
    const docProperties = PropertiesService.getDocumentProperties();
    const allProperties = docProperties.getProperties();

    const expirationDateHash = {};
    const lastHashProperty = {};
    const timestampProperties = {};
    const configProperties = {};
    const otherProperties = {};

    const conditionsMap = [
        {
            condition: key => key === 'expirationDateHash',
            action: key => expirationDateHash[key] = allProperties[key]
        },
        {
            condition: key => key === 'lastHash',
            action: key => lastHashProperty[key] = allProperties[key]
        },
        {
            condition: key => key.toLowerCase().includes('h') && !isNaN(parseInt(key.substring(1))),
            action: key => timestampProperties[key] = allProperties[key]
        },
        {
            condition: key => key.toLowerCase().includes('last') || key.toLowerCase().includes('enable') || key.toLowerCase().includes('menus'),
            action: key => configProperties[key] = allProperties[key]
        }
    ];

    for (const key in allProperties) {
        let matched = false;
        for (const { condition, action } of conditionsMap) {
            if (condition(key)) {
                action(key);
                matched = true;
                break;
            }
        }
        if (!matched) {
            otherProperties[key] = allProperties[key];
        }
    }

    Logger.log('Document Properties:');

    logProperties('EXPIRATION DATE HASH', expirationDateHash);
    logProperties('LAST HASH', lastHashProperty);
    logProperties('TIMESTAMP PROPERTIES for column H', timestampProperties);
    logProperties('CONFIG PROPERTIES', configProperties);
    logProperties('OTHER PROPERTIES', otherProperties);
}

/**
 * Logs the properties of a category to the console.
 * @param {string} categoryName The name of the category.
 * @param {Object} properties The properties to log.
 * @returns {void}
 * @customfunction
 */
function logProperties(categoryName, properties) {
    Logger.log(`${categoryName}:`);
    for (const key in properties) {
        Logger.log(`${key}: ${properties[key]}`);
    }
}

/**
 * Removes unused properties from the document properties.
 * @returns {void}
 * @customfunction
 */
function removeUnusedProperties() {
    const docProperties = PropertiesService.getDocumentProperties();
    const allProperties = docProperties.getProperties();

    const unusedKeys = [
        // add here the keys that are not used as strings:
    ];
    let removedKeys = [];

    for (const key of unusedKeys) {
        if (key in allProperties) {
            docProperties.deleteProperty(key);
            removedKeys.push(key);
        }
    }

    if (removedKeys.length > 0) {
        Logger.log('Removed the following unused properties:');
        for (const key of removedKeys) {
            Logger.log(key);
        }
    } else {
        Logger.log('No unused properties found to remove.');
    }
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

/**
 * Removes notes from empty cells in a sheet.
 *
 * @param {Sheet} sheet - The sheet to remove notes from.
 */
function removeNotesFromEmptyCells(sheet) {
    const lastRow = Math.min(40, sheet.getLastRow()); // Limit to 40 rows
    const lastColumn = sheet.getLastColumn();
    const range = sheet.getRange(2, 1, lastRow - 1, lastColumn); //From row 2 to lastRow
    const values = range.getValues();
    const notes = range.getNotes();

    for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
            if (values[row][col] === "" && notes[row][col] !== "") {
                sheet.getRange(row + 2, col + 1).setNote(""); // Clear note
            }
        }
    }
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
    Logger.log(`returning urls: ${urls}`);
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
 * Stores a hash of the current date to check for expiration.
 *
 * @customfunction
 * @returns {void}
 */
function storeExpirationDateHashTODO() {
    Logger.log('storeExpirationDateHashTODO triggered');
    // Create and store a hash of the current date to check for expiration
    const today = new Date();
    const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    const dateHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, dateString);
    const hashString = Utilities.base64Encode(dateHash);

    const docProperties = PropertiesService.getDocumentProperties();
    docProperties.setProperty('expirationDateHash', hashString);

    Logger.log('Stored expiration date hash: ' + hashString);
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
                    Logger.log(`Restoring link: ${link.url} at ${link.start}-${link.end}`);
                    builder.setLinkUrl(link.start, link.end, link.url);
                }
                Logger.log(`Restored links for cell ${cellKey}. With a total of ${cellData.links.length}`);
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

/**
 * Converts a date string in the format dd/MM/yy to a JavaScript Date object.
 * 
 * @param {string} dateString - La fecha en formato dd/MM/yy.
 * @return {Date} - Un objeto Date de JavaScript.
 */
function parseDate(dateString) {
    const [day, month, year] = dateString.split('/').map(Number);
    const fullYear = year + 2000;  // Assume years are in the 21st century
    return new Date(fullYear, month - 1, day);  // Months are 0-based
}

/**
 * Converts a date string in the format dd/MM/yyyy to a JavaScript Date object.
 * 
 * @param {string} dateString - La fecha en formato dd/MM/yyyy.
 * @return {Date} - Un objeto Date de JavaScript.
 */
function parseFullYearDate(dateString) {
    Logger.log(`Parsing date string with full year: ${dateString}`);
    const [day, month, year] = dateString.split('/').map(Number);
    const date = new Date(year, month - 1, day);  // Months are 0-based in JavaScript's Date object
    Logger.log(`Parsed date: ${date}`);
    return date;
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
 * Updates the colors of the dates based on the priority of the column.
 *  
 * @customfunction
 */
function updateTipsCellTODO() {
    Logger.log('updateTipsCellTODO called');
    const cell = sheet.getRange("I2");

    const tips = {
        English: "💡 To add an expiration date countdown, simply add the date in the format **dd/mm/yyyy** to the desired cell.(Don't forget to add the asterisks)",
        Spanish: "💡 Para añadir una fecha de expiración a modo de cuenta atrás en días, basta con añadir la fecha en formato **dd/mm/yyyy** a la celda en cuestión. (No olvides añadir los asteriscos)",
        Catalan: "💡 Per afegir una data de caducitat en mode compte enrere en dies, només cal afegir la data en format **dd/mm/yyyy** a la cel·la en qüestió. (No oblidis afegir els asteriscs)"
    };

    const titles = {
        English: "💡Tips",
        Spanish: "💡Consejos",
        Catalan: "💡Consells"
    };

    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';
    const tipText = tips[language];
    const titleText = titles[language];

    cell.setNote(tipText);
    cell.setValue(titleText);
    cell.setFontWeight("bold");
    cell.setFontSize(12);
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
    cell.setBackground("#efefef");
    cell.setBorder(true, true, true, true, true, true, '#D3D3D3', SpreadsheetApp.BorderStyle.SOLID_THICK);

    Logger.log('Tips cell updated with tips for language: ' + language);
}

/**
 * Updates the cell comment with the latest changes.
 * 
 * @customfunction
 */
function updateCellCommentTODO() {
    Logger.log('updateCellCommentTODO called');
    const cell = sheet.getRange("I3");
    const version = "v0.3.0";
    cell.setValue(version);

    const changes = {
        English: `
            NEW FEATURES:
            - You can now add an expiration date to your cells. It is a countdown by days. See the "Tips" cell for more information.
            - Cells with expiration dates come with notes added that are used to calculate and update the expiration days.
            - You can now enable or disable all the functionalities that are added by default when writing in a cell. For example, for a cell to have only text, without checkboxes or default date. From the "TODO Sheet"/"Enable/Disable Default Additions in Cells" menu.
            OLD FEATURES:
            - A checkbox is added by default from the 3rd to the 8th column when a cell is written or modified.
            - You can add, mark, restore, and delete checkboxes in cells by selecting them and using the "Custom Formats" menu.
            - The "days left" counter is updated daily in the 8th column. When the counter reaches zero, the cell is cleared.
            - A snapshot of the sheet can be saved and restored from the "Custom Formats" menu.
            - Indicative limit of cells for each priority, with a warning when the limit is reached.
            - Custom formats can be applied without refreshing the page from the "Custom Formats" menu.
            - Date color change times vary by column priority.
            - The Piechart can be shown or hidden using its dropdown cell.
            - Deleted empty cells are replaced by the immediately lower cell.
        `,
        Spanish: `
            NUEVAS FUNCIONES:
            - Ahora puedes agregar una fecha de vencimiento a tus celdas. Es una cuenta regresiva en días. Consulta la celda "Consejos" para obtener más información.
            - Las celdas con fecha de expiracion vienen con notas agregadas que son usadas para calcular y actualizar los dias de expiración.
            - Ahora puedes activar o desactivar todas las funcionalidades que se añaden por defecto al escribir en una celda. Por ejemplo para que una celda tenga solamente texto, sin casillas ni fecha por defecto. Desde el menu "Hoja TODO"/"Habilitar/Deshabilitar Adiciones por Defecto en las Celdas"
            FUNCIONES ANTIGUAS:
            - Se añade una casilla de verificación por defecto desde la 3ª a la 8ª columna cuando se escribe o modifica una celda.
            - Puedes agregar, marcar, restaurar y eliminar casillas en las celdas seleccionándolas y usando el menú "Formatos personalizados".
            - El contador de "días restantes" se actualiza diariamente en la 8ª columna. Cuando el contador llega a cero, la celda se borra.
            - Se puede guardar y restaurar una instantánea de la hoja desde el menú "Formatos personalizados".
            - Límite indicativo de celdas para cada prioridad, con una advertencia cuando se alcanza el límite.
            - Se pueden aplicar formatos personalizados sin necesidad de refrescar la página desde el menú "Formatos personalizados".
            - Los tiempos de cambio de color de las fechas varían según la prioridad de la columna.
            - El gráfico circular se puede mostrar u ocultar usando su celda desplegable.
            - Las celdas vacías eliminadas son reemplazadas por la celda inmediatamente inferior.
        `,
        Catalan: `
            NOVES FUNCIONS:
            - Ara pots agregar una data de venciment a les cel·les. És un compte enrere en dies. Consulta la cel·la "Consells" per obtenir més informació.
            - Les cel·les amb data d'expiració venen amb notes agregades que són usades per calcular i actualitzar els dies d'expiració.
            - Ara podeu activar o desactivar totes les funcionalitats que s'afegeixen per defecte en escriure en una cel·la. Per exemple perquè una cel·la tingui només text, sense caselles ni data per defecte. Des del menu "Full de TODO"/"Habilitar/Deshabilitar Addicions per Defecte a les Cel·les"
            FUNCIONS ANTIGUES:
            - S'afegeix una casella de verificació per defecte des de la 3a fins a la 8a columna quan s'escriu o es modifica una cel·la.
            - Pots afegir, marcar, restaurar i eliminar caselles en les cel·les seleccionades seleccionant-les i utilitzant el menú "Formats personalitzats".
            - El comptador de "dies restants" s'actualitza diàriament a la 8a columna. Quan el comptador arriba a zero, la cel·la s'esborra.
            - Es pot desar i restaurar una instantània del full des del menú "Formats personalitzats".
            - Límite indicatiu de cel·les per a cada prioritat, amb una advertència quan s'assoleix el límit.
            - Es poden aplicar formats personalitzats sense necessitat de refrescar la pàgina des del menú "Formats personalitzats".
            - Els temps de canvi de color de les dates varien segons la prioritat de la columna.
            - El gràfic circular es pot mostrar o ocultar utilitzant la seva cel·la desplegable.
            - Les cel·les buides eliminades són reemplaçades per la cel·la immediatament inferior.
        `
    };

    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';
    const comment = `Version: ${version}\n${changes[language]}`;

    cell.setComment(comment);
    cell.setBackground("#efefef");
    cell.setBorder(true, true, true, true, true, true, '#D3D3D3', SpreadsheetApp.BorderStyle.SOLID_THICK);
    Format(cell);

    Logger.log('Cell comment updated with changes for language: ' + language);
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
        Logger.log(`Example text set for column ${column} at ${column} 2: ${exampleText}`);
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
    const totalRows = Math.min(40, sheet.getMaxRows()); // Limited to 40 rows
    const range = sheet.getRange(1, 1, totalRows, 8);

    Logger.log('applyFormatToAllTODO()/preserveRelevantHyperlinks() called');
    const preservedLinks = preserveRelevantHyperlinks(range);

    Logger.log('applyFormatToAllTODO()/updateDateColorsTODO() called');
    updateDateColorsTODO();  // Update date colors selectively

    Logger.log('applyFormatToAllTODO()/setCellContentAndStyle(): setting cell content and style');
    setCellContentAndStyleTODO();

    Logger.log('applyFormatToAllTODO()/exampleTextTODO(): setting example text');
    applyExampleTexts(language);

    if (range) {
        Format(range);
        applyBorders(range);
    }

    Logger.log('applyFormatToAllTODO()/restoreRelevantHyperlinks() called');
    restoreRelevantHyperlinks(range, preservedLinks);

    Logger.log('applyFormatToAllTODO()/applyExpiresTextStyle() called');
    applyExpiresTextStyle();

    Logger.log('applyFormatToAllTODO()/applyThickBorders(): applying thick borders');
    applyThickBorders(sheet.getRange(1, 3, 11, 1));
    applyThickBorders(sheet.getRange(1, 4, 21, 1));
    applyThickBorders(sheet.getRange(1, 5, 21, 1));

    Logger.log('applyFormatToAllTODO()/checkAndSetColumnTODO(): checking and setting columns');
    applyColumnStyles(language);
}

/**
 * Preserves relevant hyperlinks in the specified range.
 * 
 * @param {Range} range - The range to preserve hyperlinks in.
 * @return {Array} The preserved hyperlinks.
 * @customfunction
 */
function preserveRelevantHyperlinks(range) {
    Logger.log('preserveRelevantHyperlinks called');
    const richTextValues = range.getRichTextValues();
    const preservedLinks = [];
    const maxRows = Math.min(richTextValues.length, 40); // Limited to 40 rows

    for (let row = 0; row < maxRows; row++) {
        let rowHasRelevantData = false;
        preservedLinks[row] = [];
        for (let col = 1; col <= 6; col++) { // columns B to H 
            const richText = richTextValues[row][col];
            const cellText = richText.getText().trim();

            if (cellText === '') {
                preservedLinks[row][col] = null;  // Omit empty cells
                continue;
            }

            if (richText.getLinkUrl() || cellText.includes('Expires in')) {
                preservedLinks[row][col] = richText;
                rowHasRelevantData = true;
                Logger.log(`Row ${row + 1}, Column ${col + 1}: Relevant data preserved.`);
            } else {
                preservedLinks[row][col] = null;
            }
        }

        if (!rowHasRelevantData) {
            preservedLinks[row] = null;  // delete rows with no relevant data
        }
    }

    Logger.log(`preserveRelevantHyperlinks completed: Total rows preserved ${preservedLinks.length}`);
    return preservedLinks;
}

/**
 * Restores the relevant hyperlinks in the specified range.
 * 
 * @param {Range} range - The range to restore hyperlinks in.
 * @param {Array} preservedLinks - The preserved hyperlinks.
 * @customfunction
 * @returns {void}
 */
function restoreRelevantHyperlinks(range, preservedLinks) {
    Logger.log('restoreRelevantHyperlinks called');
    const richTextValues = range.getRichTextValues();

    const maxRows = Math.min(preservedLinks.length, 40); // Limited to 40 rows

    for (let row = 0; row < maxRows; row++) {
        if (preservedLinks[row] !== null) {
            for (let col = 1; col <= 6; col++) { // columns B to H 
                if (preservedLinks[row][col] !== null) {
                    richTextValues[row][col] = preservedLinks[row][col];
                    Logger.log(`Row ${row + 1}, Column ${col + 1}: Restoring preserved data.`);
                }
            }
        }
    }

    range.setRichTextValues(richTextValues); // apply all restored rich text values at once
    Logger.log('restoreRelevantHyperlinks completed');
}

/**
 * Sets the example text for each column in the sheet.
 * 
 * @param {language} language - The language to set the content in.
 * @customfunction
 */
function applyExampleTexts(language) {
    for (const column in exampleTexts) {
        const { text } = exampleTexts[column];
        const translatedText = text[language];
        exampleTextTODO(column, translatedText);
        Logger.log(`applyFormatToAllTODO(): example text set for column ${column} - translatedText: ${translatedText}`);
    }
}

/**
 * Sets the column content and style based on the language.
 * 
 * @param {language} language - The language to set the content in.
 * @customfunction
 */
function applyColumnStyles(language) {
    for (const column in cellStyles) {
        const { limit, priority } = cellStyles[column];

        const translatedLimit = limit?.[language];
        const translatedPriority = priority?.[language];

        if (translatedLimit !== undefined && translatedPriority !== undefined) {
            checkAndSetColumnTODO(column.charAt(0), translatedLimit, translatedPriority);
            Logger.log(`applyFormatToAllTODO(): translatedText set for column ${column} - limit: ${translatedLimit}, priority: ${translatedPriority}`);
        } else {
            Logger.log(`applyFormatToAllTODO(): limit or priority not found for column ${column} and language ${language}`);
        }
    }
}

/**
 * Sets the expiration date text style in the sheet.
 * 
 * @customfunction
 * @returns {void}
 */
function applyExpiresTextStyle() {
    Logger.log('applyExpiresTextStyle called');
    const totalRows = Math.min(40, sheet.getMaxRows()); // Limited to 40 rows
    const range = sheet.getRange(1, 1, totalRows, 8);
    const richTextValues = range.getRichTextValues();

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const richText = richTextValues[row][col];
            const text = richText.getText();
            const expiresInIndex = text.indexOf('Expires in');
            const dateMatch = text.match(/\d{2}\/\d{2}\/\d{2}/);  // Match for the date format DD/MM/YY

            if (expiresInIndex !== -1 && dateMatch) {
                const dateIndex = text.indexOf(dateMatch[0]);  // Get the starting index of the date

                // Ensure indices are valid and in order
                if (expiresInIndex < dateIndex && expiresInIndex >= 0 && dateIndex >= 0) {
                    const builder = richText.copy();

                    // Preserve existing text styles
                    const existingStyle = richText.getTextStyle(0, text.length);

                    // Apply the new styles only to the "Expires in" and date sections
                    builder.setTextStyle(0, text.length, existingStyle); // Preserve the existing style
                    builder.setTextStyle(expiresInIndex, dateIndex, SpreadsheetApp.newTextStyle().setForegroundColor('#0000FF').setItalic(true).build());

                    range.getCell(row + 1, col + 1).setRichTextValue(builder.build());
                } else {
                    Logger.log(`Skipping invalid indices for styling in cell ${row + 1}, ${col + 1}`);
                }
            }
        }
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

    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';

    // Define messages based on the selected language
    const messages = {
        English: {
            cellLimitReached: "⚠️CELL LIMIT REACHED⚠️",
            alertMessage: "⚠️CELL LIMIT REACHED⚠️ \nfor " + priority
        },
        Spanish: {
            cellLimitReached: "⚠️LÍMITE DE CELDAS ALCANZADO⚠️",
            alertMessage: "⚠️LÍMITE DE CELDAS ALCANZADO⚠️ \npara la " + priority
        },
        Catalan: {
            cellLimitReached: "⚠️LÍMIT DE CEL·LES ASSOLIT⚠️",
            alertMessage: "⚠️LÍMIT DE CEL·LES ASSOLIT⚠️ \nper a la " + priority
        }
    };

    const message = messages[language];

    const dataRange = getDataRange();
    const values = sheet.getRange(column + "2:" + column + dataRange.getLastRow()).getValues().flat();
    const occupied = values.filter(String).length;
    const range = sheet.getRange(column + "2:" + column + dataRange.getLastRow());

    if (occupied > limit) {
        // Set red border with thicker style
        range.setBorder(true, true, true, true, true, true, "#FF0000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(column + "1").setValue(message.cellLimitReached);
        SpreadsheetApp.getUi().alert(message.alertMessage);
    } else {
        // Set black border with thicker style
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
    let totalRows = Math.min(40, sheet.getMaxRows()); // Limited to 40 rows
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
    setColumnBackground(sheet, 6, '#eef7ff', 2); // Column F: Light blue 3
    setColumnBackground(sheet, 7, '#fff1f1', 2); // Column G: Light red 3

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

            values[i][0] = values[i + 1][0];
            richTextValues[i][0] = richTextValues[i + 1][0];

            moveNoteToUpperCell(sheet, startRow + i + 1, startRow + i, column);

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
 * Moves the note from one cell to another.
 * 
 * @param {Sheet} sheet - The sheet object.
 * @param {number} fromRow - The row number to move the note from.
 * @param {number} toRow - The row number to move the note to.
 * @param {number} column - The column number.
 * @returns {void}
 * @customfunction
 */
function moveNoteToUpperCell(sheet, fromRow, toRow, column) {
    const note = sheet.getRange(fromRow, column).getNote();
    sheet.getRange(toRow, column).setNote(note);
    sheet.getRange(fromRow, column).setNote('');
}

/**
 * Forces empty cells to shift up in specified columns.
 *
 * @customfunction
 */
function pushUpEmptyCellsTODO() {
    Logger.log('pushUpEmptyCellsTODO called');
    const range = sheet.getDataRange();
    const numRows = Math.min(40, range.getNumRows()); // Limited to 40 rows
    const numCols = range.getNumColumns();

    for (let col = 1; col <= numCols; col++) {
        let startRow = null;
        for (let row = 2; row <= numRows; row++) {
            if (sheet.getRange(row, col).getValue() === '' && startRow === null) {
                startRow = row;
            } else if (sheet.getRange(row, col).getValue() !== '' && startRow !== null) {
                shiftCellsUpTODO(col, startRow, numRows);
                startRow = null;
            }
        }
        // If the last rows in the column are empty
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

    let originalRichTextValue;
    try {
        originalRichTextValue = range.getRichTextValue() || SpreadsheetApp.newRichTextValue().setText(originalValue).build();
    } catch (error) {
        Logger.log(`Error getting original rich text value: ${error.message}`);
        return;
    }

    // Add the date if it's not already present
    const datePattern = /\d{2}\/\d{2}\/\d{2}$/;
    if (!datePattern.test(updatedText)) {
        updatedText = `${updatedText}\n${dateFormatted}`;
        Logger.log(`No date found, updated text with new date: "${updatedText}"`);
    }

    // Add a checkbox if it's not already present
    if (!updatedText.startsWith('☑️')) {
        updatedText = `☑️${updatedText}`;
        Logger.log(`Checkbox added to the start of the text: "${updatedText}"`);
    }

    try {
        // Apply the updated text to the cell
        const newRichTextValueBuilder = SpreadsheetApp.newRichTextValue()
            .setText(updatedText);

        // preserve links from the original text
        const originalText = originalRichTextValue.getText();
        Logger.log(`Preserving links from original text: ${originalText}`);
        for (let i = 0; i < originalText.length && i < updatedText.length; i++) {
            const url = originalRichTextValue.getLinkUrl(i, i + 1);
            if (url) {
                newRichTextValueBuilder.setLinkUrl(i, i + 1, url);
            }
        }

        // Apply italic style to the date
        const dateStartIdx = updatedText.search(datePattern);
        if (dateStartIdx !== -1) {
            const dateEndIdx = updatedText.length;
            newRichTextValueBuilder.setTextStyle(dateStartIdx, dateEndIdx, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor('#A9A9A9').build());
        }

        range.setRichTextValue(newRichTextValueBuilder.build());
        Logger.log(`Set new rich text value for cell ${columnLetter}${row}`);
    } catch (error) {
        Logger.log(`Error in updateRichTextTODO: ${error.message}`);
    }
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
            "English": "TODAY",
            "Spanish": "HOY",
            "Catalan": "AVUI"
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
        backgroundColor: "#eef7ff",
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
        backgroundColor: "#fff1f1",
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
            "Spanish": "PRIORIDAD BAJA",
            "Catalan": "PRIORITAT BAIXA"
        },
        limit: {
            "English": 20,
            "Spanish": 20,
            "Catalan": 20
        },
        priority: {
            "English": "LOW PRIORITY",
            "Spanish": "PRIORIDAD BAJA",
            "Catalan": "PRIORITAT BAIXA"
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
    D: { warning: 15, danger: 180, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
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
            enableDefaultAdditions: {
                English: 'Enable Default Additions in Cells',
                Spanish: 'Habilitar Adiciones por Defecto en las Celdas',
                Catalan: 'Habilitar Addicions per Defecte a les Cel·les'
            },
            disableDefaultAdditions: {
                English: 'Disable Default Additions in Cells',
                Spanish: 'Deshabilitar Adiciones por Defecto en las Celdas',
                Catalan: 'Deshabilitar les Addicions per Defecte a les Cel·les'
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
            { key: 'enableDefaultAdditions', separatorAfter: false },
            { key: 'disableDefaultAdditions', separatorAfter: true },
            { key: 'saveSnapshot', separatorAfter: false },
            { key: 'restoreSnapshot', separatorAfter: true },
            { key: 'createPieChart', separatorAfter: false },
            { key: 'deletePieCharts', separatorAfter: true },
            { key: 'versionAndFeatureDetails', separatorAfter: false },
            { key: 'logHelloWorld', separatorAfter: false },
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
        'enableDefaultAdditions': 'enableDefaultAdditionsTODO',
        'disableDefaultAdditions': 'disableDefaultAdditionsTODO',
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
// shared/utils.js: parseDate

/**
 * Updates the colors of dates in specific columns based on the time passed.
 *
 * @customfunction
 */
function updateDateColorsTODO() {
    Logger.log('updateDateColorsTODO called');
    const columns = ['C', 'D', 'E', 'F', 'G'];
    const dataRange = getDataRange();
    const lastRow = dataRange.getLastRow();

    const datePattern = /\d{2}\/\d{2}\/\d{2}$/;  // dd/MM/yy
    const expiresPattern = /Expires in \(\d+\) days/;  // Expires in (n) days

    // index the cells to update
    const cellsToUpdate = [];

    for (const column of columns) {
        for (let row = 2; row <= lastRow; row++) {
            const cell = sheet.getRange(`${column}${row}`);
            let cellValue = cell.getValue();
            if (datePattern.test(cellValue) || expiresPattern.test(cellValue)) {
                cellsToUpdate.push({ cell, cellValue, column, row });
                Logger.log(`Identified cell ${column}${row} for update: ${cellValue}`);
            }
        }
    }

    Logger.log(`Total cells to update: ${cellsToUpdate.length}`);

    // process indexed cells
    cellsToUpdate.forEach(({ cell, cellValue, column, row }) => {
        const config = dateColorConfig[column];

        let dateText = cellValue.match(datePattern);
        if (dateText) {
            dateText = dateText[0];
            const cellDate = parseDate(dateText);
            const today = new Date();
            const diffDays = Math.floor((today - cellDate) / (1000 * 60 * 60 * 24));

            Logger.log(`Processing cell ${column}${row}: ${cellValue}`);
            Logger.log(`Days difference: ${diffDays}`);

            let dateColor = config.defaultColor;
            if (diffDays >= config.danger) {
                dateColor = config.dangerColor;
                Logger.log(`Applying danger color: ${config.dangerColor}`);
            } else if (diffDays >= config.warning) {
                dateColor = config.warningColor;
                Logger.log(`Applying warning color: ${config.warningColor}`);
            } else {
                Logger.log(`Applying default color: ${config.defaultColor}`);
            }

            const richTextValueBuilder = SpreadsheetApp.newRichTextValue().setText(cellValue);
            const dateIndex = cellValue.indexOf(dateText);
            const expiresIndex = cellValue.indexOf(`Expires in`);

            richTextValueBuilder.setTextStyle(dateIndex, dateIndex + dateText.length, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor(dateColor).build());

            if (expiresIndex !== -1) {
                const endIndex = expiresIndex + `Expires in (${diffDays}) days`.length;
                Logger.log(`Setting Expires in style from ${expiresIndex} to ${endIndex}`);
                richTextValueBuilder.setTextStyle(expiresIndex, endIndex, SpreadsheetApp.newTextStyle().setForegroundColor('#0000FF').setItalic(true).build());
            }

            Logger.log(`Updating cell ${column}${row} with new styles`);
            cell.setRichTextValue(richTextValueBuilder.build());
        }
    });
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
    const lastUpdateDate = properties.getProperty('lastUpdateDate');
    const today = new Date();
    const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

    Logger.log(`Last update was on: ${lastUpdateDate}`);
    Logger.log(`Today's date is: ${todayFormatted}`);

    // Calculate the number of days elapsed since the last update
    let daysElapsed = 0;
    if (lastUpdateDate) {
        const lastUpdate = new Date(lastUpdateDate);
        daysElapsed = Math.floor((today - lastUpdate) / (1000 * 60 * 60 * 24));
        Logger.log(`Days elapsed since last update: ${daysElapsed}`);
    }

    const lastRow = Math.min(40, sheet.getLastRow()); // Limit to 40 rows for performance
    const range = sheet.getRange('H2:H' + lastRow);
    const richTextValues = range.getRichTextValues();
    Logger.log("Starting to update days left for each cell.");

    for (let i = 0; i < richTextValues.length; i++) {
        let cellRichTextValue = richTextValues[i][0];
        const cellValue = cellRichTextValue.getText().toString();
        const match = cellValue.match(/\((\d+)\) days left/);

        if (match) {
            const originalDays = parseInt(match[1]);
            const daysLeft = Math.max(0, originalDays - daysElapsed); // No negative days left

            Logger.log(`Row ${i + 2}: original days left = ${originalDays}, new days left = ${daysLeft}`);

            if (daysLeft <= 0) {
                richTextValues[i][0] = SpreadsheetApp.newRichTextValue().setText('').build();
                Logger.log(`Row ${i + 2}: Days left counter reached zero, clearing cell.`);
            } else {
                let newText = cellValue.replace(`(${originalDays}) days left`, `(${daysLeft}) days left`);
                let newRichTextValueBuilder = SpreadsheetApp.newRichTextValue().setText(newText);

                // copy text styles from the original rich text value
                for (let j = 0; j < cellRichTextValue.getRuns().length; j++) {
                    const startOffset = cellRichTextValue.getRuns()[j].getStartIndex();
                    const endOffset = cellRichTextValue.getRuns()[j].getEndIndex();
                    const textStyle = cellRichTextValue.getTextStyle(startOffset, endOffset);

                    newRichTextValueBuilder.setTextStyle(startOffset, endOffset, textStyle);
                }

                richTextValues[i][0] = newRichTextValueBuilder.build();
            }
        }
    }

    range.setRichTextValues(richTextValues);

    if (daysElapsed > 0) {
        properties.setProperty('lastUpdateDate', todayFormatted);
    }

    Logger.log("Days left counter updated for all applicable cells.");
}

/**
 * Updates the cell with the number of days left, preserving any existing links.
 * 
 * @param {Range} range - The cell range to update.
 * @param {number} daysLeft - The number of days left to display.
 */
function updateDaysLeftCellTODO(range, daysLeft) {
    Logger.log(`updateDaysLeftCellTODO called`);
    let originalText = range.getValue().toString().split('\n')[0];
    let daysLeftText = `(${daysLeft}) days left`;
    let newText = originalText + '\n' + daysLeftText;

    // Get the original rich text value to preserve links
    const originalRichTextValue = range.getRichTextValue();
    Logger.log(`updateDaysLeftCellTODO(): getting original rich text value: ${originalRichTextValue.getText()}`);

    // Create new rich text value with updated text and styling
    let newRichTextValueBuilder = SpreadsheetApp.newRichTextValue()
        .setText(newText)
        .setTextStyle(0, originalText.length, SpreadsheetApp.newTextStyle().build())
        .setTextStyle(originalText.length + 1, newText.length,
            SpreadsheetApp.newTextStyle().setForegroundColor('#FF0000').setItalic(true).build());

    // Preserve links from the original rich text value
    const originalTextLength = originalRichTextValue.getText().length;
    Logger.log(`updateDaysLeftCellTODO(): original text length: ${originalTextLength}`);
    for (let i = 0; i < Math.min(newText.length, originalTextLength); i++) {
        const url = originalRichTextValue.getLinkUrl(i, i + 1);
        if (url) {
            newRichTextValueBuilder.setLinkUrl(i, i + 1, url);
            Logger.log(`updateDaysLeftCellTODO(): set link for character ${i}: ${url}`);
        }
    }

    let newRichTextValue = newRichTextValueBuilder.build();
    range.setRichTextValue(newRichTextValue);
    Logger.log(`updateDaysLeftCellTODO(): updated cell with value: ${newRichTextValue.getText()}`);

    // Set a custom property to store the initial date
    const now = new Date();
    PropertiesService.getDocumentProperties().setProperty(range.getA1Notation(), now.toISOString());
    Logger.log(`updateDaysLeftCellTODO(): set custom property for cell ${range.getA1Notation()}: ${now.toISOString()}`);

    Logger.log(`Updated days left for cell ${range.getA1Notation()}: ${newText}`);
}

/**
 * Parses the number of days left from a given value.
 * 
 * @param {string} value - The value to parse for days left.
 * @returns {number} The number of days left, or 60 if not parseable.
 */
function parseDaysLeftTODO(value) {
    Logger.log(`parseDaysLeftTODO called with value: ${value}`);
    const daysLeftMatch = value.match(/\((\d+)\) days left/); // regex to match the days left pattern
    if (daysLeftMatch) {
        Logger.log(`parseDaysLeftTODO(): parsed days left: ${daysLeftMatch[1]}`);
        return parseInt(daysLeftMatch[1]);
    } else if (/^\d+$/.test(value.trim())) { // regex to check if the value is a number
        Logger.log(`parseDaysLeftTODO(): parsed days left: ${value.trim()}`);
        return parseInt(value.trim());
    }
    const defaultDays = 60;
    Logger.log(`parseDaysLeftTODO(): default days left: ${defaultDays}`);
    return defaultDays;
}

/**
 * Removes multiple dates from cells, keeping only the last occurrence of today's date.
 * 
 * @customfunction
 */
function removeMultipleDatesTODO() {
    Logger.log('removeMultipleDatesTODO called');
    const dataRange = getDataRange();
    const lastRow = Math.min(40, dataRange.getLastRow()); // Limit to 40 rows for performance
    const columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];

    for (const column of columns) {
        for (let row = 2; row <= lastRow; row++) {
            const cell = sheet.getRange(`${column}${row}`);
            const richTextValue = cell.getRichTextValue();
            let text = richTextValue ? richTextValue.getText() : cell.getValue();

            Logger.log(`removeMultipleDatesTODO(): Checking cell ${column}${row}: ${text}`);

            const dateMatches = text.match(/\d{2}\/\d{2}\/\d{2}/g);
            if (dateMatches && dateMatches.length > 1) {
                Logger.log(`removeMultipleDatesTODO(): Found dates in ${column}${row}: ${dateMatches.join(', ')}`);

                // Keep only the last date
                const lastDate = dateMatches[dateMatches.length - 1];
                text = text.replace(/\d{2}\/\d{2}\/\d{2}/g, '').trim();
                text += `\n${lastDate}`;

                Logger.log(`removeMultipleDatesTODO(): Updated text for ${column}${row}: ${text}`);

                let builder = SpreadsheetApp.newRichTextValue().setText(text);
                let currentPos = 0;

                // Apply the same text styles to the new text
                const lines = text.split('\n');
                for (let i = 0; i < lines.length; i++) {
                    const part = lines[i];
                    let startPos = currentPos;
                    let endPos = startPos + part.length;

                    if (/\d{2}\/\d{2}\/\d{2}/.test(part)) {
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

/**
 * Handles the expiration date in a cell.
 * If the cell contains an expiration date in the format **dd/MM/yyyy**, it calculates the number of days left
 * and updates the cell with the new information.
 * 
 * @param {Range} range - The cell range to check for expiration date.
 * @param {string} originalValue - The original value of the cell.
 * @param {string} newValue - The new value of the cell.
 * @param {string} columnLetter - The letter of the column.
 * @param {number} row - The row number.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object for the edit trigger.
 * @return {boolean} True if the expiration date was found and updated, false otherwise.
 */
function handleExpirationDateTODO(range, originalValue, newValue, columnLetter, row, e) {
    Logger.log(`handleExpirationDateTODO called for cell ${columnLetter}${row}`);
    const expiresDatePattern = /\*\*(\d{2}\/\d{2}\/\d{4})\*\*/;
    const match = newValue.match(expiresDatePattern);

    if (match) {
        const dateString = match[1];
        const daysLeft = calcExpirationDaysTODO(dateString);
        Logger.log(`Calculated days left: ${daysLeft} for date: ${dateString}`);

        if (isNaN(daysLeft)) {
            Logger.log('Error: daysLeft is NaN');
            return;
        }

        // Save the expiration date in the cell note
        const expirationDate = parseFullYearDate(dateString);
        range.setNote(`Expiration Date: ${expirationDate.toISOString()}`);

        // Remove the old expiration date and reset text
        let updatedText = newValue.replace(expiresDatePattern, '').trim();
        updatedText = updatedText.replace(/\d{2}\/\d{2}\/\d{2}/g, '').trim(); // Remove any other dates

        // Replace any existing expiration information
        const expiresTextPattern = /Expires in \(\d+\) days|Expires today|EXPIRED/;
        updatedText = updatedText.replace(expiresTextPattern, '').trim();

        // Get today's date
        const today = new Date();
        const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yy");

        // Add the new expiration information based on daysLeft
        if (daysLeft > 0) {
            updatedText += `\nExpires in (${daysLeft}) days\n${todayFormatted}`;
        } else if (daysLeft === 0) {
            updatedText += `\nExpires today\n${todayFormatted}`;
        } else {
            updatedText += "\nEXPIRED";
        }

        range.setValue(updatedText);

        Logger.log(`Updated cell ${columnLetter}${row} with expiration information: ${updatedText}`);

        updateRichTextTODO(range, originalValue, updatedText, columnLetter, row, e);

        if (!updatedText.includes('☑️')) {
            Logger.log(`Adding default checkbox to cell ${columnLetter}${row}`);
            addCheckboxToCellTODO(range);
        }
        return true;
    }

    return false; // No expiration date found
}

/**
 * Calculates the number of days left until the expiration date.
 * 
 * @param {string} dateString - The expiration date in the format dd/MM/yyyy.
 * @returns {number} The number of days left until the expiration date.
 */
function calcExpirationDaysTODO(dateString) {
    Logger.log(`calcExpirationDaysTODO() Triggered: Calculating days left for date: ${dateString}`);

    // Parse the expiration date
    const expirationDate = parseFullYearDate(dateString);
    Logger.log(`Parsed expiration date: ${expirationDate}`);

    // Get today's date and clear time portion
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Reset hours to midnight to compare only dates
    Logger.log(`Today's date (time cleared): ${today}`);

    // Calculate the difference in days between the expiration date and today
    const oneDayInMilliseconds = 24 * 60 * 60 * 1000;
    const daysLeft = Math.floor((expirationDate - today) / oneDayInMilliseconds);
    Logger.log(`Days left: ${daysLeft}`);

    return daysLeft;
}

/**
 * Updates the "Expires in (n) days" text for all cells that contain it.
 * This function should be called on sheet open or reload to ensure that all expiration dates are accurate.
 * 
 * @customfunction
 */
function updateExpirationDatesTODO() {
    Logger.log('updateExpirationDatesTODO called');

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = Math.min(40, sheet.getLastRow()); // Limit to 40 rows for performance

    if (lastRow < 2) {
        Logger.log('No data rows available for processing.');
        return;
    }

    const numRows = lastRow - 1;
    const range = sheet.getRange(2, 3, numRows - 1, 5); // from row 2, Columns C to G
    const values = range.getValues();


    Logger.log(`Total rows being processed: ${numRows}`);

    for (let row = 0; row < values.length; row++) { // Iterating from the start of the range
        for (let col = 0; col < values[row].length; col++) { // Iterate over the selected columns C to G
            const cell = range.getCell(row + 1, col + 1);
            const cellValue = cell.getValue();

            if (typeof cellValue === 'string' && cellValue.includes('Expires')) {
                const note = cell.getNote();

                let expirationDate = null;
                if (note && note.includes('Expiration Date:')) {
                    expirationDate = new Date(note.replace('Expiration Date: ', ''));
                    Logger.log(`Parsed expiration date from note: ${expirationDate}`);
                } else {
                    Logger.log(`No expiration date found in note for cell ${cell.getA1Notation()}`);
                }

                if (expirationDate) {
                    const daysLeft = calcExpirationDaysTODO(Utilities.formatDate(expirationDate, Session.getScriptTimeZone(), "dd/MM/yyyy"));
                    Logger.log(`Calculated days left: ${daysLeft} for expiration date: ${expirationDate}`);

                    let richText = cell.getRichTextValue();
                    let newText = "";

                    if (daysLeft > 0) {
                        newText = `Expires in (${daysLeft}) days`;
                    } else if (daysLeft === 0) {
                        newText = `Expires today`;
                    } else {
                        newText = "EXPIRED";
                    }

                    let textToReplace = /Expires in \(\d+\) days|Expires today|EXPIRED/;
                    let updatedText = cellValue.replace(textToReplace, newText);

                    let builder = SpreadsheetApp.newRichTextValue();
                    builder.setText(updatedText);

                    let runs = richText.getRuns();
                    let offset = 0;

                    for (let i = 0; i < runs.length; i++) {
                        let runText = runs[i].getText();
                        let runStyle = runs[i].getTextStyle();
                        if (offset + runText.length <= updatedText.length) {
                            builder.setTextStyle(offset, offset + runText.length, runStyle);
                        }
                        offset += runText.length;
                    }

                    cell.setRichTextValue(builder.build());

                    Logger.log(`Final updated rich text for cell ${cell.getA1Notation()}: ${updatedText}`);
                }
            }
        }
    }
    storeExpirationDateHashTODO();

    Logger.log('updateExpirationDatesTODO completed');
}

/**
 * Checks if the expiration dates have been updated today.
 * 
 * @returns {boolean} True if the expiration dates have been updated today, false otherwise.
 */
function isExpirationsUpdatedDatesTODO() {
    const today = new Date();
    const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    const dateHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, dateString);
    const hashString = Utilities.base64Encode(dateHash);

    const docProperties = PropertiesService.getDocumentProperties();
    const savedHash = docProperties.getProperty('expirationDateHash');

    return savedHash === hashString;
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

// Contents of ./TODOtriggers/TODOonEdit.js

// globals.js: sheet
// TODOsheet/TODOtoggleFn.js: handlePieChartToggleTODO
// TODOsheet/TODOformatting.js: shiftCellsUpTODO, handleColumnEditTODO, addCheckboxToCellTODO
// TODOsheet/TODOtimeHandle.js: handleExpirationDateTODO
/**
 * Track changes in specified columns and add the date.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The event object for the edit trigger.
 * @customfunction
 */
function onEdit(e) {
    Logger.log('onEdit triggered');
    try {
        let isEnabledDefaultAdditions = PropertiesService.getScriptProperties().getProperty('isEnabledDefaultAdditions') === 'true';
        Logger.log(`isEnabledDefaultAdditions at start of onEdit: ${isEnabledDefaultAdditions}`);

        if (!e || !e.range) {
            Logger.log('Edit event is undefined or does not have range property');
            return;
        }

        const { range } = e;
        const column = range.getColumn();
        const row = range.getRow();
        const columnLetter = String.fromCharCode(64 + column);
        const totalRows = Math.min(40, sheet.getMaxRows()); // Limit to 40 rows

        Logger.log(`onEdit triggered: column ${column}, row ${row}`);
        Logger.log(`isEnabledDefaultAdditions is currently: ${isEnabledDefaultAdditions}`);

        const originalValue = e.oldValue || '';
        const newValue = range.getValue().toString();
        Logger.log(`onEdit(): Original value: "${originalValue}", New value: "${newValue}"`);

        // Check if the cell was cleared (newValue is empty string)
        if (newValue === '') {
            Logger.log(`onEdit(): Clearing format and notes for cell ${range.getA1Notation()}`);
            range.clear({ contentsOnly: false, formatOnly: true, commentsOnly: true });

            // Shift cells up if applicable
            if ((column === 1 || (column >= 3 && column <= 8)) && row >= 2) {
                Logger.log(`onEdit(): Shifting cells up for column ${column}`);
                shiftCellsUpTODO(column, 2, totalRows);
            }
            return;
        }

        if (column === 9 && row === 1) {
            handlePieChartToggleTODO(range);
            return;
        }

        // only handle default additions if enabled
        if (isEnabledDefaultAdditions) {
            Logger.log('Default additions are enabled.');

            if (handleExpirationDateTODO(range, originalValue, newValue, columnLetter, row, e)) {
                Logger.log('onEdit()/handleExpirationDate(): Handled expiration date');
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
        } else {
            Logger.log('Default additions are disabled.');
        }
    } catch (error) {
        Logger.log(`Error in onEdit: ${error.message}`);
        Logger.log(`Error stack: ${error.stack}`);
    }
}

/**
 * Enable default additions on cell edit.
 * @customfunction
 */
function enableDefaultAdditionsTODO() {
    Logger.log('enableDefaultAdditionsTODO called');
    PropertiesService.getScriptProperties().setProperty('isEnabledDefaultAdditions', 'true');
    Logger.log('Default additions on cell edit are enabled.');
}

/**
 * Disable default additions on cell edit.
 * @customfunction
 */
function disableDefaultAdditionsTODO() {
    Logger.log('disableDefaultAdditionsTODO called');
    PropertiesService.getScriptProperties().setProperty('isEnabledDefaultAdditions', 'false');
    Logger.log('Default additions on cell edit are disabled.');
}

// Contents of ./TODOtriggers/TODOonOpen.js

// globals.js: ui, isLoaded
// shared/utils.js: getSheetContentHash, shouldRunUpdates
// shared/formatting: applyFormatToSelected, applyFormatToAll
// TODOsheet/TODOformatting.js: applyFormatToAllTODO, customCeilBGColorTODO, createPieChartTODO, deleteAllChartsTODO, updateDateColorsTODO, setupDropdownTODO, pushUpEmptyCellsTODO, updateCellCommentTODO, removeMultipleDatesTODO, updateDaysLeftTODO
// TODOsheet/TODOcheckbox.js: addCheckboxToCellTODO, addCheckboxesToSelectedCellsTODO, markCheckboxSelectedCellsTODO, markAllCheckboxesSelectedCellsTODO, removeCheckboxesFromSelectedCellsTODO
// TODOsheet/TODOtimeHandle.js: updateExpirationDatesTODO, updateDaysLeftCounterTODO

/**
 * Initializes the UI menu in the spreadsheet.
 * Sets up custom menus and triggers functions when menu items are clicked.
 *
 * @customfunction
 */
function onOpen() {
    Logger.log('onOpen triggered');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const docProperties = PropertiesService.getDocumentProperties();
    const language = docProperties.getProperty('language') || 'English';

    // saveSnapshotTODO() // point A

    Logger.log('Current language: ' + language);
    try {
        const docProperties = PropertiesService.getDocumentProperties();
        const lastHash = docProperties.getProperty('lastHash');
        const currentHash = getSheetContentHash();
        const isExpirationDatesUpdated = isExpirationsUpdatedDatesTODO();
        createMenusTODO();
        if (shouldRunUpdates(lastHash, currentHash) || !isExpirationDatesUpdated) {
            isLoaded = false
            ss.toast(toastMessages.loading[language], 'Status:', 45);
            applyGridLoaderTODO(sheet);
            runAllFunctionsTODO();
            docProperties.setProperty('lastHash', currentHash);
            Logger.log('Running all update functions');
            ss.toast(toastMessages.updateComplete[language], 'Status:', 5);
        } else {
            isLoaded = true
            ss.toast(toastMessages.updateComplete[language], 'Status:', 5);
            Logger.log('It is not necessary to run all functions, the data has not changed significantly.');
        }

    } catch (e) {
        Logger.log('Error: ' + e.toString());
        ui.alert('Error during processing: ' + e.toString());
    }
}

/**
 * Applies a grid loader to the sheet.
 * Adds a red border to the first 20 rows and 8 columns.
 * Used to indicate that the sheet is loading.
 * 
 * @param {Sheet} sheet - The sheet to apply the grid loader to.
 * @customfunction
 */
function applyGridLoaderTODO(sheet) {
    const startRow = 1;
    const endRow = 21;
    const startColumn = 1;
    const endColumn = 8;

    const range = sheet.getRange(startRow, startColumn, endRow, endColumn);
    range.setBorder(true, true, true, true, true, true, '#FF0000', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Runs all functions needed to update the TODO sheet.
 * Calls multiple formatting and update functions.
 *
 * @customfunction
 */
function runAllFunctionsTODO() {
    Logger.log('runAllFunctionsTODO triggered');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    setupDropdownTODO();
    removeMultipleDatesTODO();
    pushUpEmptyCellsTODO();
    updateDaysLeftCounterTODO();
    updateExpirationDatesTODO();
    translateSheetTODO();
    customCellBGColorTODO();
    updateCellCommentTODO();
    updateTipsCellTODO();
    removeNotesFromEmptyCells(sheet) //could be dispensable (gaining load time) since it is difficult for a cell to remain empty with notes given onEdit functionality.
    applyFormatToAllTODO(); // overwrites the grid loader. Must always be the last function called.
    Logger.log('All functions called successfully!');
}

