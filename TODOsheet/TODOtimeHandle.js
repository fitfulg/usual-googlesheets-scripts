/* eslint-disable no-unused-vars */
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

    for (const column of columns) {
        const config = dateColorConfig[column];
        for (let row = 2; row <= lastRow; row++) {
            const cell = sheet.getRange(`${column}${row}`);
            const cellValue = cell.getValue();
            Logger.log(`updateDateColorsTODO(): Checking cell ${cell.getA1Notation()} for date and expiration`);

            // Extract the date from the cell
            let dateText = cellValue.match(datePattern);
            if (dateText) {
                dateText = dateText[0];
                const cellDate = parseDate(dateText);
                const today = new Date();
                const diffDays = Math.floor((cellDate - today) / (1000 * 60 * 60 * 24));

                // Update or add "Expires in (n) days"
                if (diffDays >= 0) {
                    const expiresInText = `Expires in (${diffDays}) days`;
                    if (expiresPattern.test(cellValue)) {
                        cellValue = cellValue.replace(expiresPattern, expiresInText);
                    } else {
                        cellValue += `\n${expiresInText}`;
                    }
                }

                // Update the rich text value with new expiration and date
                const richTextValueBuilder = SpreadsheetApp.newRichTextValue().setText(cellValue);
                const dateIndex = cellValue.indexOf(dateText);
                const expiresIndex = cellValue.indexOf(`Expires in`);

                // Style the date
                richTextValueBuilder.setTextStyle(dateIndex, dateIndex + dateText.length, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor(config.defaultColor).build());

                // Style the "Expires in"
                if (expiresIndex !== -1) {
                    richTextValueBuilder.setTextStyle(expiresIndex, cellValue.length, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor(config.defaultColor).build());
                }

                cell.setRichTextValue(richTextValueBuilder.build());
            }
        }
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

    const range = sheet.getRange('H2:H' + sheet.getLastRow());
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
    const lastRow = dataRange.getLastRow();
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
        // Remove the old expiration date
        let updatedText = newValue.replace(expiresDatePattern, '').trim();
        updatedText = updatedText.replace(/\d{2}\/\d{2}\/\d{2}/g, '').trim(); // Remove any other dates

        const expiresTextPattern = /Expires in \(\d+\) days/;
        updatedText = updatedText.replace(expiresTextPattern, '').trim();

        // Add the new expiration information
        updatedText += `\nExpires in (${daysLeft}) days\n${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy")}`;

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

    const expirationDate = parseFullYearDate(dateString);
    const today = new Date();
    Logger.log(`Expiration date: ${expirationDate}`);
    Logger.log(`Today's date: ${today}`);

    const expirationDateUTC = Date.UTC(expirationDate.getFullYear(), expirationDate.getMonth(), expirationDate.getDate());
    const todayUTC = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
    const timeDiff = expirationDateUTC - todayUTC;
    Logger.log(`Time difference in milliseconds: ${timeDiff}`);

    const daysLeft = Math.ceil(timeDiff / (1000 * 60 * 60 * 24)); // Cálculo de días restantes
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
    const dataRange = getDataRange();
    const lastRow = dataRange.getLastRow();

    for (let row = 2; row <= lastRow; row++) {
        const cell = sheet.getRange(`C${row}`);
        const cellValue = cell.getValue();

        if (cellValue.includes("Expires in")) {
            // Recalcula y actualiza los días restantes
            let expirationDateMatch = cellValue.match(/\d{2}\/\d{2}\/\d{2}$/);
            if (expirationDateMatch) {
                let expirationDate = parseDate(expirationDateMatch[0]);
                let today = new Date();
                let diffDays = Math.floor((expirationDate - today) / (1000 * 60 * 60 * 24));

                if (diffDays >= 0) {
                    cellValue = cellValue.replace(/Expires in \(\d+\) days/, `Expires in (${diffDays}) days`);
                    cell.setValue(cellValue);
                } else {
                    // Optional: clear cell or mark as expired
                    cell.setValue(cellValue.replace(/Expires in \(\d+\) days/, "Expired"));
                }
            }
        }
    }
}

// for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        updateDateColorsTODO,
        removeMultipleDatesTODO,
        updateDaysLeftCounterTODO,
        updateDaysLeftCellTODO,
        parseDaysLeftTODO,
        handleExpirationDateTODO,
        calcExpirationDaysTODO,
        updateExpirationDatesTODO
    }
}