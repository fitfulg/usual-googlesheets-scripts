/* eslint-disable no-unused-vars */
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
    const dataRange = getDataRange();
    const lastRow = dataRange.getLastRow();

    for (const column of columns) {
        const config = dateColorConfig[column];
        for (let row = 2; row <= lastRow; row++) {
            const cell = sheet.getRange(`${column}${row}`);
            const cellValue = cell.getValue();
            Logger.log(`updateDateColorsTODO(): Checking if cell ${cellValue} contains a date that matches the pattern ${datePattern}`);
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
    const lastUpdateDate = properties.getProperty('lastUpdateDate');
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    Logger.log(`Last update was on: ${lastUpdateDate}`);
    Logger.log(`Today's date is: ${today}`);

    const now = new Date();
    const endOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
    const hoursLeftUntilNextUpdate = (endOfDay - now) / 3600000; // milliseconds to hours

    Logger.log(`Hours left until the next days left update: ${hoursLeftUntilNextUpdate.toFixed(2)} hours`);

    if (lastUpdateDate === today) {
        Logger.log("No days left counter update needed yet.");
        return;
    }

    const range = sheet.getRange('H2:H' + sheet.getLastRow());
    const values = range.getValues();
    Logger.log("Starting to update days left for each cell.");

    for (let i = 0; i < values.length; i++) {
        const cellValue = values[i][0].toString();
        const match = cellValue.match(/\((\d+)\) days left/);
        if (match) {
            const originalDays = parseInt(match[1]);
            const daysLeft = Math.max(0, originalDays - 1);

            Logger.log(`Row ${i + 2}: original days left = ${originalDays}, new days left = ${daysLeft}`);

            if (daysLeft <= 0) {
                values[i][0] = '';
                Logger.log(`Row ${i + 2}: Days left counter reached zero, clearing cell.`);
            } else {
                values[i][0] = `(${daysLeft}) days left`;
            }
        }
    }
    range.setValues(values);
    properties.setProperty('lastUpdateDate', today);
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
            const cellValue = cell.getValue();
            const richTextValue = cell.getRichTextValue();
            const text = richTextValue ? richTextValue.getText() : cellValue;

            Logger.log(`removeMultipleDatesTODO(): Checking cell ${column}${row}: ${text}`);

            const dateMatches = text.match(/\d{2}\/\d{2}\/\d{2}/g);
            if (dateMatches && dateMatches.length > 1) {
                Logger.log(`removeMultipleDatesTODO(): Found dates in ${column}${row}: ${dateMatches.join(', ')}`);

                const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy");
                Logger.log(`removeMultipleDatesTODO(): Today is: ${today}`);

                // filter and keep only the last occurrence of today's date
                const datesToKeep = [today];
                for (let date of dateMatches) {
                    if (date !== today) {
                        datesToKeep.push(date);
                    }
                }
                Logger.log(`removeMultipleDatesTODO(): Dates to keep: ${datesToKeep.join(', ')}`);

                // create updated text with only the last occurrence of today's date
                let updatedText = text;
                for (let date of datesToKeep) {
                    let lastOccurrence = updatedText.lastIndexOf(date);
                    if (lastOccurrence !== -1) {
                        updatedText = updatedText.substring(0, lastOccurrence) + updatedText.substring(lastOccurrence).replace(new RegExp(date, 'g'), '');
                    }
                }
                Logger.log(`removeMultipleDatesTODO(): Updated text for ${column}${row}: ${updatedText}`);

                updatedText = updatedText.replace(new RegExp(`\\b(${dateMatches.join('|')})\\b`, 'g'), '').trim() + `\n${today}`;
                Logger.log(`removeMultipleDatesTODO(): Updated text for ${column}${row}: ${updatedText}`);

                // build new rich text value with updated text
                let builder = SpreadsheetApp.newRichTextValue().setText(updatedText);
                let currentPos = 0;

                // apply styles to the updated text
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
                Logger.log(`removeMultipleDatesTODO(): Updated rich text value for ${column}${row}: ${builder.build().getText()}`);

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
        updateDateColorsTODO,
        removeMultipleDatesTODO,
        updateDaysLeftCounterTODO,
        updateDaysLeftCellTODO,
        parseDaysLeftTODO
    }
}