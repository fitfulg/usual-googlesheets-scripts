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
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        updateDateColorsTODO,
        removeMultipleDatesTODO,
        updateDaysLeftCounterTODO,
        updateDaysLeftCellTODO,
        parseDaysLeftTODO
    }
}