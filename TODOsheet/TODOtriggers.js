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

        const range = e.range;
        const column = range.getColumn();
        const row = range.getRow();
        const columnLetter = String.fromCharCode(64 + column);
        const totalRows = sheet.getMaxRows();

        Logger.log(`onEdit triggered: column ${column}, row ${row}`);

        // Check column for the toggle piechart action
        if (column === 9 && row === 1) {
            const action = range.getValue().toString().trim();
            Logger.log(`Action selected: ${action}`);
            if (action === 'Show Piechart' || action === 'Hide Piechart') {
                togglePieChartTODO(action);
            } else {
                Logger.log('Invalid action selected');
            }
            sheet.getRange("I1").setValue("Piechart");
            return;
        }

        // Store original value before edit
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
            const cellValue = newValue;
            Logger.log(`Cell value after edit: ${cellValue}`);

            let richTextValue = range.getRichTextValue();
            const text = richTextValue ? richTextValue.getText() : cellValue;

            // Store old rich text value and URLs
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

            const date = new Date();
            const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");
            const dateFormatted = ` ${formattedDate}`;

            const newRichTextValue = datePattern.test(text)
                ? updateDateWithStyle(text, dateFormatted, columnLetter, dateColorConfig)
                : appendDateWithStyle(text, dateFormatted, columnLetter, dateColorConfig);

            Logger.log(`Setting rich text value for cell ${columnLetter}${row}`);
            range.setRichTextValue(newRichTextValue);

            // Preserve URLs after updating rich text value
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
    } catch (error) {
        Logger.log(`Error in onEdit: ${error.message}`);
    }
}














