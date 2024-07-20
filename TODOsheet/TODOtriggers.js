// globals.js: sheet, datePattern, getDataRange
// shared/formatting.js: resetTextStyle, appendDateWithStyle, updateDateWithStyle

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

            // Skip if value has not changed
            if (originalValue === newValue) {
                Logger.log('No change in cell value, skipping update');
                return;
            }

            // Get current rich text value to preserve formatting
            const richTextValue = range.getRichTextValue();
            const text = richTextValue ? richTextValue.getText() : cellValue;

            if (text.trim() === "") return resetTextStyle(range);

            const date = new Date();
            const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");
            const dateFormatted = ` ${formattedDate}`;

            const newRichTextValue = datePattern.test(text)
                ? updateDateWithStyle(text, dateFormatted, columnLetter, dateColorConfig)
                : appendDateWithStyle(text, dateFormatted, columnLetter, dateColorConfig);

            Logger.log(`Setting rich text value for cell ${columnLetter}${row}`);
            range.setRichTextValue(newRichTextValue);
        }
    } catch (error) {
        Logger.log(`Error in onEdit: ${error.message}`);
    }
}







