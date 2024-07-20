// globals.js: sheet, datePattern, getDataRange
// shared/formatting.js: resetTextStyle, appendDateWithStyle, updateDateWithStyle

// Track changes in specified columns and add the date
function onEdit(e) {
    const range = e.range;
    const column = range.getColumn();
    const row = range.getRow();
    const columnLetter = String.fromCharCode(64 + column);

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

    // Check column for the toggle dates action
    if (column === 9 && row === 2) {
        const action = range.getValue().toString().trim();
        Logger.log(`Action selected: ${action}`);
        if (action === 'Show Dates' || action === 'Hide Dates') {
            toggleDatesTODO(action);
        } else {
            Logger.log('Invalid action selected');
        }
        sheet.getRange("I2").setValue("Date Toggle");
        return;
    }

    // Check if the edit is in columns C, D, E, F, G, H and from row 2 onwards
    if (column >= 3 && column <= 8 && row >= 2) {
        const cellValue = range.getValue();

        if (cellValue.trim() === "") return resetTextStyle(range);

        const date = new Date();
        const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");
        const dateFormatted = ` ${formattedDate}`;

        const richTextValue = datePattern.test(cellValue)
            ? updateDateWithStyle(cellValue, dateFormatted, columnLetter, dateColorConfig)
            : appendDateWithStyle(cellValue, dateFormatted, columnLetter, dateColorConfig);

        range.setRichTextValue(richTextValue);
    }
}
