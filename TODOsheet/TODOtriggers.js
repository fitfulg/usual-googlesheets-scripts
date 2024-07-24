 
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

        const { range } = e;
        const column = range.getColumn();
        const row = range.getRow();
        const columnLetter = String.fromCharCode(64 + column);
        const totalRows = sheet.getMaxRows();

        Logger.log(`onEdit triggered: column ${column}, row ${row}`);

        // Check if the edited cell is for toggling the pie chart (cell I1)
        if (column === 9 && row === 1) {
            handlePieChartToggle(range);
            return;
        }

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
            updateRichTextTODO(range, originalValue, newValue, columnLetter, row, e);
        }
        Logger.log('Calling removeMultipleDatesTODO from onEdit');
        removeMultipleDatesTODO();
    } catch (error) {
        Logger.log(`Error in onEdit: ${error.message}`);
    }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        onEdit
    }
}














