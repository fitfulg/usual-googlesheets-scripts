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

        // Check if the edited cell is for toggling the pie chart (cell I1)
        if (column === 9 && row === 1) {
            handlePieChartToggleTODO(range);
            return;
        }

        const originalValue = e.oldValue || '';
        const newValue = range.getValue().toString();

        Logger.log(`onEdit(): Original value: "${originalValue}", New value: "${newValue}"`);

        // Shift cells up if the edited cell is now empty
        if ((column === 1 || (column >= 3 && column <= 8)) && row >= 2 && newValue.trim() === '') {
            Logger.log(`onEdit(): Shifting cells up for column ${column}`);
            shiftCellsUpTODO(column, 2, totalRows);
            return;
        }

        // Handle edits in different columns
        if (row >= 2 && column >= 3 && column <= 8) {
            Logger.log(`onEdit()/handleColumnEditTODO(): Handling column edit for column ${column}`);
            handleColumnEditTODO(range, originalValue, newValue, columnLetter, row, e);
            // Only add a checkbox if the newValue is non-empty and doesn't already contain a checkbox
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

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        onEdit
    }
}
