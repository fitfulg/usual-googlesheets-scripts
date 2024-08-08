

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
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        addCheckboxToCellTODO,
        addCheckboxesTODO,
        removeCheckboxesTODO,
        markCheckboxTODO,
        markAllCheckboxesTODO,
        restoreCheckboxesTODO
    };
}