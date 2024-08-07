

// globals.js: sheet

/**
 * Adds by default a checkbox to a cell while preserving existing rich text styles and links.
 * @param {Range} range - The range of the cell to which the checkbox is added.
 * @customfunction
 * @returns {void}
 */
function addCheckboxToCellTODO(range) {
    const cellValue = range.getValue().toString();
    const richTextValue = range.getRichTextValue() || SpreadsheetApp.newRichTextValue().setText(cellValue).build();

    // Check if any checkbox is already present at the beginning
    if (cellValue.startsWith('☑️') || cellValue.startsWith('✅')) {
        Logger.log(`Checkbox already present at the start of cell ${range.getA1Notation()}`);
        return;
    }

    const newRichTextValueBuilder = SpreadsheetApp.newRichTextValue().setText('☑️' + cellValue);

    // Apply style to the checkbox
    newRichTextValueBuilder.setTextStyle(0, 2, SpreadsheetApp.newTextStyle().setBold(true).build());

    // Preserve existing text styles and links starting from the next character
    for (let i = 0; i < cellValue.length; i++) {
        const textStyle = richTextValue.getTextStyle(i, i + 1);
        const url = richTextValue.getLinkUrl(i, i + 1);
        newRichTextValueBuilder.setTextStyle(i + 2, i + 3, textStyle);
        if (url) {
            newRichTextValueBuilder.setLinkUrl(i + 2, i + 3, url);
        }
    }

    range.setRichTextValue(newRichTextValueBuilder.build());
    Logger.log(`Checkbox added to the start of cell ${range.getA1Notation()}`);
}

/**
 * Adds a checkbox to all selected cells, preserving existing rich text styles and links.
 * @customfunction
 * @returns {void}
 */
function addCheckboxesTODO() {
    const range = sheet.getActiveRange();
    const richTextValues = range.getRichTextValues();

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellValue = richTextValues[row][col];
            if (cellValue) {
                const originalText = cellValue.getText();
                Logger.log(`Original cell text: "${originalText}"`);

                // Check if the cell contains only the default checkbox
                const onlyDefaultCheckbox = originalText === '☑️';

                // If only the default checkbox is present, replace it with two checkboxes
                let newText;
                if (onlyDefaultCheckbox) {
                    newText = '☑️☑️';
                    Logger.log('Only default checkbox found, replacing with two checkboxes.');
                } else {
                    // Otherwise, add an additional checkbox
                    newText = `☑️${originalText}`;
                    Logger.log(`New text with added checkbox: "${newText}"`);
                }

                const builder = SpreadsheetApp.newRichTextValue().setText(newText);

                // Preserve existing styles for the rest of the text
                for (let i = 0; i < originalText.length; i++) {
                    const style = cellValue.getTextStyle(i, i + 1);
                    builder.setTextStyle(i + (onlyDefaultCheckbox ? 1 : 2), i + (onlyDefaultCheckbox ? 2 : 3), style);

                    const url = cellValue.getLinkUrl(i, i + 1);
                    if (url) {
                        builder.setLinkUrl(i + (onlyDefaultCheckbox ? 1 : 2), i + (onlyDefaultCheckbox ? 2 : 3), url);
                    }
                }

                // Set the new rich text value for the cell
                range.getCell(row + 1, col + 1).setRichTextValue(builder.build());
            }
        }
    }
    Logger.log("Checkboxes added to selected cells.");
}


/**
 * Changes the first checkbox in each selected cell to a green checkbox.
 * @customfunction
 * @returns {void}
*/
function markCheckboxTODO() {
    const range = sheet.getActiveRange();
    const richTextValues = range.getRichTextValues();

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellValue = richTextValues[row][col];
            if (cellValue) {
                let newText = cellValue.getText();
                const firstCheckboxIndex = newText.indexOf('☑️');
                if (firstCheckboxIndex !== -1) {
                    // Change first checkbox to green checkbox
                    newText = newText.substring(0, firstCheckboxIndex) + '✅' + newText.substring(firstCheckboxIndex + 2);

                    // Create new rich text builder with updated checkbox
                    const builder = SpreadsheetApp.newRichTextValue().setText(newText);

                    // Preserve existing styles
                    for (let i = 0; i < newText.length; i++) {
                        const style = cellValue.getTextStyle(i, i + 1);
                        builder.setTextStyle(i, i + 1, style);

                        const url = cellValue.getLinkUrl(i, i + 1);
                        if (url) {
                            builder.setLinkUrl(i, i + 1, url);
                        }
                    }

                    // Set the new rich text value for the cell
                    range.getCell(row + 1, col + 1).setRichTextValue(builder.build());
                }
            }
        }
    }
    Logger.log("One checkbox changed to green in selected cells.");
}

/**
 * Changes all checkboxes in each selected cell to green checkboxes.
 * @customfunction
 * @returns {void}
*/
function markAllCheckboxesTODO() {
    const range = sheet.getActiveRange();
    const richTextValues = range.getRichTextValues();

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellValue = richTextValues[row][col];
            if (cellValue) {
                let newText = cellValue.getText();
                // Change all checkboxes to green checkboxes
                newText = newText.replace(/☑️/g, '✅');

                // Create new rich text builder with updated checkboxes
                const builder = SpreadsheetApp.newRichTextValue().setText(newText);

                // Preserve existing styles
                for (let i = 0; i < newText.length; i++) {
                    const style = cellValue.getTextStyle(i, i + 1);
                    builder.setTextStyle(i, i + 1, style);

                    const url = cellValue.getLinkUrl(i, i + 1);
                    if (url) {
                        builder.setLinkUrl(i, i + 1, url);
                    }
                }

                // Set the new rich text value for the cell
                range.getCell(row + 1, col + 1).setRichTextValue(builder.build());
            }
        }
    }
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
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const range = sheet.getActiveRange();
        const richTextValues = range.getRichTextValues();

        Logger.log(`Starting to process range: ${range.getA1Notation()}`);

        for (let row = 0; row < richTextValues.length; row++) {
            for (let col = 0; col < richTextValues[row].length; col++) {
                const cellValue = richTextValues[row][col];

                if (cellValue) {
                    let newText = cellValue.getText();
                    Logger.log(`Original text for cell (${row + 1}, ${col + 1}): ${newText}`);

                    // Replace all green checkboxes with default checkboxes
                    newText = newText.replace(/✅/g, '☑️');

                    Logger.log(`Updated text for cell (${row + 1}, ${col + 1}): ${newText}`);

                    // Create a new rich text builder with updated checkboxes
                    const builder = SpreadsheetApp.newRichTextValue().setText(newText);

                    // Preserve existing styles and links
                    for (let i = 0; i < newText.length; i++) {
                        try {
                            const style = cellValue.getTextStyle(i, i + 1);
                            if (style) {
                                builder.setTextStyle(i, i + 1, style);
                                Logger.log(`Applied style from position ${i} to ${i + 1} for cell (${row + 1}, ${col + 1})`);
                            }

                            const url = cellValue.getLinkUrl(i, i + 1);
                            if (url) {
                                builder.setLinkUrl(i, i + 1, url);
                                Logger.log(`Applied link from position ${i} to ${i + 1} for cell (${row + 1}, ${col + 1})`);
                            }
                        } catch (innerError) {
                            Logger.log(`Error applying style or link at position ${i} for cell (${row + 1}, ${col + 1}): ${innerError.message}`);
                        }
                    }

                    // Set the new rich text value for the cell
                    range.getCell(row + 1, col + 1).setRichTextValue(builder.build());
                } else {
                    Logger.log(`Empty cell or no rich text value at (${row + 1}, ${col + 1})`);
                }
            }
        }
        Logger.log("All checkboxes restored to default in selected cells.");
    } catch (e) {
        Logger.log(`Error in restoreCheckboxesTODO: ${e.message}`);
    }
}

/**
 * Removes all checkboxes from the selected cells while preserving existing rich text styles and links.
 * @customfunction
 * @returns {void}
 */
function removeCheckboxesTODO() {
    Logger.log("removeCheckboxesTODO triggered");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getActiveRange();
    const richTextValues = range.getRichTextValues();

    richTextValues.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
            const cellText = cell.getText();
            Logger.log(`Processing cell at row ${rowIndex + 1}, column ${colIndex + 1}`);
            Logger.log(`Original cell text: "${cellText}"`);

            const newText = cellText.replace(/☑️|✅/g, '');
            Logger.log(`Text after checkbox removal: "${newText}"`);

            const builder = SpreadsheetApp.newRichTextValue().setText(newText);

            // Apply existing text styles
            Logger.log('removeCheckboxesTODO(): Applying existing text styles to the cell.');
            for (let i = 0; i < newText.length; i++) {
                const textStyle = cell.getTextStyle(i, i + 1);
                builder.setTextStyle(i, i + 1, textStyle);
                Logger.log(`Applied text style from position ${i} to ${i + 1}.`);
            }

            // Restore existing links
            Logger.log('removeCheckboxesTODO(): Restoring existing links to the cell.');
            for (let i = 0; i < newText.length; i++) {
                const originalIndex = cellText.indexOf(newText[i]);
                if (originalIndex !== -1) {
                    const url = cell.getLinkUrl(originalIndex, originalIndex + 1);
                    if (url) {
                        Logger.log(`Url found at position ${i}: ${url}`);
                        builder.setLinkUrl(i, i + 1, url);
                        Logger.log(`Restored ${url} at position ${i}.`);
                    }
                }
            }

            range.getCell(rowIndex + 1, colIndex + 1).setRichTextValue(builder.build());
            Logger.log(`Checkboxes removed from selected cells.`);
        });
    });
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