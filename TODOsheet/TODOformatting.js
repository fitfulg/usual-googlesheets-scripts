/* eslint-disable no-unused-vars */
// globals.js: sheet, getDataRange, datePattern
// shared/formatting.js: Format, applyBorders, applyThickBorders, setCellStyle, appendDateWithStyle, updateDateWithStyle, resetTextStyle, clearTextFormatting
// TODOsheet/TODOtimeHandle.js: updateDaysLeftCellTODO
// TODOsheet/TODOlibrary.js: dateColorConfig

/**
 * Updates the colors of the dates based on the priority of the column.
 *  
 * @customfunction
 */
function updateTipsCellTODO() {
    Logger.log('updateTipsCellTODO called');
    const cell = sheet.getRange("I2");

    const tips = {
        English: "💡 To add an expiration date countdown, simply add the date in the format **dd/mm/yyyy** to the desired cell.(Don't forget to add the asterisks)",
        Spanish: "💡 Para añadir una fecha de expiración a modo de cuenta atrás en días, basta con añadir la fecha en formato **dd/mm/yyyy** a la celda en cuestión. (No olvides añadir los asteriscos)",
        Catalan: "💡 Per afegir una data de caducitat en mode compte enrere en dies, només cal afegir la data en format **dd/mm/yyyy** a la cel·la en qüestió. (No oblidis afegir els asteriscs)"
    };

    const titles = {
        English: "💡Tips",
        Spanish: "💡Consejos",
        Catalan: "💡Consells"
    };

    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';
    const tipText = tips[language];
    const titleText = titles[language];

    cell.setNote(tipText);
    cell.setValue(titleText);
    cell.setFontWeight("bold");
    cell.setFontSize(12);
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
    cell.setBackground("#efefef");
    cell.setBorder(true, true, true, true, true, true, '#D3D3D3', SpreadsheetApp.BorderStyle.SOLID_THICK);

    Logger.log('Tips cell updated with tips for language: ' + language);
}

/**
 * Updates the cell comment with the latest changes.
 * 
 * @customfunction
 */
function updateCellCommentTODO() {
    Logger.log('updateCellCommentTODO called');
    const cell = sheet.getRange("I3");
    const version = "v1.3";
    cell.setValue(version);

    const changes = {
        English: `
            NEW FEATURES:
            - You can now add an expiration date to your cells. It is a countdown by days. See the "Tips" cell for more information.
            - Cells with expiration dates come with notes added that are used to calculate and update the expiration days.
            - You can now enable or disable all the functionalities that are added by default when writing in a cell. For example, for a cell to have only text, without checkboxes or default date. From the "TODO Sheet"/"Enable/Disable Default Additions in Cells" menu.
            OLD FEATURES:
            - A checkbox is added by default from the 3rd to the 8th column when a cell is written or modified.
            - You can add, mark, restore, and delete checkboxes in cells by selecting them and using the "Custom Formats" menu.
            - The "days left" counter is updated daily in the 8th column. When the counter reaches zero, the cell is cleared.
            - A snapshot of the sheet can be saved and restored from the "Custom Formats" menu.
            - Indicative limit of cells for each priority, with a warning when the limit is reached.
            - Custom formats can be applied without refreshing the page from the "Custom Formats" menu.
            - Date color change times vary by column priority.
            - The Piechart can be shown or hidden using its dropdown cell.
            - Deleted empty cells are replaced by the immediately lower cell.
        `,
        Spanish: `
            NUEVAS FUNCIONES:
            - Ahora puedes agregar una fecha de vencimiento a tus celdas. Es una cuenta regresiva en días. Consulta la celda "Consejos" para obtener más información.
            - Las celdas con fecha de expiracion vienen con notas agregadas que son usadas para calcular y actualizar los dias de expiración.
            - Ahora puedes activar o desactivar todas las funcionalidades que se añaden por defecto al escribir en una celda. Por ejemplo para que una celda tenga solamente texto, sin casillas ni fecha por defecto. Desde el menu "Hoja TODO"/"Habilitar/Deshabilitar Adiciones por Defecto en las Celdas"
            FUNCIONES ANTIGUAS:
            - Se añade una casilla de verificación por defecto desde la 3ª a la 8ª columna cuando se escribe o modifica una celda.
            - Puedes agregar, marcar, restaurar y eliminar casillas en las celdas seleccionándolas y usando el menú "Formatos personalizados".
            - El contador de "días restantes" se actualiza diariamente en la 8ª columna. Cuando el contador llega a cero, la celda se borra.
            - Se puede guardar y restaurar una instantánea de la hoja desde el menú "Formatos personalizados".
            - Límite indicativo de celdas para cada prioridad, con una advertencia cuando se alcanza el límite.
            - Se pueden aplicar formatos personalizados sin necesidad de refrescar la página desde el menú "Formatos personalizados".
            - Los tiempos de cambio de color de las fechas varían según la prioridad de la columna.
            - El gráfico circular se puede mostrar u ocultar usando su celda desplegable.
            - Las celdas vacías eliminadas son reemplazadas por la celda inmediatamente inferior.
        `,
        Catalan: `
            NOVES FUNCIONS:
            - Ara pots agregar una data de venciment a les cel·les. És un compte enrere en dies. Consulta la cel·la "Consells" per obtenir més informació.
            - Les cel·les amb data d'expiració venen amb notes agregades que són usades per calcular i actualitzar els dies d'expiració.
            - Ara podeu activar o desactivar totes les funcionalitats que s'afegeixen per defecte en escriure en una cel·la. Per exemple perquè una cel·la tingui només text, sense caselles ni data per defecte. Des del menu "Full de TODO"/"Habilitar/Deshabilitar Addicions per Defecte a les Cel·les"
            FUNCIONS ANTIGUES:
            - S'afegeix una casella de verificació per defecte des de la 3a fins a la 8a columna quan s'escriu o es modifica una cel·la.
            - Pots afegir, marcar, restaurar i eliminar caselles en les cel·les seleccionades seleccionant-les i utilitzant el menú "Formats personalitzats".
            - El comptador de "dies restants" s'actualitza diàriament a la 8a columna. Quan el comptador arriba a zero, la cel·la s'esborra.
            - Es pot desar i restaurar una instantània del full des del menú "Formats personalitzats".
            - Límite indicatiu de cel·les per a cada prioritat, amb una advertència quan s'assoleix el límit.
            - Es poden aplicar formats personalitzats sense necessitat de refrescar la pàgina des del menú "Formats personalitzats".
            - Els temps de canvi de color de les dates varien segons la prioritat de la columna.
            - El gràfic circular es pot mostrar o ocultar utilitzant la seva cel·la desplegable.
            - Les cel·les buides eliminades són reemplaçades per la cel·la immediatament inferior.
        `
    };

    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';
    const comment = `Version: ${version}\n${changes[language]}`;

    cell.setComment(comment);
    cell.setBackground("#efefef");
    cell.setBorder(true, true, true, true, true, true, '#D3D3D3', SpreadsheetApp.BorderStyle.SOLID_THICK);
    Format(cell);

    Logger.log('Cell comment updated with changes for language: ' + language);
}

/**
 * Sets example text for a specific column if the cells are empty.
 * 
 * @customfunction
 * @param {string} column - The column to check for empty cells.
 * @param {string} exampleText - The example text to set if cells are empty.
 */
function exampleTextTODO(column, exampleText) {
    Logger.log(`exampleTextTODO called for column: ${column}, example text: ${exampleText}`);
    const dataRange = getDataRange();
    const lastRow = dataRange.getLastRow();  // Get the last row with data
    let values;

    if (column === "B") {
        // Get values excluding B3 and B8
        const firstPart = sheet.getRange(column + "2").getValues().flat();  // Get value from B2
        const middlePart = sheet.getRange(column + "4:" + column + "7").getValues().flat();  // Get values from B4-B7
        const lastPart = sheet.getRange(column + "9:" + column + lastRow).getValues().flat();  // Get values from B9 to the last row

        values = [...firstPart, ...middlePart, ...lastPart];
    } else {
        values = sheet.getRange(column + "2:" + column + lastRow).getValues().flat();  // Get values from the column's 2nd row to the last row
    }

    Logger.log(`Values in column ${column}: ${values}`);

    // Check if the first cell of the column is empty
    const firstCellEmpty = values[0].toString().trim() === '';

    if (firstCellEmpty) {
        const cell = sheet.getRange(column + "2");
        cell.setValue(exampleText);  // Set example text if the first cell is empty
        Logger.log(`Example text set for column ${column} at ${column} 2: ${exampleText}`);
    } else {
        Logger.log(`Column ${column} is not empty at ${column}2, skipping setting example text.`);
    }
}

/**
 * Applies formatting to the entire sheet and sets example text.
 * 
 * @customfunction
 */
function applyFormatToAllTODO() {
    Logger.log('applyFormatToAllTODO called');
    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';
    const totalRows = Math.min(40, sheet.getMaxRows()); // Limited to 40 rows
    const range = sheet.getRange(1, 1, totalRows, 8);

    Logger.log('applyFormatToAllTODO()/preserveRelevantHyperlinks() called');
    const preservedLinks = preserveRelevantHyperlinks(range);

    Logger.log('applyFormatToAllTODO()/updateDateColorsTODO() called');
    updateDateColorsTODO();  // Update date colors selectively

    Logger.log('applyFormatToAllTODO()/setCellContentAndStyle(): setting cell content and style');
    setCellContentAndStyleTODO();

    Logger.log('applyFormatToAllTODO()/exampleTextTODO(): setting example text');
    applyExampleTexts(language);

    if (range) {
        Format(range);
        applyBorders(range);
    }

    Logger.log('applyFormatToAllTODO()/restoreRelevantHyperlinks() called');
    restoreRelevantHyperlinks(range, preservedLinks);

    Logger.log('applyFormatToAllTODO()/applyExpiresTextStyle() called');
    applyExpiresTextStyle();

    Logger.log('applyFormatToAllTODO()/applyThickBorders(): applying thick borders');
    applyThickBorders(sheet.getRange(1, 3, 11, 1));
    applyThickBorders(sheet.getRange(1, 4, 21, 1));
    applyThickBorders(sheet.getRange(1, 5, 21, 1));

    Logger.log('applyFormatToAllTODO()/checkAndSetColumnTODO(): checking and setting columns');
    applyColumnStyles(language);
}

/**
 * Preserves relevant hyperlinks in the specified range.
 * 
 * @param {Range} range - The range to preserve hyperlinks in.
 * @return {Array} The preserved hyperlinks.
 * @customfunction
 */
function preserveRelevantHyperlinks(range) {
    Logger.log('preserveRelevantHyperlinks called');
    const richTextValues = range.getRichTextValues();
    const preservedLinks = [];
    const maxRows = Math.min(richTextValues.length, 40); // Limited to 40 rows

    for (let row = 0; row < maxRows; row++) {
        let rowHasRelevantData = false;
        preservedLinks[row] = [];
        for (let col = 1; col <= 6; col++) { // columns B to H 
            const richText = richTextValues[row][col];
            const cellText = richText.getText().trim();

            if (cellText === '') {
                preservedLinks[row][col] = null;  // Omit empty cells
                continue;
            }

            if (richText.getLinkUrl() || cellText.includes('Expires in')) {
                preservedLinks[row][col] = richText;
                rowHasRelevantData = true;
                Logger.log(`Row ${row + 1}, Column ${col + 1}: Relevant data preserved.`);
            } else {
                preservedLinks[row][col] = null;
            }
        }

        if (!rowHasRelevantData) {
            preservedLinks[row] = null;  // delete rows with no relevant data
        }
    }

    Logger.log(`preserveRelevantHyperlinks completed: Total rows preserved ${preservedLinks.length}`);
    return preservedLinks;
}

/**
 * Restores the relevant hyperlinks in the specified range.
 * 
 * @param {Range} range - The range to restore hyperlinks in.
 * @param {Array} preservedLinks - The preserved hyperlinks.
 * @customfunction
 * @returns {void}
 */
function restoreRelevantHyperlinks(range, preservedLinks) {
    Logger.log('restoreRelevantHyperlinks called');
    const richTextValues = range.getRichTextValues();

    const maxRows = Math.min(preservedLinks.length, 40); // Limited to 40 rows

    for (let row = 0; row < maxRows; row++) {
        if (preservedLinks[row] !== null) {
            for (let col = 1; col <= 6; col++) { // columns B to H 
                if (preservedLinks[row][col] !== null) {
                    richTextValues[row][col] = preservedLinks[row][col];
                    Logger.log(`Row ${row + 1}, Column ${col + 1}: Restoring preserved data.`);
                }
            }
        }
    }

    range.setRichTextValues(richTextValues); // apply all restored rich text values at once
    Logger.log('restoreRelevantHyperlinks completed');
}

/**
 * Sets the example text for each column in the sheet.
 * 
 * @param {language} language - The language to set the content in.
 * @customfunction
 */
function applyExampleTexts(language) {
    for (const column in exampleTexts) {
        const { text } = exampleTexts[column];
        const translatedText = text[language];
        exampleTextTODO(column, translatedText);
        Logger.log(`applyFormatToAllTODO(): example text set for column ${column} - translatedText: ${translatedText}`);
    }
}

/**
 * Sets the column content and style based on the language.
 * 
 * @param {language} language - The language to set the content in.
 * @customfunction
 */
function applyColumnStyles(language) {
    for (const column in cellStyles) {
        const { limit, priority } = cellStyles[column];

        const translatedLimit = limit?.[language];
        const translatedPriority = priority?.[language];

        if (translatedLimit !== undefined && translatedPriority !== undefined) {
            checkAndSetColumnTODO(column.charAt(0), translatedLimit, translatedPriority);
            Logger.log(`applyFormatToAllTODO(): translatedText set for column ${column} - limit: ${translatedLimit}, priority: ${translatedPriority}`);
        } else {
            Logger.log(`applyFormatToAllTODO(): limit or priority not found for column ${column} and language ${language}`);
        }
    }
}

/**
 * Sets the expiration date text style in the sheet.
 * 
 * @customfunction
 * @returns {void}
 */
function applyExpiresTextStyle() {
    Logger.log('applyExpiresTextStyle called');
    const totalRows = Math.min(40, sheet.getMaxRows()); // Limited to 40 rows
    const range = sheet.getRange(1, 1, totalRows, 8);
    const richTextValues = range.getRichTextValues();

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const richText = richTextValues[row][col];
            const text = richText.getText();
            const expiresInIndex = text.indexOf('Expires in');
            const dateMatch = text.match(/\d{2}\/\d{2}\/\d{2}/);  // Match for the date format DD/MM/YY

            if (expiresInIndex !== -1 && dateMatch) {
                const dateIndex = text.indexOf(dateMatch[0]);  // Get the starting index of the date

                // Ensure indices are valid and in order
                if (expiresInIndex < dateIndex && expiresInIndex >= 0 && dateIndex >= 0) {
                    const builder = richText.copy();

                    // Preserve existing text styles
                    const existingStyle = richText.getTextStyle(0, text.length);

                    // Apply the new styles only to the "Expires in" and date sections
                    builder.setTextStyle(0, text.length, existingStyle); // Preserve the existing style
                    builder.setTextStyle(expiresInIndex, dateIndex, SpreadsheetApp.newTextStyle().setForegroundColor('#0000FF').setItalic(true).build());

                    range.getCell(row + 1, col + 1).setRichTextValue(builder.build());
                } else {
                    Logger.log(`Skipping invalid indices for styling in cell ${row + 1}, ${col + 1}`);
                }
            }
        }
    }
}

/**
 * Checks and sets the column based on the limit of occupied cells.
 * 
 * @customfunction
 * @param {string} column - The column to check.
 * @param {number} limit - The limit of occupied cells.
 * @param {string} priority - The priority level.
 */
function checkAndSetColumnTODO(column, limit, priority) {
    Logger.log(`checkAndSetColumnTODO called for column: ${column}, limit: ${limit}, priority: ${priority}`);

    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';

    // Define messages based on the selected language
    const messages = {
        English: {
            cellLimitReached: "⚠️CELL LIMIT REACHED⚠️",
            alertMessage: "⚠️CELL LIMIT REACHED⚠️ \nfor " + priority
        },
        Spanish: {
            cellLimitReached: "⚠️LÍMITE DE CELDAS ALCANZADO⚠️",
            alertMessage: "⚠️LÍMITE DE CELDAS ALCANZADO⚠️ \npara la " + priority
        },
        Catalan: {
            cellLimitReached: "⚠️LÍMIT DE CEL·LES ASSOLIT⚠️",
            alertMessage: "⚠️LÍMIT DE CEL·LES ASSOLIT⚠️ \nper a la " + priority
        }
    };

    const message = messages[language];

    const dataRange = getDataRange();
    const values = sheet.getRange(column + "2:" + column + dataRange.getLastRow()).getValues().flat();
    const occupied = values.filter(String).length;
    const range = sheet.getRange(column + "2:" + column + dataRange.getLastRow());

    if (occupied > limit) {
        // Set red border with thicker style
        range.setBorder(true, true, true, true, true, true, "#FF0000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(column + "1").setValue(message.cellLimitReached);
        SpreadsheetApp.getUi().alert(message.alertMessage);
    } else {
        // Set black border with thicker style
        range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(column + "1").setValue(priority);
    }
}

/**
 * Sets the background color of a specific column.
 * 
 * @customfunction
 * @param {Sheet} sheet - The sheet object.
 * @param {number} col - The column number.
 * @param {string} color - The background color to set.
 * @param {number} [startRow=2] - The starting row number.
 */
function setColumnBackground(sheet, col, color, startRow = 2) {
    Logger.log(`setColumnBackground called for column: ${col}, color: ${color}, startRow: ${startRow}`);
    let totalRows = Math.min(40, sheet.getMaxRows()); // Limited to 40 rows
    let range = sheet.getRange(startRow, col, totalRows - startRow + 1, 1);
    range.setBackground(color);
}

/**
 * Customizes the background colors of specific columns and cells.
 * 
 * @customfunction
 */
function customCellBGColorTODO() {
    Logger.log('customCellBGColorTODO called');
    // Apply background colors to specific columns
    setColumnBackground(sheet, 1, '#d3d3d3', 2); // Column A: Light gray 3
    setColumnBackground(sheet, 6, '#eef7ff', 2); // Column F: Light blue 3
    setColumnBackground(sheet, 7, '#fff1f1', 2); // Column G: Light red 3

    // Apply white background to columns B, C, D, E, H, I starting from row 2
    let whiteColumns = [2, 3, 4, 5, 8, 9]; // Makes cell I2 momentarily white(column 8) while loading rest of the sheet. Useful for testing. Then turns dark gray(updateCellCommentTODO)
    for (let col of whiteColumns) {
        setColumnBackground(sheet, col, '#ffffff', 2);
    }

    // Apply dark yellow background to specific cells in column B
    sheet.getRange('B3').setBackground('#b5a642'); // Dark yellow 3
    sheet.getRange('B8').setBackground('#b5a642'); // Dark yellow 3
}

/**
 * Sets content and style for specific cells based on predefined configurations.
 * 
 * @customfunction
 */
function setCellContentAndStyleTODO() {
    Logger.log('setCellContentAndStyleTODO called');
    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';
    for (const cell in cellStyles) {
        const { value, fontWeight, fontColor, backgroundColor, alignment } = cellStyles[cell];
        const translatedValue = value[language];
        setCellStyle(cell, translatedValue, fontWeight, fontColor, backgroundColor, alignment);
    }
}

/**
 * Sets up a dropdown menu in cell I1 with options to show or hide the pie chart.
 *
 * @customfunction
 */
function setupDropdownTODO() {
    Logger.log('setupDropdownTODO called');
    // Setup dropdown in I1
    const buttonCell = sheet.getRange("I1");
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Piechart', 'Show Piechart', 'Hide Piechart'], true).build();
    buttonCell.setDataValidation(rule);
    buttonCell.setValue('Piechart');
    buttonCell.setFontWeight('bold');
    buttonCell.setFontSize(12);
    buttonCell.setHorizontalAlignment("center");
    buttonCell.setVerticalAlignment("middle");
}

/**
 * Shifts cells up in a column if they are empty, filling with the values below.
 *
 * @customfunction
 * @param {number} column - The column to shift cells up (1-indexed).
 * @param {number} startRow - The starting row number.
 * @param {number} endRow - The ending row number.
 */
function shiftCellsUpTODO(column, startRow, endRow) {
    Logger.log(`shiftCellsUpTODO called for column: ${column}, from row ${startRow} to ${endRow}`);

    const range = sheet.getRange(startRow, column, endRow - startRow + 1, 1);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();

    let hasChanges = false;

    for (let i = 0; i < values.length - 1; i++) {
        if (values[i][0] === '' && values[i + 1][0] !== '') {
            Logger.log(`Empty cell found at row ${i + startRow}, shifting cells up`);

            values[i][0] = values[i + 1][0];
            richTextValues[i][0] = richTextValues[i + 1][0];

            moveNoteToUpperCell(sheet, startRow + i + 1, startRow + i, column);

            values[i + 1][0] = '';
            richTextValues[i + 1][0] = SpreadsheetApp.newRichTextValue().setText('').build();

            hasChanges = true;
            Logger.log(`After shifting: Row ${i + startRow}, New Value: ${values[i][0]}, New RichText: ${richTextValues[i][0].getText()}`);
        }
    }

    if (hasChanges) {
        Logger.log(`Setting values for range: ${startRow} to ${endRow}, column: ${column}`);
        range.setValues(values);
        range.setRichTextValues(richTextValues);
    }

    Logger.log(`shiftCellsUpTODO completed for column: ${column}`);
}

/**
 * Moves the note from one cell to another.
 * 
 * @param {Sheet} sheet - The sheet object.
 * @param {number} fromRow - The row number to move the note from.
 * @param {number} toRow - The row number to move the note to.
 * @param {number} column - The column number.
 * @returns {void}
 * @customfunction
 */
function moveNoteToUpperCell(sheet, fromRow, toRow, column) {
    const note = sheet.getRange(fromRow, column).getNote();
    sheet.getRange(toRow, column).setNote(note);
    sheet.getRange(fromRow, column).setNote('');
}

/**
 * Forces empty cells to shift up in specified columns.
 *
 * @customfunction
 */
function pushUpEmptyCellsTODO() {
    Logger.log('pushUpEmptyCellsTODO called');
    const range = sheet.getDataRange();
    const numRows = Math.min(40, range.getNumRows()); // Limited to 40 rows
    const numCols = range.getNumColumns();

    for (let col = 1; col <= numCols; col++) {
        let startRow = null;
        for (let row = 2; row <= numRows; row++) {
            if (sheet.getRange(row, col).getValue() === '' && startRow === null) {
                startRow = row;
            } else if (sheet.getRange(row, col).getValue() !== '' && startRow !== null) {
                shiftCellsUpTODO(col, startRow, numRows);
                startRow = null;
            }
        }
        // If the last rows in the column are empty
        if (startRow !== null) {
            shiftCellsUpTODO(col, startRow, numRows);
        }
    }
    Logger.log('pushUpEmptyCells completed');
}

/**
 * Updates rich text content of a cell based on original and new values.
 *
 * @customfunction
 * @param {Range} range - The cell range to update.
 * @param {string} originalValue - The original value of the cell.
 * @param {string} newValue - The new value of the cell.
 * @param {string} columnLetter - The column letter of the cell.
 * @param {number} row - The row number of the cell.
 * @param {Event} e - The edit event object.
 */
function updateRichTextTODO(range, originalValue, newValue, columnLetter, row, e) {
    Logger.log(`updateRichTextTODO called for column: ${columnLetter}, row: ${row}, original value: "${originalValue}", new value: "${newValue}"`);

    let updatedText = newValue.toString().trim();
    const dateFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy");
    Logger.log(`Today's date: ${dateFormatted}`);

    let originalRichTextValue;
    try {
        originalRichTextValue = range.getRichTextValue() || SpreadsheetApp.newRichTextValue().setText(originalValue).build();
    } catch (error) {
        Logger.log(`Error getting original rich text value: ${error.message}`);
        return;
    }

    // Add the date if it's not already present
    const datePattern = /\d{2}\/\d{2}\/\d{2}$/;
    if (!datePattern.test(updatedText)) {
        updatedText = `${updatedText}\n${dateFormatted}`;
        Logger.log(`No date found, updated text with new date: "${updatedText}"`);
    }

    // Add a checkbox if it's not already present
    if (!updatedText.startsWith('☑️')) {
        updatedText = `☑️${updatedText}`;
        Logger.log(`Checkbox added to the start of the text: "${updatedText}"`);
    }

    try {
        // Apply the updated text to the cell
        const newRichTextValueBuilder = SpreadsheetApp.newRichTextValue()
            .setText(updatedText);

        // preserve links from the original text
        const originalText = originalRichTextValue.getText();
        Logger.log(`Preserving links from original text: ${originalText}`);
        for (let i = 0; i < originalText.length && i < updatedText.length; i++) {
            const url = originalRichTextValue.getLinkUrl(i, i + 1);
            if (url) {
                newRichTextValueBuilder.setLinkUrl(i, i + 1, url);
            }
        }

        // Apply italic style to the date
        const dateStartIdx = updatedText.search(datePattern);
        if (dateStartIdx !== -1) {
            const dateEndIdx = updatedText.length;
            newRichTextValueBuilder.setTextStyle(dateStartIdx, dateEndIdx, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor('#A9A9A9').build());
        }

        range.setRichTextValue(newRichTextValueBuilder.build());
        Logger.log(`Set new rich text value for cell ${columnLetter}${row}`);
    } catch (error) {
        Logger.log(`Error in updateRichTextTODO: ${error.message}`);
    }
}

/**
 * Handles the editing of a cell based on its column.
 * 
 * @param {Range} range - The cell range that was edited.
 * @param {string} originalValue - The original value of the cell before editing.
 * @param {string} newValue - The new value of the cell after editing.
 * @param {string} columnLetter - The letter of the column that was edited.
 * @param {number} row - The row number of the edited cell.
 * @param {Event} e - The edit event object.
 */
function handleColumnEditTODO(range, originalValue, newValue, columnLetter, row, e) {
    Logger.log(`handleColumnEditTODO called for column: ${columnLetter}, row: ${row}, original value: "${originalValue}", new value: "${newValue}"`);
    if (columnLetter === 'H') {
        let daysLeft = parseDaysLeftTODO(newValue);
        updateDaysLeftCellTODO(range, daysLeft);
    } else {
        updateRichTextTODO(range, originalValue, newValue, columnLetter, row, e);
        removeMultipleDatesTODO();
    }
}

// for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        updateCellCommentTODO,
        exampleTextTODO,
        applyFormatToAllTODO,
        checkAndSetColumnTODO,
        setColumnBackground,
        customCellBGColorTODO,
        setCellContentAndStyleTODO,
        setupDropdownTODO,
        pushUpEmptyCellsTODO,
        updateRichTextTODO,
        shiftCellsUpTODO,
        handleColumnEditTODO,
        updateTipsCellTODO,
        applyExpiresTextStyle,
        applyColumnStyles,
        applyExampleTexts,
        preserveRelevantHyperlinks,
        restoreRelevantHyperlinks
    }
}