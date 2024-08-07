/* eslint-disable no-unused-vars */

// globals.js: sheet
// TODOsheet/TODOlibrary.js: languages

const setLanguageEnglish = () => setLanguage('English');
const setLanguageSpanish = () => setLanguage('Spanish');
const setLanguageCatalan = () => setLanguage('Catalan');

function setLanguage(language) {
    Logger.log('setLanguage called with language: ' + language);
    if (languages[language]) {
        PropertiesService.getDocumentProperties().setProperty('language', language);
        translateSheetTODO();
        const ui = SpreadsheetApp.getUi();
        const message = {
            'English': 'Language changed.\n Please reload the sheet to update menus.',
            'Spanish': 'Idioma cambiado.\n Por favor, recargue la hoja para actualizar los menús.',
            'Catalan': 'Idioma canviat.\n Si us plau, recarregui el full per actualitzar els menús.'
        };
        ui.alert(message[language]);
    } else {
        Logger.log('Language not supported: ' + language);
    }
}

/**
 * Translates the sheet to the selected language
 * @returns {void}
 * @customfunction
 */
function translateSheetTODO() {
    Logger.log('translateSheetTODO called');
    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';

    // Update with the corresponding styles
    for (const cell in cellStyles) {
        const cellData = cellStyles[cell];
        if (cellData.value[language]) {
            let range = sheet.getRange(cell);
            range.setValue(cellData.value[language])
                .setFontWeight(cellData.fontWeight)
                .setFontColor(cellData.fontColor)
                .setHorizontalAlignment(cellData.alignment);

            if (cellData.backgroundColor) {
                range.setBackground(cellData.backgroundColor);
            }
        }
    }

    // Update the example texts with the corresponding language
    const range = sheet.getDataRange();
    const values = range.getValues();
    for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
            for (const exampleKey in exampleTexts) {
                if (typeof values[i][j] === 'string' && values[i][j].startsWith("Example:")) {
                    const exampleData = exampleTexts[exampleKey];
                    if (exampleData.text[language]) {
                        sheet.getRange(i + 1, j + 1).setValue(exampleData.text[language]);
                    }
                }
            }
        }
    }
}

// for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        setLanguageEnglish,
        setLanguageSpanish,
        setLanguageCatalan,
        setLanguage,
        translateSheetTODO
    }
}