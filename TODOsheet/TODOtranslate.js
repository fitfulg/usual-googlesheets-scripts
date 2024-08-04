/* eslint-disable no-unused-vars */

const setLanguageEnglish = () => setLanguage('English');
const setLanguageSpanish = () => setLanguage('Spanish');
const setLanguageCatalan = () => setLanguage('Catalan');

function setLanguage(language) {
    if (languages[language]) {
        PropertiesService.getDocumentProperties().setProperty('language', language);
        translateSheetTODO();
    } else {
        Logger.log('Language not supported: ' + language);
    }
}

function translateSheetTODO() {
    const language = PropertiesService.getDocumentProperties().getProperty('language') || 'English';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Upda with the corresponding styles
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
