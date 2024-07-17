// Higher-order function to apply formatting to a range only if it is valid
const withValidRange = (fn) => (range, ...args) => range && fn(range, ...args);

const Format = withValidRange((range) => {
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
});

const applyBordersWithStyle = withValidRange((range, borderStyle) => range.setBorder(true, true, true, true, true, true, "#000000", borderStyle));
const applyBorders = range => applyBordersWithStyle(range, SpreadsheetApp.BorderStyle.SOLID);
const applyThickBorders = range => applyBordersWithStyle(range, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

function applyFormatToSelected() {
    let range = sheet.getActiveRange();
    Format(range);
    applyBorders(range);
}

function applyFormatToAll() {
    let range = dataRange;
    Format(range);
    applyBorders(range);
}

function setCellStyle(cell, value, fontWeight, fontColor, backgroundColor, alignment) {
    let range = sheet.getRange(cell);
    range.setValue(value)
        .setFontWeight(fontWeight)
        .setFontColor(fontColor)
        .setHorizontalAlignment(alignment);

    if (backgroundColor) {
        range.setBackground(backgroundColor);
    }
}

// append DATE to cell
function appendDateWithStyle(cellValue, dateFormatted) {
    const newText = cellValue + '\n' + dateFormatted;
    return createRichTextValue(newText, dateFormatted);
}

// Update DATE in cell if it already exists
function updateDateWithStyle(cellValue, dateFormatted) {
    const newText = cellValue.replace(datePattern, '\n' + dateFormatted);
    return createRichTextValue(newText, dateFormatted);
}

// create rich text value with italic date
function createRichTextValue(text, dateFormatted) {
    return SpreadsheetApp.newRichTextValue()
        .setText(text)
        .setTextStyle(text.length - dateFormatted.length, text.length, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor('#A9A9A9').build())
        .build();
}

// Reset the text style of a cell
function resetTextStyle(range) {
    const richTextValue = SpreadsheetApp.newRichTextValue()
        .setText(range.getValue())
        .setTextStyle(SpreadsheetApp.newTextStyle().build())
        .build();

    range.setRichTextValue(richTextValue);
}
