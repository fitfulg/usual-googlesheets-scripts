const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

function withValidRange(fn) {
    return function (range, ...args) {
        if (range) {
            fn(range, ...args);
        }
    };
}

const Format = withValidRange(function (range) {
    // Apply the desired formats
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
});

const applyBordersWithStyle = withValidRange(function (range, borderStyle) {
    // Apply borders with specified style
    range.setBorder(true, true, true, true, true, true, "#000000", borderStyle);
});

const applyBorders = range => applyBordersWithStyle(range, SpreadsheetApp.BorderStyle.SOLID);
const applyThickBorders = range => applyBordersWithStyle(range, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

function applyFormatToSelected() {
    // Get the active sheet and the selected range
    let range = sheet.getActiveRange();
    // Apply formatting to the selected range
    Format(range);
    applyBorders(range);
}

function applyFormatToAll() {
    // Get the active sheet and the entire data range
    let range = sheet.getDataRange();
    // Apply formatting to the entire data range
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

// Append DATE in italic and dark gray without changing the original cell text
function appendDateWithStyle(cellValue, dateFormatted) {
    const richTextValue = SpreadsheetApp.newRichTextValue()
        .setText(cellValue + dateFormatted)
        .setTextStyle(cellValue.length, (cellValue + dateFormatted).length, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor('#A9A9A9').build())
        .build();

    return richTextValue;
}

// Function to reset the text style of a cell
function resetTextStyle(range) {
    const richTextValue = SpreadsheetApp.newRichTextValue()
        .setText(range.getValue())
        .setTextStyle(SpreadsheetApp.newTextStyle().build())
        .build();

    range.setRichTextValue(richTextValue);
}
