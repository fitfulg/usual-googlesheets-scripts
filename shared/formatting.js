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
