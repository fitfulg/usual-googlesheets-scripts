// Add the onEdit function to track changes in specified columns and add the date
function onEdit(e) {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const column = range.getColumn();
    const row = range.getRow();

    // Check if the edit is in columns C, D, E, F, G, H and from row 2 onwards
    if (column >= 3 && column <= 8 && row >= 2) {
        const cellValue = range.getValue();
        const date = new Date();
        const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");

        // Create a RichTextValueBuilder to combine text and style
        let richTextBuilder = SpreadsheetApp.newRichTextValue()
            .setText(cellValue + ' ' + formattedDate);

        // current text styles
        const textStyle = SpreadsheetApp.newTextStyle()
            .setFontStyle('normal')
            .setForegroundColor('#000000')
            .build();

        richTextBuilder = richTextBuilder.setTextStyle(0, cellValue.length, textStyle);

        // styles just for the date
        const dateStyle = SpreadsheetApp.newTextStyle()
            .setFontStyle('italic')
            .setForegroundColor('#555555')
            .build();

        richTextBuilder = richTextBuilder.setTextStyle(cellValue.length, cellValue.length + formattedDate.length + 1, dateStyle);

        // Set the RichTextValue in the cell
        range.setRichTextValue(richTextBuilder.build());
    }
}

