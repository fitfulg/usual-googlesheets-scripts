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

        // RichTextValueBuilder to combine text and style
        const richTextValueBuilder = SpreadsheetApp.newRichTextValue()
            .setText(cellValue + ' ' + formattedDate);

        // Get current text styles
        const textStyle = SpreadsheetApp.newTextStyle()
            .setFontStyle('normal')
            .setForegroundColor('#000000')
            .build();

        richTextValueBuilder.setTextStyle(0, cellValue.length, textStyle);

        // Create styles for the date only
        const dateStyle = SpreadsheetApp.newTextStyle()
            .setFontStyle('italic')
            .setForegroundColor('#555555')
            .build();

        richTextValueBuilder.setTextStyle(cellValue.length, cellValue.length + formattedDate.length + 1, dateStyle);

        // Build the RichTextValue and set it in the cell
        const richTextValue = richTextValueBuilder.build();
        range.setRichTextValue(richTextValue);
    }
}
