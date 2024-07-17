// Add the onEdit function to track changes in specified columns and add the date
function onEdit(e) {
    const range = e.range;
    const column = range.getColumn();
    const row = range.getRow();

    // Check if the edit is in columns C, D, E, F, G, H and from row 2 onwards
    if (column >= 3 && column <= 8 && row >= 2) {
        const cellValue = range.getValue();

        if (cellValue.trim() === "") return resetTextStyle(range);

        const date = new Date();
        const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");

        // Append or update the formatted date at the end of the cell content
        const dateFormatted = `\n${formattedDate}`;

        const richTextValue = datePattern.test(cellValue)
            ? updateDateWithStyle(cellValue, dateFormatted)
            : appendDateWithStyle(cellValue, dateFormatted);

        // Set the value with the date and apply the rich text formatting
        range.setRichTextValue(richTextValue);
    }
}
