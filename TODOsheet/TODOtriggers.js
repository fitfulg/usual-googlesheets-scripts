// Add the onEdit function to track changes in specified columns and add the date
function onEdit(e) {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const column = range.getColumn();
    const row = range.getRow();

    // Check if the edit is in columns C, D, E, F, G, H and from row 2 onwards
    if (column >= 3 && column <= 8 && row >= 2) {
        const cellValue = range.getValue();

        // If the cell is empty, reset text style and do nothing else
        if (cellValue.trim() === "") {
            resetTextStyle(range);
            return;
        }

        const date = new Date();
        const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");

        // Append or update the formatted date at the end of the cell content
        const dateFormatted = ` ${formattedDate}`;
        const datePattern = /\s\d{2}\/\d{2}\/\d{2}$/;

        // Check if there is already a date, and update it if so; otherwise, append it
        let richTextValue;
        if (datePattern.test(cellValue)) {
            richTextValue = updateDateWithStyle(cellValue, dateFormatted);
        } else {
            richTextValue = appendDateWithStyle(cellValue, dateFormatted);
        }

        // Set the value with the date and apply the rich text formatting
        range.setRichTextValue(richTextValue);
    }
}
