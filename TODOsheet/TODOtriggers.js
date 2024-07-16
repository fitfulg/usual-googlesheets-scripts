// Add the onEdit function to track changes in specified columns and add the date
function onEdit(e) {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const column = range.getColumn();
    const row = range.getRow();

    // Check if the edition is in columns C, D, E, F, G, H and from row 2 down
    if (column >= 3 && column <= 8 && row >= 2) {
        const cellValue = range.getValue();
        const date = new Date();
        const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");

        // Formatear la fecha en cursiva y gris oscuro
        const dateFormatted = ` ${formattedDate}`;
        range.setValue(cellValue + dateFormatted)
            .setFontStyle('italic')
            .setFontColor('#555555');
    }
}
