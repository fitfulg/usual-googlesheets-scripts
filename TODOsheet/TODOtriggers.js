function onEdit(e) {
    const range = e.range;
    const value = e.value;

    // Check if the cell is in the specified columns and contains example text
    const columnLetter = range.getA1Notation().charAt(0);

    if (exampleTexts[columnLetter] && value === exampleTexts[columnLetter].text) {
        // Remove the example text formatting
        range.setFontStyle("normal")
            .setFontColor("#000000"); // Set font color to black or default
    }
}
