// function onEdit(e) {
//     const sheet = e.source.getActiveSheet();
//     const range = e.range;
//     const columnLetter = range.getA1Notation().charAt(0);

//     if (exampleTexts[columnLetter]) {
//         const { text, color } = exampleTexts[columnLetter];

//         // If the cell contains the example text, clear it
//         if (range.getValue() === text) {
//             range.setValue("")
//                 .setFontStyle("normal")
//                 .setFontColor("#000000"); // Set font color to black or default
//         }
//     }
// }
