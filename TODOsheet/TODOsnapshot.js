/* eslint-disable no-unused-vars */
// shared/util.js: saveSnapshot, restoreSnapshot

/**
 * Saves a snapshot of the current state of the active sheet while ignoring specific cells.
 * Ignores cells C1, D1, and E1 so we retain the changed column titles when cell max limit is reached.
 * 
 * @return {void}
 */
function saveSnapshotTODO() {
    Logger.log('saveSnapshotTODO triggered');
    const cellsToIgnore = ["R1C1", "R1C2", "R1C3", "R1C4", "R1C5", "R1C6", "R1C7", "R1C8"]
    Logger.log(`Ignoring cells ${cellsToIgnore.join(', ')} from snapshot.`);
    const snapshot = saveSnapshot(cellsToIgnore);

    // Save filtered snapshot to script properties
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('sheetSnapshot', JSON.stringify(snapshot));
    Logger.log("Snapshot saved, excluding specified cells.");
}


/**
 * Restores the sheet snapshot and applies custom formatting for dates and "days left".
 *
 * @return {void}
 */
function restoreSnapshotTODO() {
    Logger.log('restoreSnapshotTODO triggered');
    restoreSnapshot((builder, text) => {
        // Reapply formatting for dates and "days left"
        const dateMatches = text.match(/\d{2}\/\d{2}\/\d{2}/g);
        const daysLeftPattern = /\((\d+)\) days left/;
        const daysLeftMatch = text.match(daysLeftPattern);

        if (dateMatches) {
            Logger.log('restoreSnapshotTODO)(): dateMatches :', dateMatches);
            for (const date of dateMatches) {
                const start = text.lastIndexOf(date);
                const end = start + date.length;
                builder.setTextStyle(start, end, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor('#A9A9A9').build());
                Logger.log('restoreSnapshotTODO() date to be formatted :', date);
            }
        }

        if (daysLeftMatch) {
            Logger.log('restoreSnapshotTODO() daysLeftMatch :', daysLeftMatch);
            const start = text.lastIndexOf(daysLeftMatch[0]);
            const end = start + daysLeftMatch[0].length;
            builder.setTextStyle(start, end, SpreadsheetApp.newTextStyle().setItalic(true).setForegroundColor('#FF0000').build());
            Logger.log('restoreSnapshotTODO() days left to be formatted :', daysLeftMatch[0]);
        }
    });
}

// for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        saveSnapshotTODO,
        restoreSnapshotTODO
    }
}