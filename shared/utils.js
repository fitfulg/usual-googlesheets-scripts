

/**
 * Extracts URLs from a rich text value.
 *
 * @param {RichTextValue} richTextValue - The rich text value to extract URLs from.
 * @return {Array} The extracted URLs with their start and end positions.
 */
function extractUrls(richTextValue) {
    Logger.log('extractUrls triggered');
    const urls = [];
    const text = richTextValue.getText();
    for (let i = 0; i < text.length; i++) {
        const url = richTextValue.getLinkUrl(i, i + 1);
        if (url) {
            urls.push({ url, start: i, end: i + 1 });
        }
    }
    Logger.log(`returning urls: ${urls}`);
    return urls;
}

/**
 * Checks if two arrays are equal.
 *
 * @param {Array} arr1 - The first array.
 * @param {Array} arr2 - The second array.
 * @return {boolean} True if the arrays are equal, false otherwise.
 */
function arraysEqual(arr1, arr2) {
    Logger.log('arraysEqual triggered');
    if (arr1.length !== arr2.length) return false;
    for (let i = 0; i < arr1.length; i++) {
        if (arr1[i] !== arr2[i]) return false;
    }
    Logger.log('arrays are equal');
    return true;
}

/**
 * Generates a SHA-256 hash for the given content.
 *
 * @param {string} content - The content to hash.
 * @return {string} The generated hash in base64 encoding.
 */
function generateHash(content) {
    Logger.log('generateHash triggered');
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, content));
    Logger.log('hash generated');
    return hash;
}

/**
 * Checks if the hash of the content of the sheet has changed.
 *
 * @param {string} lastHash - The previous hash value.
 * @param {string} currentHash - The current hash value.
 * @return {boolean} True if the hash has changed, false otherwise.
 */
function shouldRunUpdates(lastHash, currentHash) {
    Logger.log('shouldRunUpdates triggered');
    const hasChanged = lastHash !== currentHash;
    Logger.log(`hash has changed: ${hasChanged}`);
    return hasChanged;
}

/**
 * Gets the content of the sheet and generates a hash for it.
 *
 * @return {string} The generated hash of the sheet content.
 */
function getSheetContentHash() {
    Logger.log('getSheetContentHash triggered');
    const range = getDataRange();
    const values = range.getValues().flat().join(",");
    Logger.log('getSheetContentHash: returning generateHash');
    const hash = generateHash(values);
    Logger.log(`generated hash: ${hash}`);
    return hash;
}

/**
 * Saves a snapshot of the current state of the active sheet.
 * The snapshot includes the text content and links of each cell.
 * You can specify cells to ignore by passing an array of cell references.
 * 
 * @param {Array<string>} cellsToIgnore - (e.g., ["R1C3", "R1C4", "R1C5"] for C1, D1, E1).
 * @return {object} The snapshot object.
 */
function saveSnapshot(cellsToIgnore = []) {
    Logger.log('saveSnapshot triggered');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getDataRange();
    const richTextValues = range.getRichTextValues();
    const snapshot = {};

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellKey = `R${row + 1}C${col + 1}`;
            if (cellsToIgnore.includes(cellKey)) {
                Logger.log(`Ignoring cell ${cellKey} from snapshot.`);
                continue;
            }

            const cellValue = richTextValues[row][col];
            if (cellValue) {
                const urls = extractUrls(cellValue);
                snapshot[cellKey] = {
                    text: cellValue.getText(),
                    links: urls
                };
                Logger.log(`Snapshot saved for cell ${cellKey}.`);
            }
        }
    }

    // Save snapshot to script properties
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('sheetSnapshot', JSON.stringify(snapshot));
    Logger.log("Snapshot saved.");
    return snapshot;
}

/**
 * Restores the sheet to a previously saved snapshot state.
 * This includes restoring text content, links, and optional custom formatting.
 *
 * @param {function} formatCallback - Optional callback function to apply custom formatting.
 * @return {void}
 */
function restoreSnapshot(formatCallback) {
    Logger.log('restoreSnapshot triggered');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getDataRange();
    const properties = PropertiesService.getScriptProperties();
    const snapshotJson = properties.getProperty('sheetSnapshot');

    if (!snapshotJson) {
        Logger.log("No snapshot found.");
        return;
    }

    const snapshot = JSON.parse(snapshotJson);
    const richTextValues = range.getRichTextValues();

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellKey = `R${row + 1}C${col + 1}`;
            if (snapshot[cellKey]) {
                const cellData = snapshot[cellKey];
                const builder = SpreadsheetApp.newRichTextValue()
                    .setText(cellData.text);
                Logger.log(`Restoring snapshot for cell ${cellKey}.`);
                // Restore links
                for (const link of cellData.links) {
                    Logger.log(`Restoring link: ${link.url} at ${link.start}-${link.end}`);
                    builder.setLinkUrl(link.start, link.end, link.url);
                }
                Logger.log(`Restored links for cell ${cellKey}. With a total of ${cellData.links.length}`);
                // Apply custom formatting if a callback is provided
                if (formatCallback) {
                    Logger.log(`restoreSnapshot()/formatCallback(): Applying custom formatting for cell ${cellKey}.`);
                    formatCallback(builder, cellData.text);
                }
                richTextValues[row][col] = builder.build();
            }
        }
    }

    range.setRichTextValues(richTextValues);
    Logger.log("Snapshot restored.");
}

/**
 * Iterates over each cell in the selected range and applies the specified function to it.
 * @param {function(Range, RichTextValue): void} cellFunction - The function to apply to each cell.
 */
function processCells(cellFunction) {
    Logger.log('processCells triggered');
    const range = sheet.getActiveRange();
    const richTextValues = range.getRichTextValues();

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellValue = richTextValues[row][col];
            if (cellValue) {
                const cellRange = range.getCell(row + 1, col + 1);
                cellFunction(cellRange, cellValue);
            }
        }
    }
    Logger.log('processCells completed');
}

/**
 * Preserves existing text styles and links from the original text to the new text.
 * @param {RichTextValue} originalTextValue - The original rich text value.
 * @param {RichTextValueBuilder} newTextValueBuilder - The rich text value builder for the new text.
 * @param {number} offset - The offset to apply to the new text positions.
 */
function preserveStylesAndLinks(originalTextValue, newTextValueBuilder, offset) {
    Logger.log('preserveStylesAndLinks triggered');
    const originalText = originalTextValue.getText();
    const newText = newTextValueBuilder.build().getText();
    const minLength = Math.min(originalText.length, newText.length - offset);

    for (let i = 0; i < minLength; i++) {
        const style = originalTextValue.getTextStyle(i, i + 1);
        newTextValueBuilder.setTextStyle(i + offset, i + offset + 1, style);

        const url = originalTextValue.getLinkUrl(i, i + 1);
        if (url) {
            newTextValueBuilder.setLinkUrl(i + offset, i + offset + 1, url);
        }
    }
    Logger.log('preserveStylesAndLinks completed');
}


if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        extractUrls,
        arraysEqual,
        generateHash,
        shouldRunUpdates,
        getSheetContentHash,
        saveSnapshot,
        restoreSnapshot,
        processCells,
        preserveStylesAndLinks
    };
}
