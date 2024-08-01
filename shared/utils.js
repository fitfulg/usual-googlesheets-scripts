

/**
 * Extracts URLs from a rich text value.
 *
 * @param {RichTextValue} richTextValue - The rich text value to extract URLs from.
 * @return {string[]} The extracted URLs.
 */
function extractUrls(richTextValue) {
    const urls = [];
    const text = richTextValue.getText();
    for (let i = 0; i < text.length; i++) {
        const url = richTextValue.getLinkUrl(i, i + 1);
        if (url) {
            urls.push(url);
        }
    }
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
    if (arr1.length !== arr2.length) return false;
    for (let i = 0; i < arr1.length; i++) {
        if (arr1[i] !== arr2[i]) return false;
    }
    return true;
}

/**
 * Generates a SHA-256 hash for the given content.
 *
 * @param {string} content - The content to hash.
 * @return {string} The generated hash in base64 encoding.
 */
function generateHash(content) {
    return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, content));
}

/**
 * Checks if the hash of the content of the sheet has changed.
 *
 * @param {string} lastHash - The previous hash value.
 * @param {string} currentHash - The current hash value.
 * @return {boolean} True if the hash has changed, false otherwise.
 */
function shouldRunUpdates(lastHash, currentHash) {
    return lastHash !== currentHash;
}

/**
 * Gets the content of the sheet and generates a hash for it.
 *
 * @return {string} The generated hash of the sheet content.
 */
function getSheetContentHash() {
    const range = getDataRange();
    const values = range.getValues().flat().join(",");
    return generateHash(values);
}

/**
 * Saves a snapshot of the current state of the active sheet.
 * The snapshot includes the text content and links of each cell.
 * 
 * @return {void}
 */
function saveSnapshot() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getDataRange();
    const richTextValues = range.getRichTextValues();
    const snapshot = {};

    for (let row = 0; row < richTextValues.length; row++) {
        for (let col = 0; col < richTextValues[row].length; col++) {
            const cellValue = richTextValues[row][col];
            if (cellValue) {
                const cellKey = `R${row + 1}C${col + 1}`;
                snapshot[cellKey] = {
                    text: cellValue.getText(),
                    links: []
                };

                for (let i = 0; i < cellValue.getText().length; i++) {
                    const url = cellValue.getLinkUrl(i, i + 1);
                    if (url) {
                        snapshot[cellKey].links.push({ start: i, end: i + 1, url });
                    }
                }
            }
        }
    }

    // Save snapshot to script properties
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('sheetSnapshot', JSON.stringify(snapshot));
    Logger.log("Snapshot saved.");
}



if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        extractUrls,
        arraysEqual,
        generateHash,
        shouldRunUpdates,
        getSheetContentHash,
        saveSnapshot
    };
}
