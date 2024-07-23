/* eslint-disable no-unused-vars */

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

function arraysEqual(arr1, arr2) {
    if (arr1.length !== arr2.length) return false;
    for (let i = 0; i < arr1.length; i++) {
        if (arr1[i] !== arr2[i]) return false;
    }
    return true;
}

function generateHash(content) {
    return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, content));
}
// check if the hash of the content of the sheet has changed
function shouldRunUpdates(lastHash, currentHash) {
    return lastHash !== currentHash;
}
//  get the content of the sheet and generate a hash for it
function getSheetContentHash() {
    const range = getDataRange();
    const values = range.getValues().flat().join(",");
    return generateHash(values);
}