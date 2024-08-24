/**
 * Logs all the document properties to the console.
 * @returns {void}
 * @customfunction
 */
function logAllDocumentProperties() {
    const docProperties = PropertiesService.getDocumentProperties();
    const allProperties = docProperties.getProperties();

    const expirationDateHash = {};
    const lastHashProperty = {};
    const timestampProperties = {};
    const configProperties = {};
    const otherProperties = {};

    const conditionsMap = [
        {
            condition: key => key === 'expirationDateHash',
            action: key => expirationDateHash[key] = allProperties[key]
        },
        {
            condition: key => key === 'lastHash',
            action: key => lastHashProperty[key] = allProperties[key]
        },
        {
            condition: key => key.toLowerCase().includes('h') && !isNaN(parseInt(key.substring(1))),
            action: key => timestampProperties[key] = allProperties[key]
        },
        {
            condition: key => key.toLowerCase().includes('last') || key.toLowerCase().includes('enable') || key.toLowerCase().includes('menus'),
            action: key => configProperties[key] = allProperties[key]
        }
    ];

    for (const key in allProperties) {
        let matched = false;
        for (const { condition, action } of conditionsMap) {
            if (condition(key)) {
                action(key);
                matched = true;
                break;
            }
        }
        if (!matched) {
            otherProperties[key] = allProperties[key];
        }
    }

    Logger.log('Document Properties:');

    Logger.log('EXPIRATION DATE HASH:');
    for (const key in expirationDateHash) {
        Logger.log(`${key}: ${expirationDateHash[key]}`);
    }

    Logger.log('LAST HASH:');
    for (const key in lastHashProperty) {
        Logger.log(`${key}: ${lastHashProperty[key]}`);
    }

    Logger.log('TIMESTAMP PROPERTIES for column H:');
    for (const key in timestampProperties) {
        Logger.log(`${key}: ${timestampProperties[key]}`);
    }

    Logger.log('CONFIG PROPERTIES:');
    for (const key in configProperties) {
        Logger.log(`${key}: ${configProperties[key]}`);
    }

    Logger.log('OTHER PROPERTIES:');
    for (const key in otherProperties) {
        Logger.log(`${key}: ${otherProperties[key]}`);
    }
}

/**
 * Removes unused properties from the document properties.
 * @returns {void}
 * @customfunction
 */
function removeUnusedProperties() {
    const docProperties = PropertiesService.getDocumentProperties();
    const allProperties = docProperties.getProperties();

    const unusedKeys = [
        // add here the keys that are not used as strings:
    ];
    let removedKeys = [];

    for (const key of unusedKeys) {
        if (key in allProperties) {
            docProperties.deleteProperty(key);
            removedKeys.push(key);
        }
    }

    if (removedKeys.length > 0) {
        Logger.log('Removed the following unused properties:');
        for (const key of removedKeys) {
            Logger.log(key);
        }
    } else {
        Logger.log('No unused properties found to remove.');
    }
}


if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        logAllDocumentProperties,
        removeUnusedProperties
    };
}