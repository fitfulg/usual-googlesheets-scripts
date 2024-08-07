// globals.js: ui

/**
 * Gets the current language from document properties or returns 'English' as default.
 * @returns {string} The current language
 */
const getCurrentLanguageTODO = () => PropertiesService.getDocumentProperties().getProperty('language') || 'English';

/**
 * Creates custom menus in the spreadsheet.
 * Adds menu items to the UI and assigns functions to them.
 * 
 * @customfunction
 */
function createMenusTODO() {
    Logger.log('createMenusTODO triggered');
    const currentLanguage = getCurrentLanguageTODO();

    const functionNameMap = {
        'restoreDefaultTodoTemplate': 'applyFormatToAllTODO',
        'restoreCellBackgroundColors': 'customCellBGColorTODO',
        'addCheckboxesToSelectedCells': 'addCheckboxesTODO',
        'markCheckboxInSelectedCells': 'markCheckboxTODO',
        'markAllCheckboxesInSelectedCells': 'markAllCheckboxesTODO',
        'restoreCheckboxes': 'restoreCheckboxesTODO',
        'removeAllCheckboxesInSelectedCells': 'removeCheckboxesTODO',
        'applyFormat': 'applyFormatToSelected',
        'applyFormatToAll': 'applyFormatToAll',
        'createPieChart': 'createPieChartTODO',
        'deletePieCharts': 'deleteAllChartsTODO',
        'logHelloWorld': 'logHelloWorld',
        'versionAndFeatureDetails': 'updateCellCommentTODO',
        'saveSnapshot': 'saveSnapshotTODO',
        'restoreSnapshot': 'restoreSnapshotTODO'
    };

    for (const { config, items } of menus) {
        const menuTitle = config[0].title[currentLanguage];
        const menu = ui.createMenu(menuTitle);

        for (const { key, separatorAfter } of items) {
            const itemTitle = config[0].items[key][currentLanguage];
            const functionName = functionNameMap[key] || key;

            menu.addItem(itemTitle, functionName);
            if (separatorAfter) {
                menu.addSeparator();
            }
        }

        menu.addToUi();
    }
}

/**
 * Displays a "Hello World" message in an alert.
 *
 * @customfunction
 */
function logHelloWorld() {
    ui.alert('Hello World!!');
    Logger.log('hello world test');
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        createMenusTODO,
        getCurrentLanguageTODO,
        logHelloWorld
    }
}