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
    const ui = SpreadsheetApp.getUi();
    const currentLanguage = getCurrentLanguageTODO();

    const functionNameMap = {
        'restoreDefaultTodoTemplate': 'applyFormatToAllTODO',
        'restoreCellBackgroundColors': 'customCeilBGColorTODO',
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

    menus.forEach(menuConfig => {
        const menuTitle = menuConfig.config[0].title[currentLanguage];
        let menu = ui.createMenu(menuTitle);

        menuConfig.items.forEach(item => {
            const itemTitle = menuConfig.config[0].items[item.key][currentLanguage];
            let functionName = functionNameMap[item.key] || item.key;

            menu.addItem(itemTitle, functionName);
            if (item.separatorAfter) {
                menu.addSeparator();
            }
        });

        menu.addToUi();
    });
}

/**
 * Displays a "Hello World" message in an alert.
 *
 * @customfunction
 */
function logHelloWorld() {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Hello World!!');
    Logger.log('hello world test');
}
