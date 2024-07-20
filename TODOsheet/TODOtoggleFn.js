// TODOsheet/TODOtoggleFn.js: createPieChartTODO, deleteAllChartsTODO

function togglePieChartTODO(action) {
    Logger.log(`togglePieChartTODO called with action: ${action}`);
    if (action === 'Hide Piechart') {
        Logger.log('Attempting to hide piechart');
        deleteAllChartsTODO();
        isPieChartVisible = false;
        Logger.log('Piechart hidden');
    } else if (action === 'Show Piechart') {
        Logger.log('Attempting to show piechart');
        createPieChartTODO();
        isPieChartVisible = true;
        Logger.log('Piechart shown');
    } else {
        Logger.log('Invalid action selected');
    }
}

function toggleDatesTODO(action) {
    Logger.log(`toggleDatesTODO called with action: ${action}`);
    if (action === 'Hide Dates' && areDatesVisible) {
        Logger.log('Hiding dates');
        hideDatesTODO();
        areDatesVisible = false;
        Logger.log('Dates hidden');
    } else if (action === 'Show Dates' && !areDatesVisible) {
        Logger.log('Showing dates');
        showDatesTODO();
        areDatesVisible = true;
        Logger.log('Dates shown');
    } else {
        Logger.log('No action taken');
    }
}


