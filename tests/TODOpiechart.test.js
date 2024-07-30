const TODOpiechart = require('../TODOsheet/TODOpiechart');

describe('TODOpiechart Functions', () => {
    test('createPieChartTODO function exists', () => {
        expect(typeof TODOpiechart.createPieChartTODO).toBe('function');
    });

    test('deleteAllChartsTODO function exists', () => {
        expect(typeof TODOpiechart.deleteAllChartsTODO).toBe('function');
    });
});