const TODOtoggleFn = require('../TODOsheet/TODOtoggleFn');

describe('TODOtoggleFn Functions', () => {
    test('togglePieChartTODO function exists', () => {
        expect(typeof TODOtoggleFn.togglePieChartTODO).toBe('function');
    });

    test('handlePieChartToggleTODO function exists', () => {
        expect(typeof TODOtoggleFn.handlePieChartToggleTODO).toBe('function');
    });
});