const Menu = require('../Menu');

describe('Menu Functions', () => {
    test('onOpen function exists', () => {
        expect(typeof Menu.onOpen).toBe('function');
    });

    test('runAllFunctionsTODO function exists', () => {
        expect(typeof Menu.runAllFunctionsTODO).toBe('function');
    });
});