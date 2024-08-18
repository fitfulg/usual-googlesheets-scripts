const TODOtriggers = require('../TODOsheet/TODOtriggers');

describe('TODOtriggers Functions', () => {
    test('onEdit function exists', () => {
        expect(typeof TODOtriggers.onEdit).toBe('function');
    });
});