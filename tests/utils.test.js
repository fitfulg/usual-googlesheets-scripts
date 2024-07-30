const utils = require('../shared/utils');

describe('Utils Functions', () => {
    test('extractUrls function exists', () => {
        expect(typeof utils.extractUrls).toBe('function');
    });

    test('arraysEqual function exists', () => {
        expect(typeof utils.arraysEqual).toBe('function');
    });

    test('generateHash function exists', () => {
        expect(typeof utils.generateHash).toBe('function');
    });

    test('shouldRunUpdates function exists', () => {
        expect(typeof utils.shouldRunUpdates).toBe('function');
    });

    test('getSheetContentHash function exists', () => {
        expect(typeof utils.getSheetContentHash).toBe('function');
    });
});