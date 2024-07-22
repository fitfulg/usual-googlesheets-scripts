// tests/formatting.test.js

const { setCellStyle } = require('../shared/formatting');

describe('setCellStyle', () => {
    test('should set cell style correctly', () => {
        const mockRange = {
            setValue: jest.fn().mockReturnThis(),
            setFontWeight: jest.fn().mockReturnThis(),
            setFontColor: jest.fn().mockReturnThis(),
            setHorizontalAlignment: jest.fn().mockReturnThis(),
            setBackground: jest.fn().mockReturnThis(),
        };

        global.SpreadsheetApp = {
            getActiveSpreadsheet: jest.fn().mockReturnThis(),
            getActiveSheet: jest.fn().mockReturnThis(),
            newTextStyle: jest.fn().mockReturnThis(),
            newRichTextValue: jest.fn().mockReturnThis(),
        };

        global.sheet = {
            getRange: jest.fn().mockReturnValue(mockRange),
        };

        setCellStyle('A1', 'Test Value', 'bold', '#000000', '#ffffff', 'center');

        expect(mockRange.setValue).toHaveBeenCalledWith('Test Value');
        expect(mockRange.setFontWeight).toHaveBeenCalledWith('bold');
        expect(mockRange.setFontColor).toHaveBeenCalledWith('#000000');
        expect(mockRange.setHorizontalAlignment).toHaveBeenCalledWith('center');
        expect(mockRange.setBackground).toHaveBeenCalledWith('#ffffff');
    });
});
